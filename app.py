#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WordFormat —— Word 论文/公文格式一键排版工具

用法（命令行）:
    python app.py 文件1.docx 文件2.docx            # 另存为 "原名_已排版.docx"
    python app.py 文件夹 -o 输出文件夹              # 批量处理整个文件夹
    python app.py a.docx --inplace --backup        # 原地覆盖，并先备份
    python app.py a.docx --no-page-number          # 不自动补页码
    python app.py a.docx --plain-title             # 首段没用标题样式时，也尝试识别为主标题

用法（图形界面）:
    python app.py            # 不带任何参数时启动 GUI
    python app.py --gui

打包 exe（Windows）:
    pip install pyinstaller
    pyinstaller app.spec        # 或双击 build.bat
"""

from __future__ import annotations

import argparse
import os
import queue
import shutil
import subprocess
import sys
import threading
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from formatter import (format_document, preview_profile,  # noqa: E402
                       resolve_profile)
from profile import (FormatProfile, ProfileError,  # noqa: E402
                     builtin_profiles, write_default_template)

CONFIG_DIR_NAME = "profiles"

APP_NAME = "WordFormat"
SUFFIX = "_已排版"


def _log_dir():
    """崩溃日志目录：打包成 exe 时放在 exe 旁边，否则放脚本目录。"""
    import tempfile
    try:
        base = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
            else os.path.dirname(os.path.abspath(__file__))
        with open(os.path.join(base, ".wtest"), "w") as f:
            f.write("1")
        os.remove(os.path.join(base, ".wtest"))
    except Exception:
        base = tempfile.gettempdir()
    return base


def write_crash_log(exc_text):
    """把崩溃信息写到 crash.log，并尽量弹窗提示（GUI 模式下否则会直接闪退）。"""
    path = os.path.join(_log_dir(), "crash.log")
    try:
        import datetime
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"\n===== {datetime.datetime.now():%Y-%m-%d %H:%M:%S} =====\n")
            f.write(exc_text + "\n")
    except Exception:
        path = ""
    return path


def install_crash_handler():
    """全局兜底：任何未捕获异常都写 crash.log（对无控制台的 exe 尤其重要）。"""
    import datetime

    def hook(exc_type, exc, tb):
        text = "".join(traceback.format_exception(exc_type, exc, tb))
        path = write_crash_log(text)
        try:
            import tkinter.messagebox as mb
            mb.showerror(APP_NAME, f"程序出错，详情已写入：\n{path}\n\n{exc}")
        except Exception:
            pass
        return sys.__excepthook__(exc_type, exc, tb)

    sys.excepthook = hook


def selftest(log=print):
    """环境自检：输出 Python / python-docx / tkinter / 打包状态，便于排查。"""
    import platform
    log(f"{APP_NAME} 环境自检")
    log(f"  Python      : {platform.python_version()}  ({sys.executable})")
    log(f"  运行方式    : {'打包后的 exe' if getattr(sys, 'frozen', False) else '源码'}")
    try:
        import docx
        log(f"  python-docx : OK ({getattr(docx, '__version__', 'unknown')})")
    except Exception as e:
        log(f"  python-docx : 缺失 -> {e}")
    try:
        import tkinter
        log(f"  tkinter     : OK (GUI 可用)")
    except Exception as e:
        log(f"  tkinter     : 不可用 -> {e}（只能用命令行模式）")
    log(f"  日志目录    : {_log_dir()}")
    return 0


# --------------------------------------------------------------------------- #
# .doc 转换（Windows 用 Word COM，其它系统退化为 LibreOffice）
# --------------------------------------------------------------------------- #

def convert_doc_to_docx(path, log=print):
    """.doc -> .docx，无法转换时抛出 RuntimeError。"""
    out_dir = os.path.dirname(os.path.abspath(path)) or "."
    lower = path.lower()

    if lower.endswith(".docx"):
        return path

    if not lower.endswith(".doc"):
        raise RuntimeError(f"不支持的文件类型：{path}")

    # 1) LibreOffice / WPS 命令行
    for exe in ("soffice", "libreoffice", "wps", "et"):
        bin_path = shutil.which(exe)
        if not bin_path:
            continue
        try:
            subprocess.run([bin_path, "--headless", "--convert-to", "docx",
                            "--outdir", out_dir, path],
                           check=True, timeout=120,
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            candidate = os.path.join(
                out_dir, os.path.splitext(os.path.basename(path))[0] + ".docx")
            if os.path.exists(candidate):
                log(f"  已转换 .doc -> {os.path.basename(candidate)}")
                return candidate
        except Exception:
            continue

    # 2) Windows + Microsoft Word
    if sys.platform.startswith("win"):
        try:
            import win32com.client  # type: ignore
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(os.path.abspath(path))
                new_path = os.path.splitext(os.path.abspath(path))[0] + ".docx"
                doc.SaveAs(new_path, FileFormat=16)  # 16 = docx
                doc.Close()
                log(f"  已用 Word 转换 -> {os.path.basename(new_path)}")
                return new_path
            finally:
                word.Quit()
        except Exception as e:
            raise RuntimeError(f"转换 .doc 失败（{e}），请先用 Word 另存为 .docx")

    raise RuntimeError("未能自动转换 .doc，请先另存为 .docx")


# --------------------------------------------------------------------------- #
# 处理流程
# --------------------------------------------------------------------------- #

def build_output_path(src, mode, outdir):
    base = os.path.basename(src)
    name, ext = os.path.splitext(base)
    if mode == "inplace":
        return src
    if mode == "folder":
        return os.path.join(outdir, name + SUFFIX + ext)
    return os.path.join(os.path.dirname(os.path.abspath(src)), name + SUFFIX + ext)


def process_one(src, mode="suffix", outdir="", backup=False,
                add_page_number=None, plain_title=None, profile=None, log=print):
    log(f"▶ {os.path.basename(src)}")
    try:
        src_docx = convert_doc_to_docx(src, log)
    except Exception as e:
        log(f"  [跳过] {e}")
        return False, str(e)

    if backup:
        bak = os.path.splitext(src)[0] + ".bak" + os.path.splitext(src)[1]
        try:
            shutil.copy2(src, bak)
        except Exception as e:
            log(f"  [警告] 备份失败：{e}")

    dst = build_output_path(src_docx, mode, outdir)
    if dst == src_docx and backup is False:
        pass
    if outdir and mode == "folder":
        os.makedirs(outdir, exist_ok=True)

    try:
        stats = format_document(src_docx, dst,
                                add_page_number=add_page_number,
                                detect_plain_title=plain_title,
                                profile=profile,
                                log=lambda m: log("  " + m))
    except Exception as e:
        log(f"  [失败] {e}")
        log(traceback.format_exc(limit=2))
        return False, str(e)

    log(f"  ✔ 已保存：{dst}")
    return True, stats


def collect_files(inputs, recursive=False):
    files = []
    for item in inputs:
        if os.path.isdir(item):
            if recursive:
                for root, _dirs, names in os.walk(item):
                    for n in names:
                        if n.lower().endswith((".doc", ".docx")) and not n.startswith("~$"):
                            files.append(os.path.join(root, n))
            else:
                for n in sorted(os.listdir(item)):
                    if n.lower().endswith((".doc", ".docx")) and not n.startswith("~$"):
                        files.append(os.path.join(item, n))
        elif os.path.isfile(item):
            files.append(item)
    # 去重保序
    seen, out = set(), []
    for f in files:
        k = os.path.abspath(f)
        if k not in seen:
            seen.add(k)
            out.append(f)
    return out


def run_batch(files, mode, outdir, backup, add_page_number, plain_title,
              profile=None, log=print):
    ok = fail = 0
    for f in files:
        good, _ = process_one(f, mode, outdir, backup, add_page_number,
                              plain_title, profile, log)
        ok += good
        fail += (not good)
    log(f"\n完成：成功 {ok} 个，失败 {fail} 个")
    return ok, fail


def load_profile_quiet(path, log=print):
    """加载配置，失败时给出可读提示并返回 None。"""
    try:
        prof = FormatProfile.load(path)
        log(f"已载入配置：{prof.name}  （{path}）")
        return prof
    except ProfileError as e:
        log(f"[错误] 配置无法使用：{e}")
        return None
    except Exception as e:
        log(f"[错误] 读取配置失败：{e}")
        return None


# --------------------------------------------------------------------------- #
# 命令行
# --------------------------------------------------------------------------- #

def cli_main(argv=None):
    install_crash_handler()
    parser = argparse.ArgumentParser(
        prog=APP_NAME,
        description="Word 论文/公文格式一键排版（正文/标题/脚注/页码/图表题注）")
    parser.add_argument("inputs", nargs="*", help=".docx 文件或文件夹")
    parser.add_argument("-o", "--outdir", default="", help="输出文件夹")
    parser.add_argument("--inplace", action="store_true", help="直接覆盖原文件")
    parser.add_argument("--backup", action="store_true", help="处理前先备份原文件")
    parser.add_argument("--no-page-number", dest="page", action="store_false",
                        default=None, help="缺少页码时不自动添加")
    parser.add_argument("--page-number", dest="page", action="store_true",
                        default=None, help="缺少页码时自动添加")
    parser.add_argument("--plain-title", action="store_true",
                        help="首段未使用标题样式时，也尝试将其识别为主标题")
    parser.add_argument("-r", "--recursive", action="store_true",
                        help="递归处理子文件夹")
    parser.add_argument("--gui", action="store_true", help="启动图形界面")
    parser.add_argument("--selftest", action="store_true", help="环境自检")
    parser.add_argument("-c", "--config", default="",
                        help="排版方案 JSON 文件（导入自定义格式）")
    parser.add_argument("--export-config", metavar="PATH", default="",
                        help="导出当前方案为 JSON（默认方案或 -c 指定的方案）")
    parser.add_argument("--export-template", metavar="PATH", default="",
                        help="导出带说明的模板 JSON")
    parser.add_argument("--show-config", action="store_true",
                        help="打印当前方案摘要后退出")
    args = parser.parse_args(argv)

    if args.selftest:
        return selftest(print)

    profile = None
    if args.config:
        profile = load_profile_quiet(args.config)
        if profile is None:
            return 1

    if args.export_template:
        path = write_default_template(args.export_template)
        print(f"模板已导出：{path}")
        print("改完后用 -c 导入即可生效。")
        return 0

    if args.export_config:
        prof = profile or FormatProfile.default()
        prof.save(args.export_config)
        print(f"方案「{prof.name}」已导出：{args.export_config}")
        return 0

    if args.show_config:
        print(preview_profile(profile))
        return 0

    if args.gui or not args.inputs:
        try:
            import tkinter  # noqa: F401
        except ImportError:
            parser.print_help()
            print("\n[提示] 当前环境没有 tkinter，无法启动图形界面，请使用命令行参数。")
            print("        Windows 官方安装包默认带有 tkinter，重装时记得勾选 tcl/tk。")
            return 1
        try:
            return gui_main()
        except Exception as e:
            text = traceback.format_exc()
            path = write_crash_log(text)
            print(f"[错误] GUI 启动失败：{e}\n详情已写入：{path}\n{text}")
            return 1

    files = collect_files(args.inputs, args.recursive)
    if not files:
        print("没有找到任何 .doc/.docx 文件")
        return 1

    mode = "inplace" if args.inplace else ("folder" if args.outdir else "suffix")
    add_page = None if args.page is None else args.page
    ok, fail = run_batch(files, mode, args.outdir, args.backup,
                         add_page, None, profile, print)
    return 1 if fail and not ok else 0


# --------------------------------------------------------------------------- #
# 图形界面
# --------------------------------------------------------------------------- #

def gui_main():
    install_crash_handler()
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText

    root = tk.Tk()
    root.title(f"{APP_NAME} · Word 格式一键排版")
    root.geometry("760x600")
    root.minsize(680, 540)

    files = []
    q = queue.Queue()

    # ---------- 变量 ----------
    presets = builtin_profiles()
    preset_names = list(presets)
    current_profile = {"obj": FormatProfile.default()}

    var_mode = tk.StringVar(value="suffix")
    var_outdir = tk.StringVar(value="")
    var_backup = tk.BooleanVar(value=False)
    var_page = tk.BooleanVar(value=True)
    var_plain = tk.BooleanVar(value=False)
    var_quotes = tk.BooleanVar(value=True)
    var_ref_trim = tk.BooleanVar(value=True)

    # ---------- 布局 ----------
    top = ttk.Frame(root, padding=10)
    top.pack(fill="x")

    ttk.Label(top, text="待处理文件（可多选，也支持 .doc）",
              font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w")

    listbox = tk.Listbox(top, height=7, activestyle="none")
    listbox.pack(fill="x", pady=(4, 6))

    def refresh_list():
        listbox.delete(0, "end")
        for f in files:
            listbox.insert("end", f)

    def add_files():
        sel = filedialog.askopenfilenames(
            title="选择 Word 文件",
            filetypes=[("Word 文档", "*.docx *.doc"), ("所有文件", "*.*")])
        for s in sel:
            if s not in files:
                files.append(s)
        refresh_list()

    def add_folder():
        d = filedialog.askdirectory(title="选择文件夹")
        if not d:
            return
        for n in sorted(os.listdir(d)):
            if n.lower().endswith((".doc", ".docx")) and not n.startswith("~$"):
                p = os.path.join(d, n)
                if p not in files:
                    files.append(p)
        refresh_list()

    def clear_files():
        files.clear()
        refresh_list()

    btns = ttk.Frame(top)
    btns.pack(fill="x")
    ttk.Button(btns, text="添加文件", command=add_files).pack(side="left")
    ttk.Button(btns, text="添加文件夹", command=add_folder).pack(side="left", padx=6)
    ttk.Button(btns, text="清空", command=clear_files).pack(side="left")

    # ---------- 排版方案 ----------
    cfg = ttk.LabelFrame(root, text="排版方案（可导入 JSON 自定义）", padding=10)
    cfg.pack(fill="x", padx=10, pady=4)

    var_profile_name = tk.StringVar(value=current_profile["obj"].summary())

    row0 = ttk.Frame(cfg)
    row0.pack(fill="x")
    ttk.Label(row0, text="当前：", width=6).pack(side="left")
    ttk.Label(row0, textvariable=var_profile_name,
              foreground="#0a5").pack(side="left", fill="x", expand=True)

    def apply_profile(prof, origin=""):
        current_profile["obj"] = prof
        var_profile_name.set(prof.summary())
        suffix = f"（{origin}）" if origin else ""
        log(f"已切换方案：{prof.name}{suffix}")
        status.set(f"当前方案：{prof.name}")

    def choose_preset(event=None):
        name = cb_preset.get()
        if name in presets:
            apply_profile(presets[name].copy(), "内置预设")

    row1 = ttk.Frame(cfg)
    row1.pack(fill="x", pady=(6, 0))
    ttk.Label(row1, text="预设", width=6).pack(side="left")
    cb_preset = ttk.Combobox(row1, values=preset_names, state="readonly", width=34)
    cb_preset.set(preset_names[0])
    cb_preset.pack(side="left")
    cb_preset.bind("<<ComboboxSelected>>", choose_preset)

    row2 = ttk.Frame(cfg)
    row2.pack(fill="x", pady=(6, 0))

    def import_config():
        path = filedialog.askopenfilename(
            title="导入排版方案",
            filetypes=[("JSON 配置", "*.json"), ("所有文件", "*.*")])
        if not path:
            return
        prof = load_profile_quiet(path, log)
        if prof is None:
            messagebox.showerror(APP_NAME,
                                 f"配置无法使用，详情见下方日志。\n无法识别：{path}")
            return
        apply_profile(prof, os.path.basename(path))
        cb_preset.set("")
        messagebox.showinfo(APP_NAME, f"已导入方案「{prof.name}」")

    def export_config():
        prof = current_profile["obj"]
        path = filedialog.asksaveasfilename(
            title="导出当前方案", defaultextension=".json",
            initialfile=f"{prof.name}.json",
            filetypes=[("JSON 配置", "*.json")])
        if not path:
            return
        try:
            prof.save(path)
            log(f"方案已导出：{path}")
            messagebox.showinfo(APP_NAME, f"已导出到：\n{path}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"导出失败：{e}")

    def export_template():
        path = filedialog.asksaveasfilename(
            title="导出带说明的模板", defaultextension=".json",
            initialfile="排版方案模板.json",
            filetypes=[("JSON 配置", "*.json")])
        if not path:
            return
        try:
            write_default_template(path)
            log(f"模板已导出：{path}")
            messagebox.showinfo(APP_NAME,
                                f"模板已导出：\n{path}\n\n"
                                "改完保存后，用「导入配置」载入即可生效。")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"导出失败：{e}")

    def show_config():
        messagebox.showinfo(APP_NAME + " · 当前方案",
                            preview_profile(current_profile["obj"]))

    def reset_config():
        apply_profile(FormatProfile.default(), "已恢复默认")
        cb_preset.set(preset_names[0])

    ttk.Button(row2, text="导入配置…", command=import_config).pack(side="left")
    ttk.Button(row2, text="导出当前方案", command=export_config).pack(side="left", padx=4)
    ttk.Button(row2, text="导出模板（带说明）",
               command=export_template).pack(side="left")
    ttk.Button(row2, text="查看方案", command=show_config).pack(side="left", padx=4)
    ttk.Button(row2, text="恢复默认", command=reset_config).pack(side="left")

    opts = ttk.LabelFrame(root, text="输出与选项", padding=10)
    opts.pack(fill="x", padx=10, pady=4)

    ttk.Radiobutton(opts, text="另存为「原名_已排版.docx」",
                    variable=var_mode, value="suffix").grid(row=0, column=0, sticky="w")
    ttk.Radiobutton(opts, text="直接覆盖原文件",
                    variable=var_mode, value="inplace").grid(row=0, column=1, sticky="w")
    ttk.Radiobutton(opts, text="输出到文件夹",
                    variable=var_mode, value="folder").grid(row=1, column=0, sticky="w")
    ttk.Entry(opts, textvariable=var_outdir, width=42).grid(row=1, column=1, sticky="w")
    ttk.Button(opts, text="选择…", width=8,
               command=lambda: var_outdir.set(
                   filedialog.askdirectory(title="选择输出文件夹") or var_outdir.get())
               ).grid(row=1, column=2, sticky="w", padx=4)

    ttk.Checkbutton(opts, text="缺少页码时自动添加页码（底部居中·小五·Times New Roman）",
                    variable=var_page).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))
    ttk.Checkbutton(opts, text="首段没有标题样式时，也尝试识别为文章主标题",
                    variable=var_plain).grid(row=3, column=0, columnspan=3, sticky="w")
    ttk.Checkbutton(opts, text="处理前备份原文件（*.bak）",
                    variable=var_backup).grid(row=4, column=0, columnspan=3, sticky="w")
    ttk.Checkbutton(opts,
                    text='西文引文用英文弯引号（“Concept” / \"Concept\" -> “Concept”，中文引文不受影响）',
                    variable=var_quotes).grid(row=5, column=0, columnspan=3, sticky="w")
    ttk.Checkbutton(opts,
                    text="脚注序号小四上标，且与正文之间不留空格",
                    variable=var_ref_trim).grid(row=6, column=0, columnspan=3, sticky="w")

    logbox = ScrolledText(root, height=14, font=("Consolas", 10))
    logbox.pack(fill="both", expand=True, padx=10, pady=6)

    def log(msg=""):
        q.put(str(msg))

    def pump():
        try:
            while True:
                msg = q.get_nowait()
                logbox.insert("end", msg + "\n")
                logbox.see("end")
        except queue.Empty:
            pass
        root.after(120, pump)

    pump()

    progress = ttk.Progressbar(root, mode="indeterminate")
    progress.pack(fill="x", padx=10)

    status = tk.StringVar(value=f"就绪 · 支持 .docx（.doc 会尝试自动转换）")
    ttk.Label(root, textvariable=status, anchor="w").pack(fill="x", padx=10, pady=(2, 8))

    def start():
        if not files:
            messagebox.showwarning(APP_NAME, "请先添加要处理的文件")
            return
        mode = var_mode.get()
        outdir = var_outdir.get().strip()
        if mode == "folder" and not outdir:
            messagebox.showwarning(APP_NAME, "请选择输出文件夹")
            return
        if mode == "inplace" and not var_backup.get():
            if not messagebox.askyesno(
                    APP_NAME,
                    "将直接覆盖原文件，且未勾选备份。确定继续吗？\n（建议先勾选「处理前备份原文件」）"):
                return
        prof = current_profile["obj"].copy()
        prof.quotes.normalize_western = var_quotes.get()
        prof.footnote.reference_trim_spaces = var_ref_trim.get()
        prof.footnote.reference_superscript = var_ref_trim.get()

        btn_start.config(state="disabled")
        progress.start(12)
        status.set("正在处理…")

        def work():
            try:
                ok, fail = run_batch(
                    list(files), mode, outdir, var_backup.get(),
                    var_page.get(), var_plain.get(),
                    prof, log)
                q.put(f"—— 全部完成，成功 {ok} 个，失败 {fail} 个 ——")
                status.set(f"完成 · 成功 {ok} 个，失败 {fail} 个")
            except Exception as e:
                q.put(f"[错误] {e}\n{traceback.format_exc(limit=3)}")
                status.set("出错：" + str(e))
            finally:
                root.after(0, _finish)

        def _finish():
            progress.stop()
            btn_start.config(state="normal")

        threading.Thread(target=work, daemon=True).start()

    btn_start = ttk.Button(root, text="开始排版", command=start)
    btn_start.pack(pady=(0, 10))

    log(f"当前方案：{current_profile['obj'].summary()}")
    log("提示：可用「导入配置」载入自己的 JSON 方案。")
    tip = ("格式规则：正文/标题段前后 0 行距 1.5 倍｜正文小四宋体+Times New Roman 首行缩进 2 字符｜"
           "脚注小五单倍行距不缩进｜页码底部居中｜图表题注五号居中｜全文字色黑色")
    ttk.Label(root, text=tip, wraplength=720, foreground="#666").pack(padx=10, pady=(0, 8))

    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(cli_main())
