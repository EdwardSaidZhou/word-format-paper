# -*- coding: utf-8 -*-
"""
profile.py —— 排版方案（FormatProfile）

把 formatter.py 里所有排版参数抽成一份可序列化的配置，
支持导出为 JSON、从 JSON 导入，也支持在代码里直接构造。

    from profile import FormatProfile
    p = FormatProfile.load("my.json")
    p.body.size_pt = 12
    p.save("my.json")

JSON 里：
  · 字号可写 "小四" / "12" / "12pt" / {"pt": 12}
  · 行距可写 1.5（倍数）/ "1.5" / "20pt"（固定值）/ {"line": 360, "rule": "auto"}
  · 对齐可写 "left" / "center" / "right" / "both"（也接受 start/end/justify）
  · 颜色写 6 位十六进制，如 "000000"、"C00000"
"""

from __future__ import annotations

import copy
import json
import os
from dataclasses import asdict, dataclass, field, fields
from typing import Any, Dict, List, Optional, Union

# --------------------------------------------------------------------------- #
# 字号
# --------------------------------------------------------------------------- #

CN_SIZE_TO_PT = {
    "初号": 42.0, "小初": 36.0,
    "一号": 26.0, "小一": 24.0,
    "二号": 22.0, "小二": 18.0,
    "三号": 16.0, "小三": 15.0,
    "四号": 14.0, "小四": 12.0,
    "五号": 10.5, "小五": 9.0,
    "六号": 7.5, "小六": 6.5,
    "七号": 5.5, "八号": 5.0,
}

# 反向：pt -> 常用中文名
PT_TO_CN_SIZE = {
    42.0: "初号", 36.0: "小初", 26.0: "一号", 24.0: "小一",
    22.0: "二号", 18.0: "小二", 16.0: "三号", 15.0: "小三",
    14.0: "四号", 12.0: "小四", 10.5: "五号", 9.0: "小五",
    7.5: "六号", 6.5: "小六", 5.5: "七号", 5.0: "八号",
}

ALIGN_MAP = {
    "left": "left", "start": "left", "l": "left",
    "center": "center", "centre": "center", "c": "center", "居中": "center",
    "right": "right", "end": "right", "r": "right",
    "both": "both", "justify": "both", "j": "both", "两端对齐": "both",
    "": None, "none": None, "null": None,
}


class ProfileError(ValueError):
    """配置格式错误。"""


def parse_size(value, default=12.0) -> float:
    """把 "小四" / 12 / "12pt" / {"pt": 12} 解析成磅值。"""
    if value is None:
        return float(default)
    if isinstance(value, dict):
        value = value.get("pt", value.get("size", default))
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().lower()
    if not s:
        return float(default)
    if s in CN_SIZE_TO_PT:
        return CN_SIZE_TO_PT[s]
    s = s.replace("磅", "").replace("pt", "").replace("p", "").strip()
    try:
        return float(s)
    except ValueError:
        raise ProfileError(f"无法识别的字号：{value!r}（可用：小四/12/12pt/{{\"pt\":12}}）")


def parse_align(value, default=None):
    """把 "center" / "居中" 解析成 OOXML 的 jc 值。"""
    if value is None:
        return default
    key = str(value).strip().lower()
    if key in ALIGN_MAP:
        return ALIGN_MAP[key]
    raise ProfileError(f"无法识别的对齐方式：{value!r}（left/center/right/both）")


def parse_line(value, default=1.5):
    """
    返回 (w:line 值, w:lineRule)。
      1.5        -> (360, "auto")   多倍行距
      "20pt"     -> (400, "exact")  固定值 20 磅
      {"line":360,"rule":"auto"} -> (360, "auto")
    """
    if value is None:
        value = default
    if isinstance(value, dict):
        line = value.get("line")
        rule = value.get("rule", "auto")
        if line is None:
            raise ProfileError(f"行距配置缺少 line：{value!r}")
        return int(line), str(rule)
    if isinstance(value, (int, float)):
        return int(round(240 * float(value))), "auto"
    s = str(value).strip().lower()
    if s.endswith("pt"):
        try:
            pt = float(s[:-2].strip())
        except ValueError:
            raise ProfileError(f"无法识别的行距：{value!r}")
        return int(round(pt * 20)), "exact"
    try:
        return int(round(240 * float(s))), "auto"
    except ValueError:
        raise ProfileError(f"无法识别的行距：{value!r}（1.5 / \"20pt\" / "
                           '{"line":360,"rule":"auto"}）')


def parse_color(value, default="000000"):
    if value is None:
        return default
    s = str(value).strip().lstrip("#").upper()
    if len(s) == 6 and all(c in "0123456789ABCDEF" for c in s):
        return s
    if s in ("AUTO", "自动"):
        return "auto"
    raise ProfileError(f"无法识别的颜色：{value!r}（6 位十六进制，如 000000）")


def size_to_json(pt: float):
    """导出时优先写中文号名，读起来更直观。"""
    return PT_TO_CN_SIZE.get(round(float(pt), 2), round(float(pt), 2))


# --------------------------------------------------------------------------- #
# 数据结构
# --------------------------------------------------------------------------- #

@dataclass
class FontSpec:
    """字体：中文用 east_asia，西文/数字用 west，ascii/hAnsi 默认跟 west。"""
    east_asia: str = "宋体"
    west: str = "Times New Roman"

    @classmethod
    def from_json(cls, d, default=None):
        d = d or {}
        if not isinstance(d, dict):
            raise ProfileError(f"字体配置应为对象：{d!r}")
        base = default or cls()
        return cls(east_asia=str(d.get("east_asia", d.get("cn", base.east_asia))),
                   west=str(d.get("west", d.get("en", base.west))))

    def to_json(self):
        return {"east_asia": self.east_asia, "west": self.west}


@dataclass
class ParaSpec:
    """一类段落的完整格式。"""
    size_pt: float = 12.0
    font: FontSpec = field(default_factory=FontSpec)
    bold: Optional[bool] = False          # None = 不改变原有加粗
    italic: Optional[bool] = None
    line: int = 360                       # w:line
    line_rule: str = "auto"               # auto / exact / atLeast
    space_before: float = 0               # 磅
    space_after: float = 0                # 磅
    first_line_indent_chars: Optional[float] = None   # 首行缩进字符数，None=不缩进
    align: Optional[str] = None           # left/center/right/both/None
    color: str = "000000"

    @classmethod
    def from_json(cls, d, default=None):
        base = default or cls()
        if d is None:
            d = {}
        if not isinstance(d, dict):
            raise ProfileError(f"段落配置应为对象：{d!r}")
        d = dict(d)
        line, rule = parse_line(d.get("line_spacing", d.get("line")), default=None) \
            if ("line_spacing" in d or "line" in d) else (base.line, base.line_rule)
        return cls(
            size_pt=parse_size(d.get("size", d.get("size_pt")), base.size_pt),
            font=FontSpec.from_json(d.get("font"), base.font),
            bold=_tri_bool(d.get("bold"), base.bold),
            italic=_tri_bool(d.get("italic"), base.italic),
            line=line, line_rule=rule,
            space_before=float(d.get("space_before", base.space_before) or 0),
            space_after=float(d.get("space_after", base.space_after) or 0),
            first_line_indent_chars=_opt_float(
                d.get("first_line_indent_chars", d.get("indent_chars")),
                base.first_line_indent_chars),
            align=parse_align(d.get("align"), base.align),
            color=parse_color(d.get("color"), base.color),
        )

    def to_json(self):
        d = {
            "size": size_to_json(self.size_pt),
            "font": self.font.to_json(),
            "bold": self.bold,
            "italic": self.italic,
            "line_spacing": (round(self.line / 240.0, 3) if self.line_rule == "auto"
                             else f"{round(self.line / 20.0, 2)}pt"),
            "space_before": self.space_before,
            "space_after": self.space_after,
            "first_line_indent_chars": self.first_line_indent_chars,
            "align": self.align,
            "color": self.color,
        }
        return d


def _tri_bool(v, default):
    if v is None:
        return default
    if isinstance(v, str) and v.strip().lower() in ("", "null", "none", "保持"):
        return None
    return bool(v)


def _opt_float(v, default):
    if v is None:
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


@dataclass
class CaptionSpec(ParaSpec):
    """图表题注：额外带前缀识别规则。"""
    prefixes: List[str] = field(default_factory=lambda: list(DEFAULT_CAPTION_PREFIXES))
    max_length: int = 60

    @classmethod
    def from_json(cls, d, default=None):
        base = default or cls()
        if d is None:
            d = {}
        if not isinstance(d, dict):
            raise ProfileError(f"配置应为对象：{d!r}")
        ps = ParaSpec.from_json(d, ParaSpec(**{k: getattr(base, k)
                                               for k in _PARA_FIELDS}))
        obj = cls(**{**{k: getattr(ps, k) for k in _PARA_FIELDS},
                    "prefixes": list(d.get("prefixes", base.prefixes) or []),
                    "max_length": int(d.get("max_length", base.max_length))})
        return obj

    def to_json(self):
        d = ParaSpec.to_json(self)
        d["prefixes"] = self.prefixes
        d["max_length"] = self.max_length
        return d


_PARA_FIELDS = ("size_pt", "font", "bold", "italic", "line", "line_rule",
                "space_before", "space_after", "first_line_indent_chars",
                "align", "color")

DEFAULT_CAPTION_PREFIXES = ("图", "表", "图表", "Figure", "Table",
                            "FIGURE", "TABLE", "Fig", "Tab")


@dataclass
class FootnoteSpec(ParaSpec):
    """
    脚注区文本。另含正文里"脚注序号"的样式控制：

      reference_size_pt   : 正文中脚注序号的字号；null = 跟随正文字号（默认，即小四）
      reference_superscript: 序号是否上标
      reference_trim_spaces: 去掉序号与正文之间的空格（相邻序号之间的空格保留，
                             否则 "40 41" 会变成 "4041"）
    """
    reference_superscript: bool = True
    reference_size_pt: Optional[float] = None   # None = 跟随 body.size_pt
    reference_trim_spaces: bool = True
    separator_indent: float = 0       # 分割线缩进（磅），0 = 顶格

    @classmethod
    def from_json(cls, d, default=None):
        base = default or cls()
        if d is None:
            d = {}
        if not isinstance(d, dict):
            raise ProfileError(f"配置应为对象：{d!r}")
        ps = ParaSpec.from_json(d, ParaSpec(**{k: getattr(base, k)
                                               for k in _PARA_FIELDS}))
        return cls(**{**{k: getattr(ps, k) for k in _PARA_FIELDS},
                     "reference_superscript": bool(
                         d.get("reference_superscript", base.reference_superscript)),
                     "reference_size_pt": (
                         None if d.get("reference_size_pt",
                                        base.reference_size_pt) is None
                         else parse_size(d.get("reference_size_pt"),
                                         base.reference_size_pt or 12.0)),
                     "reference_trim_spaces": bool(
                         d.get("reference_trim_spaces", base.reference_trim_spaces)),
                     "separator_indent": float(
                         d.get("separator_indent", base.separator_indent) or 0)})

    def to_json(self):
        d = ParaSpec.to_json(self)
        d["reference_superscript"] = self.reference_superscript
        d["reference_size_pt"] = self.reference_size_pt
        d["reference_trim_spaces"] = self.reference_trim_spaces
        d["separator_indent"] = self.separator_indent
        return d


@dataclass
class PageNumberSpec(ParaSpec):
    add_if_missing: bool = True
    position: str = "footer"          # footer / header
    prefix: str = ""                  # 页码前后缀，如 "- "
    suffix: str = ""

    @classmethod
    def from_json(cls, d, default=None):
        base = default or cls()
        if d is None:
            d = {}
        if not isinstance(d, dict):
            raise ProfileError(f"配置应为对象：{d!r}")
        ps = ParaSpec.from_json(d, ParaSpec(**{k: getattr(base, k)
                                               for k in _PARA_FIELDS}))
        return cls(**{**{k: getattr(ps, k) for k in _PARA_FIELDS},
                     "add_if_missing": bool(
                         d.get("add_if_missing", base.add_if_missing)),
                     "position": str(d.get("position", base.position)).lower(),
                     "prefix": str(d.get("prefix", base.prefix)),
                     "suffix": str(d.get("suffix", base.suffix))})

    def to_json(self):
        d = ParaSpec.to_json(self)
        d["add_if_missing"] = self.add_if_missing
        d["position"] = self.position
        d["prefix"] = self.prefix
        d["suffix"] = self.suffix
        return d


@dataclass
class QuoteSpec:
    """
    引号智能归一化。

    normalize_western : 引号内是纯西文时，把引号规范化成英文印刷体引号
    style             : "smart"    -> 英文印刷体弯引号 “ ” / ‘ ’（默认，学术英文标准）
                        "straight" -> 直引号 " 和 '（打字机风格）
                        "curly"    -> 只改字体不动字符（原文已是弯引号时用）
                        "keep"     -> 完全不处理
    scope             : "all" 正文与脚注都处理 / "footnote" 只处理脚注 / "none" 关闭
    require_latin     : inner 里至少要有一个拉丁字母才算"西文"（避免纯数字被误判）
    """

    normalize_western: bool = True
    style: str = "smart"
    scope: str = "all"
    require_latin: bool = True

    @classmethod
    def from_json(cls, d, default=None):
        base = default or cls()
        d = d or {}
        if not isinstance(d, dict):
            raise ProfileError(f"quotes 应为对象：{d!r}")
        style = str(d.get("style", base.style)).strip().lower()
        if style not in ("smart", "straight", "curly", "keep"):
            raise ProfileError(
                f"quotes.style 只能是 smart / straight / curly / keep，收到：{style!r}")
        scope = str(d.get("scope", base.scope)).strip().lower()
        if scope not in ("all", "footnote", "body", "none"):
            raise ProfileError(
                f"quotes.scope 只能是 all / footnote / body / none，收到：{scope!r}")
        return cls(
            normalize_western=bool(
                d.get("normalize_western", base.normalize_western)),
            style=style,
            scope=scope,
            require_latin=bool(d.get("require_latin", base.require_latin)),
        )

    def to_json(self):
        return {"normalize_western": self.normalize_western,
                "style": self.style,
                "scope": self.scope,
                "require_latin": self.require_latin}


@dataclass
class DetectionSpec:
    """标题识别相关开关。"""
    auto_detect_title: bool = True          # 自动识别文章主标题
    detect_plain_title: bool = False        # 首段无标题样式时也当主标题
    skip_table_paragraphs: bool = True      # 表格内文字不参与标题识别
    max_title_length: int = 60              # 主标题最大字数
    max_plain_heading_length: int = 40      # 纯文本编号标题最大字数
    level1_pattern: str = r"^\s*[一二三四五六七八九十百]+[、\.]"
    level2_pattern: str = r"^\s*[（(]\s*[一二三四五六七八九十百]+\s*[)）]"
    level3_pattern: str = r"^\s*\d{1,3}\s*[\.、]\s*\S"
    level3_alt_pattern: str = r"^\s*[（(]\s*\d{1,3}\s*[)）]"

    @classmethod
    def from_json(cls, d, default=None):
        base = default or cls()
        d = d or {}
        if not isinstance(d, dict):
            raise ProfileError(f"heading_detection 应为对象：{d!r}")
        out = cls()
        for f in fields(cls):
            if f.name in d:
                v = d[f.name]
                if isinstance(f.type, str) and f.type == "bool":
                    v = bool(v)
                elif f.type == "int":
                    v = int(v)
                setattr(out, f.name, v)
            else:
                setattr(out, f.name, getattr(base, f.name))
        # 校验正则可编译
        import re
        for name in ("level1_pattern", "level2_pattern",
                     "level3_pattern", "level3_alt_pattern"):
            try:
                re.compile(getattr(out, name))
            except re.error as e:
                raise ProfileError(f"{name} 不是合法正则：{e}")
        return out


@dataclass
class FormatProfile:
    """一份完整的排版方案。"""
    name: str = "默认（中文论文/公文）"
    version: int = 1
    color: str = "000000"                  # 全文字色兜底
    body: ParaSpec = field(default_factory=ParaSpec)
    title: ParaSpec = field(default_factory=ParaSpec)
    heading1: ParaSpec = field(default_factory=ParaSpec)
    heading2: ParaSpec = field(default_factory=ParaSpec)
    heading3: ParaSpec = field(default_factory=ParaSpec)
    caption: CaptionSpec = field(default_factory=CaptionSpec)
    footnote: FootnoteSpec = field(default_factory=FootnoteSpec)
    page_number: PageNumberSpec = field(default_factory=PageNumberSpec)
    detection: DetectionSpec = field(default_factory=DetectionSpec)
    quotes: QuoteSpec = field(default_factory=QuoteSpec)
    cjk_punct_enabled: bool = True
    cjk_punct_chars: str = "“”‘’「」『』《》〈〉、。，；：！？——…"
    force_all_black: bool = True

    # ------------------------------------------------------------------ #
    # 默认方案：完全对应你最初提出的要求
    # ------------------------------------------------------------------ #
    @classmethod
    def default(cls):
        return cls(
            name="默认（中文论文/公文）",
            color="000000",
            body=ParaSpec(
                size_pt=12.0,                       # 小四
                font=FontSpec("宋体", "Times New Roman"),
                bold=False,
                line=360, line_rule="auto",         # 1.5 倍
                space_before=0, space_after=0,
                first_line_indent_chars=2,
                align=None,
                color="000000",
            ),
            title=ParaSpec(
                size_pt=15.0,                       # 小三
                font=FontSpec("黑体", "Times New Roman"),
                bold=False,
                line=360, line_rule="auto",
                space_before=0, space_after=0,
                first_line_indent_chars=None,
                align="center",
                color="000000",
            ),
            heading1=ParaSpec(
                size_pt=14.0,                       # 四号
                font=FontSpec("黑体", "Times New Roman"),
                bold=False,
                line=360, line_rule="auto",
                space_before=0, space_after=0,
                first_line_indent_chars=None,
                align="center",
                color="000000",
            ),
            heading2=ParaSpec(
                size_pt=12.0,                       # 小四
                font=FontSpec("宋体", "Times New Roman"),
                bold=True,
                line=360, line_rule="auto",
                space_before=0, space_after=0,
                first_line_indent_chars=2,
                align=None,
                color="000000",
            ),
            heading3=ParaSpec(
                size_pt=12.0,
                font=FontSpec("宋体", "Times New Roman"),
                bold=False,
                line=360, line_rule="auto",
                space_before=0, space_after=0,
                first_line_indent_chars=2,
                align=None,
                color="000000",
            ),
            caption=CaptionSpec(
                size_pt=10.5,                       # 五号
                font=FontSpec("宋体", "Times New Roman"),
                bold=False,
                line=360, line_rule="auto",
                space_before=0, space_after=0,
                first_line_indent_chars=None,
                align="center",
                color="000000",
                prefixes=list(DEFAULT_CAPTION_PREFIXES),
            ),
            footnote=FootnoteSpec(
                size_pt=9.0,                        # 小五
                font=FontSpec("宋体", "Times New Roman"),
                bold=False,
                line=240, line_rule="auto",         # 单倍
                space_before=0, space_after=0,
                first_line_indent_chars=None,       # 不缩进
                align=None,
                color="000000",
                reference_superscript=True,
                separator_indent=0,                 # 分割线顶格
            ),
            page_number=PageNumberSpec(
                size_pt=9.0,                        # 小五
                font=FontSpec("宋体", "Times New Roman"),
                bold=False,
                line=240, line_rule="auto",
                space_before=0, space_after=0,
                first_line_indent_chars=None,
                align="center",                     # 居中
                color="000000",
                add_if_missing=True,
                position="footer",
            ),
            detection=DetectionSpec(),
            quotes=QuoteSpec(),
        )

    # ------------------------------------------------------------------ #
    # 序列化
    # ------------------------------------------------------------------ #
    def to_dict(self):
        return {
            "name": self.name,
            "version": self.version,
            "color": self.color,
            "force_all_black": self.force_all_black,
            "body": self.body.to_json(),
            "title": self.title.to_json(),
            "heading1": self.heading1.to_json(),
            "heading2": self.heading2.to_json(),
            "heading3": self.heading3.to_json(),
            "caption": self.caption.to_json(),
            "footnote": self.footnote.to_json(),
            "page_number": self.page_number.to_json(),
            "heading_detection": _detection_to_json(self.detection),
            "quotes": self.quotes.to_json(),
            "cjk_punctuation": {
                "enabled": self.cjk_punct_enabled,
                "chars": self.cjk_punct_chars,
            },
        }

    def to_json(self, ensure_ascii=False, indent=2):
        return json.dumps(self.to_dict(), ensure_ascii=ensure_ascii, indent=indent)

    def save(self, path):
        d = os.path.dirname(os.path.abspath(path))
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(self.to_json())
        return path

    @classmethod
    def from_dict(cls, d):
        if not isinstance(d, dict):
            raise ProfileError("配置根节点必须是 JSON 对象")
        base = cls.default()
        for key in d:
            if str(key).startswith("_"):
                continue          # _readme 之类的说明字段，直接忽略
            if key not in _ALLOWED_ROOT_KEYS:
                raise ProfileError(f"未知的配置项：{key!r}（拼写错误？"
                                   f"可用：{', '.join(sorted(_ALLOWED_ROOT_KEYS))}）")
        return cls(
            name=str(d.get("name", base.name)),
            version=int(d.get("version", base.version)),
            color=parse_color(d.get("color"), base.color),
            force_all_black=bool(d.get("force_all_black", base.force_all_black)),
            body=ParaSpec.from_json(d.get("body"), base.body),
            title=ParaSpec.from_json(d.get("title"), base.title),
            heading1=ParaSpec.from_json(d.get("heading1"), base.heading1),
            heading2=ParaSpec.from_json(d.get("heading2"), base.heading2),
            heading3=ParaSpec.from_json(d.get("heading3"), base.heading3),
            caption=CaptionSpec.from_json(d.get("caption"), base.caption),
            footnote=FootnoteSpec.from_json(d.get("footnote"), base.footnote),
            page_number=PageNumberSpec.from_json(d.get("page_number"),
                                                 base.page_number),
            detection=DetectionSpec.from_json(d.get("heading_detection"),
                                              base.detection),
            quotes=QuoteSpec.from_json(d.get("quotes"), base.quotes),
            cjk_punct_enabled=bool(
                (d.get("cjk_punctuation") or {}).get("enabled", True)),
            cjk_punct_chars=str(
                (d.get("cjk_punctuation") or {}).get("chars",
                                                     base.cjk_punct_chars)),
        )

    @classmethod
    def load(cls, path_or_str):
        """
        从文件路径或 JSON 字符串加载。
        出错时抛出 ProfileError，消息里带具体字段，便于排查。
        """
        try:
            if isinstance(path_or_str, dict):
                data = path_or_str
            elif os.path.exists(str(path_or_str)):
                with open(path_or_str, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = json.loads(str(path_or_str))
        except json.JSONDecodeError as e:
            raise ProfileError(f"JSON 语法错误：第 {e.lineno} 行第 {e.colno} 列 —— {e.msg}")
        except OSError as e:
            raise ProfileError(f"读取配置失败：{e}")
        if not isinstance(data, dict):
            raise ProfileError("配置根节点必须是 JSON 对象")
        return cls.from_dict(data)

    def copy(self):
        return copy.deepcopy(self)

    def summary(self):
        """一行式摘要，用于 GUI 状态栏。"""
        return (f"{self.name}｜正文 {_size_name(self.body.size_pt)}"
                f"·{_line_name(self.body.line, self.body.line_rule)}｜"
                f"主标题 {_size_name(self.title.size_pt)}｜"
                f"脚注 {_size_name(self.footnote.size_pt)}")


_ALLOWED_ROOT_KEYS = {
    "name", "version", "color", "force_all_black",
    "body", "title", "heading1", "heading2", "heading3",
    "caption", "footnote", "page_number", "heading_detection",
    "cjk_punctuation", "quotes",
}


def _detection_to_json(det: DetectionSpec):
    return {f.name: getattr(det, f.name) for f in fields(DetectionSpec)}


def _size_name(pt):
    return PT_TO_CN_SIZE.get(round(float(pt), 2), f"{pt:g}pt")


def _line_name(line, rule):
    if rule == "auto":
        return f"{line / 240.0:g}倍"
    return f"{line / 20.0:g}pt"


def write_default_template(path):
    """导出一份带中文注释说明的模板（JSON 不支持注释，说明放在 _readme 字段里）。"""
    prof = FormatProfile.default()
    d = prof.to_dict()
    d["_readme"] = _TEMPLATE_README
    d = {"_readme": d.pop("_readme"), **d}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    return path


_TEMPLATE_README = [
    "WordFormat 排版方案模板 —— 改完保存，在程序里「导入配置」即可生效。",
    "",
    "【字号】写中文号名（初号/小初/一号/小一/二号/小二/三号/小三/四号/小四/五号/小五/六号/小六）或磅数（12 / \"12pt\"）。",
    "【行距】数字=倍数（1.5 即 1.5 倍）；\"20pt\"=固定值 20 磅；也可写 {\"line\":360,\"rule\":\"auto\"}（240=单倍，360=1.5倍）。",
    "【对齐】left / center / right / both（两端对齐）；null 表示不改原文对齐。",
    "【bold/italic】true / false / null（null = 保持原文不变）。",
    "【first_line_indent_chars】首行缩进字符数（2 = 缩进 2 字符）；null 或 0 = 不缩进。",
    "【color】6 位十六进制，如 000000 黑色、C00000 红色；全文字色兜底见根节点 color。",
    "【font】east_asia=中文字体（含中文引号），west=西文/数字字体。",
    "【caption.prefixes】图表题注识别前缀，可自行增删。",
    "【footnote.reference_size_pt】正文中脚注序号的字号（如 12 或 \"小四\"）；null = 跟随正文字号。",
    "【footnote.reference_trim_spaces】去掉序号与正文之间的空格（默认 true）；相邻序号之间的空格保留。",
    "【page_number.add_if_missing】文档没有页码时是否自动添加；position 可填 footer/header。",
    "【heading_detection】标题识别开关；auto_detect_title=false 时不再自动识别文章主标题。",
    "【cjk_punctuation】中文标点（，。、等）是否强制用中文字体。"
    "引号的半角/全角由【quotes】控制。",
    "【quotes】西文引文的引号处理：normalize_western=true 时，引号内是纯西文就自动转半角；",
    "          style=straight 转成直引号 \" 和 '（默认，真正的半角）；curly 保留弯引号但用西文字体渲染；keep 不转换。",
    "          scope=all（正文+脚注，默认）/ footnote（只处理脚注）/ body（只处理正文）/ none（关闭）。",
    "          例：脚注里「“The ‘Concept’ of Communication”」会自动变成「\"The 'Concept' of Communication\"」，而「“国际传播”」保持全角。",
    "",
    "不需要的字段可以直接删掉，删掉的项会沿用默认方案。",
]


# --------------------------------------------------------------------------- #
# 内置预设方案
# --------------------------------------------------------------------------- #

def _preset(name, **kw):
    """基于默认方案改几个字段生成预设。"""
    p = FormatProfile.default()
    p.name = name
    for k, v in kw.items():
        setattr(p, k, v)
    return p


def builtin_profiles():
    """返回 {显示名: FormatProfile}。"""
    from dataclasses import replace

    # --- 毕业论文：正文小四宋体 1.5 倍，标题黑体 ---
    thesis = FormatProfile.default()
    thesis.name = "毕业论文（小四·1.5倍）"
    thesis.title = replace(thesis.title, size_pt=22.0, bold=False)     # 二号
    thesis.heading1 = replace(thesis.heading1, size_pt=15.0)           # 小三
    thesis.heading2 = replace(thesis.heading2, size_pt=13.0, bold=True)  # 13pt
    thesis.heading3 = replace(thesis.heading3, size_pt=12.0, bold=True)

    # --- 党政机关公文 GB/T 9704-2012 ---
    official = FormatProfile.default()
    official.name = "党政机关公文（GB/T 9704）"
    song3 = FontSpec("仿宋_GB2312", "Times New Roman")
    official.body = ParaSpec(size_pt=16.0, font=song3, bold=False,
                             line=600, line_rule="exact",      # 固定值 30 磅
                             space_before=0, space_after=0,
                             first_line_indent_chars=2, align="both")
    official.title = ParaSpec(size_pt=22.0,                    # 二号
                              font=FontSpec("方正小标宋简体", "Times New Roman"),
                              bold=False, line=600, line_rule="exact",
                              space_before=0, space_after=0,
                              first_line_indent_chars=None, align="center")
    official.heading1 = ParaSpec(size_pt=16.0,                 # 三号黑体
                                 font=FontSpec("黑体", "Times New Roman"),
                                 bold=False, line=600, line_rule="exact",
                                 space_before=0, space_after=0,
                                 first_line_indent_chars=2, align=None)
    official.heading2 = ParaSpec(size_pt=16.0,                 # 三号楷体加粗
                                 font=FontSpec("楷体_GB2312", "Times New Roman"),
                                 bold=True, line=600, line_rule="exact",
                                 space_before=0, space_after=0,
                                 first_line_indent_chars=2, align=None)
    official.heading3 = ParaSpec(size_pt=16.0,                 # 三号仿宋加粗
                                 font=FontSpec("仿宋_GB2312", "Times New Roman"),
                                 bold=True, line=600, line_rule="exact",
                                 space_before=0, space_after=0,
                                 first_line_indent_chars=2, align=None)

    # --- 期刊投稿：五号宋体，行距 1.25 倍，更紧凑 ---
    journal = FormatProfile.default()
    journal.name = "期刊投稿（五号·紧凑）"
    journal.body = ParaSpec(size_pt=10.5, font=FontSpec("宋体", "Times New Roman"),
                            bold=False, line=300, line_rule="auto",   # 1.25 倍
                            space_before=0, space_after=0,
                            first_line_indent_chars=2, align="both")
    journal.title = ParaSpec(size_pt=16.0, font=FontSpec("黑体", "Times New Roman"),
                             bold=False, line=300, line_rule="auto",
                             space_before=0, space_after=0,
                             first_line_indent_chars=None, align="center")
    journal.heading1 = ParaSpec(size_pt=12.0, font=FontSpec("黑体", "Times New Roman"),
                                bold=False, line=300, line_rule="auto",
                                space_before=0, space_after=0,
                                first_line_indent_chars=None, align=None)
    journal.heading2 = ParaSpec(size_pt=10.5, font=FontSpec("黑体", "Times New Roman"),
                                bold=False, line=300, line_rule="auto",
                                space_before=0, space_after=0,
                                first_line_indent_chars=2, align=None)
    journal.heading3 = ParaSpec(size_pt=10.5, font=FontSpec("宋体", "Times New Roman"),
                                bold=True, line=300, line_rule="auto",
                                space_before=0, space_after=0,
                                first_line_indent_chars=2, align=None)

    return {
        FormatProfile.default().name: FormatProfile.default(),
        thesis.name: thesis,
        official.name: official,
        journal.name: journal,
    }
