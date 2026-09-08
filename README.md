# WordFormat —— Word 论文/公文格式一键排版工具

一个零配置、可打包成 exe 的开源小工具。把 `.docx` 丢进去，一键套用统一的中文论文/公文排版规范：
正文、各级标题、脚注、页码、图表题注、字色，全部一次到位。

> 重点解决了一个常见痛点：**很多排版工具写死"Heading1 = 一级标题"**，
> 于是用 Heading1 排的**文章主标题**被当成"一、二、三"级，真正的"一、"级（Heading2）被当成"（一）"级，
> **全局错一位**。本工具改为按文档中**实际出现的最浅层级**自动对齐，永不错位（详见 [标题层级判定](#标题层级判定为什么不会错一位)）。

---

## 一、排版规则

| 对象 | 字体 | 字号 | 对齐 / 缩进 | 行距 | 段前段后 |
|---|---|---|---|---|---|
| 文章主标题 | 黑体（中文）+ Times New Roman（西文） | 小三 15pt | 居中 | 1.5 倍 | 0 / 0 |
| 一级标题（一、二、三…） | 黑体 + Times New Roman | 四号 14pt | 居中 | 1.5 倍 | 0 / 0 |
| 二级标题（（一）（二）…） | 宋体 + Times New Roman | 小四 12pt **加粗** | 首行缩进 2 字符 | 1.5 倍 | 0 / 0 |
| 三级标题（1. 2. 3. …） | 宋体 + Times New Roman | 小四 12pt | 首行缩进 2 字符 | 1.5 倍 | 0 / 0 |
| 正文 | 宋体 + Times New Roman | 小四 12pt | 首行缩进 2 字符 | 1.5 倍 | 0 / 0 |
| 脚注区 | 宋体 + Times New Roman | 小五 9pt | **不缩进** | **单倍** | 0 / 0 |
| 脚注分割线 | — | — | **顶格不缩进** | — | — |
| 页码 | Times New Roman | 小五 9pt | 页面底部**居中** | — | — |
| 图表标题 | 宋体 + Times New Roman | 五号 10.5pt | 居中 | 1.5 倍 | 0 / 0 |
| 全文 | **字色统一黑色**（含页眉页脚、脚注、表格、文本框） | | | | |

补充说明：

- **引号跟随内容语言**（默认开启）：看引号里写的是什么，而不是简单地一刀切。

  | 原文 | 处理后 | 字体 |
  |---|---|---|
  | `“The ‘Concept’ of Communication”` | `“The ‘Concept’ of Communication”` | 弯引号 + Times New Roman（视觉半角） |
  | `"A Conceptual Model for Communications Research"` | `“A Conceptual Model for Communications Research”` | 弯引号 + Times New Roman |
  | `“国际传播”` | `“国际传播”` | 全角 + 宋体 |

  书名号、顿号、逗号等纯中文标点始终用中文字体，不会出现「正文是宋体、标点却是 Times New Roman」的割裂效果。
  详见 [西文引文自动转半角引号](#西文引文自动转半角引号)。
- **正文中的脚注序号**：小四（跟随正文字号）、上标、Times New Roman、黑色，
  且序号与正文之间不留空格（`……意识形态。 40 41 另有他说` -> `……意识形态。40 41另有他说`；
  相邻序号之间的空格保留并规范为 1 个，避免出现 `4041`）。
- **首行缩进**使用 `w:firstLineChars`（按字符），不是按磅值，所以字号变化时缩进始终是 2 个字符。
- 文档**没有页码**时，会自动在页脚插入居中页码（可用 `--no-page-number` 关闭）。
- 自动编号/项目符号段落会保留原有缩进，不会叠加首行缩进。
- 表格单元格里的文字按正文格式处理，且不会参与标题识别（避免误判）。

---

## 二、自定义格式（JSON 方案）

内置规则不满足时，改一份 JSON 就行，**不用改代码**。

### 图形界面

主界面中部「排版方案」区：

- **预设下拉框**：内置 4 套常用方案
- **导入配置…**：选一个 `.json` 立即生效
- **导出当前方案**：把当前配置存成 JSON（改一改再导入，最省事）
- **导出模板（带说明）**：导出一份带 `_readme` 字段的说明模板，照着改
- **查看方案**：弹出当前方案的人可读摘要
- **恢复默认**：一键回到默认规则

### 命令行

```bash
python app.py -c my.json  a.docx b.docx      # 用自定义方案处理
python app.py --export-template 模板.json     # 导出带说明的模板
python app.py -c my.json --export-config out.json   # 方案另存
python app.py -c my.json --show-config       # 打印方案摘要后退出
```

`format_document()` 也直接接受方案对象 / 路径 / 字典：

```python
from formatter import format_document
from profile import FormatProfile

format_document("in.docx", "out.docx", profile="my.json")
format_document("in.docx", "out.docx", profile={"body": {"size": "三号"}})
```

### 配置长什么样

```json
{
  "name": "我的方案",
  "body": {
    "size": "小四",
    "font": { "east_asia": "宋体", "west": "Times New Roman" },
    "line_spacing": 1.5,
    "space_before": 0,
    "space_after": 0,
    "first_line_indent_chars": 2,
    "align": null,
    "bold": false,
    "color": "000000"
  },
  "title":   { "size": "小三", "font": {"east_asia": "黑体"}, "align": "center" },
  "heading1":{ "size": "四号", "font": {"east_asia": "黑体"}, "align": "center" },
  "heading2":{ "size": "小四", "bold": true, "first_line_indent_chars": 2 },
  "heading3":{ "size": "小四", "first_line_indent_chars": 2 },
  "caption": { "size": "五号", "align": "center",
               "prefixes": ["图", "表", "Figure", "Table"] },
  "footnote":{ "size": "小五", "line_spacing": 1.0,
               "first_line_indent_chars": null, "separator_indent": 0 },
  "page_number": { "size": "小五", "align": "center",
                   "add_if_missing": true, "prefix": "", "suffix": "" },
  "heading_detection": { "auto_detect_title": true,
                         "level1_pattern": "^\\s*[一二三四五六七八九十百]+[、.]" }
}
```

**可以只写要改的字段**，没写的自动沿用默认方案。上面这份等价于：

```json
{ "body": { "size": "小四" } }
```

### 写法约定

| 字段 | 可写的值 |
|---|---|
| `size` | `小四` / `12` / `"12pt"` / `{"pt": 12}`；中文号名支持 初号~八号 |
| `line_spacing` | `1.5` = 1.5 倍；`"20pt"` = 固定值 20 磅；`{"line":360,"rule":"auto"}`（240=单倍，360=1.5倍） |
| `align` | `left` / `center` / `right` / `both`（两端对齐）；`null` = 不改原文 |
| `bold` / `italic` | `true` / `false` / `null`（null = 保持原文不变） |
| `first_line_indent_chars` | 首行缩进字符数，`2` = 缩进 2 字符；`null` 或 `0` = 不缩进 |
| `color` | 6 位十六进制：`000000` 黑色、`C00000` 红色；根节点 `color` 是全文字色兜底 |
| `font.east_asia` / `font.west` | 中文字体（含中文引号）/ 西文与数字字体 |

### 其它可选项

| 位置 | 字段 | 说明 |
|---|---|---|
| `caption` | `prefixes` / `max_length` | 图表题注识别前缀与最大长度 |
| `footnote` | `separator_indent` | 脚注分割线缩进（磅），`0` = 顶格 |
| `footnote` | `reference_size_pt` | 正文里脚注序号的字号（如 `12` / `"小四"`）；`null` = 跟随正文字号（默认） |
| `footnote` | `reference_superscript` | 正文里的脚注序号是否上标（默认 true） |
| `footnote` | `reference_trim_spaces` | 去掉序号与正文之间的空格（默认 true） |
| `page_number` | `add_if_missing` / `position` / `prefix` / `suffix` | 缺页码时是否自动补、放页脚还是页眉、页码前后缀（如 `- 1 -`） |
| `heading_detection` | `auto_detect_title` | 关掉后不再自动识别文章主标题 |
| `heading_detection` | `detect_plain_title` | 首段没有标题样式时，也当主标题 |
| `heading_detection` | `level1~3_pattern` | 编号识别正则，可适配「第一章」「1.1」等写法 |
| `cjk_punctuation` | `enabled` / `chars` | 中文标点（引号等）是否强制用中文字体 |
| `quotes` | `normalize_western` | 引号内是西文时自动规范为英文印刷体引号（默认开） |
| `quotes` | `style` | `smart`（默认，英文弯引号 `“ ”` `‘ ’`）/ `straight`（直引号 `" "` `' '`）/ `curly`（只改字体不动字符）/ `keep`（不处理） |
| `quotes` | `scope` | `all`（默认，正文+脚注）/ `footnote` / `body` / `none` |
| 根节点 | `force_all_black` | 是否把全文字色统一刷黑 |

配置写错时会明确指出问题所在，例如：

```
[错误] 配置无法使用：未知的配置项：'bodyy'（拼写错误？可用：body, caption, ...）
[错误] 配置无法使用：JSON 语法错误：第 8 行第 5 列 —— Expecting ',' delimiter
```

### 西文引文的引号（smart quotes）

默认开启，默认风格 `smart` —— **英文印刷体弯引号** `“ ”` `‘ ’`，用 Times New Roman 渲染，
视觉上就是半角。这是 MLA / APA / Chicago 等学术英文排版的标准做法。

判断依据是**看引号里写的是什么**：不含中日韩字符、且至少含一个拉丁字母 → 判为西文引文；
含中文则保持全角 `“ ”` + 宋体。

```json
{ "quotes": { "normalize_western": true, "style": "smart", "scope": "all" } }
```

效果（脚注里的文献条目）：

| 处理前 | 处理后（smart，默认） | 处理后（straight） |
|---|---|---|
| `“The ‘Concept’ of Communication”` | `“The ‘Concept’ of Communication”` | `"The 'Concept' of Communication"` |
| `"A Conceptual Model for Communications Research"` | `“A Conceptual Model for Communications Research”` | `"A Conceptual Model for Communications Research"` |
| `'An Extension of the "Lasswell Formula",'` | `‘An Extension of the “Lasswell Formula”,’` | `'An Extension of the "Lasswell Formula",'` |
| `“国际传播”` | `“国际传播”`（不变） | `“国际传播”`（不变） |
| `“2023—2024”` | `“2023—2024”`（不变，纯数字不算西文） | 同左 |

四种风格怎么选：

| style | 结果 | 适合 |
|---|---|---|
| `smart` | `“Concept”` `‘inner’` | **默认**。英文学术文献、论文参考文献 |
| `straight` | `"Concept"` `'inner'` | 偏好打字机直引号，或目标刊物明确要求 |
| `curly` | 字符原样不动，只把全角引号改用西文字体渲染 | 原文已是规范的弯引号，只想修字体 |
| `keep` | 完全不处理 | 想自己控制 |

细节：

- **直引号也会被扳弯**：`smart` 不只处理全角 `“ ”`，也会把半角 `"` `'` 转成弯引号
  （按前后文判断是开引号还是闭引号），所以原文已经是半角的文献条目同样会被规范。
- **嵌套层级保持原样**：`‘外 “内” 外’` 不会重排成 `“外 ‘内’ 外”`，只把直引号扳弯、字体改对。
- **字体**：转换后的弯引号按西文字体渲染，同 run 里的中文仍是宋体。
- **不误伤单位**：`5' 6"` 这类数字后的直引号不参与转换。
- 脚注与尾注（`endnotes.xml`）都处理；`scope` 可选 `all` / `footnote` / `body` / `none`。

### 正文中的脚注序号

默认：**小四（跟随正文字号）、上标、Times New Roman、黑色**，且序号与正文之间不留空格。

```
处理前    ……族主义和意识形态。 40 41 另有他说。
处理后    ……族主义和意识形态。40 41另有他说。
          （40 后的空格删掉；40 与 41 之间的空格保留 1 个，否则会变成 4041）
```

细节：

- 只删**序号与正文之间**的空格；与序号无关的普通空格（如 `中文 与 English`）一个都不动。
- 序号之间的连续空格规范化为 1 个。
- 半角、全角、不换行空格都处理。
- 跨 run 也能处理（Word 常把序号拆成独立 run）。
- 尾注序号（endnoteReference）同样处理。

```json
{ "footnote": { "reference_size_pt": null, "reference_superscript": true,
                "reference_trim_spaces": true } }
```

`reference_size_pt` 写 `null` 表示跟随正文字号；也可单独指定 `"小五"` 或 `9`。

### 内置预设

| 预设 | 要点 |
|---|---|
| 默认（中文论文/公文） | 你最初要求的那套规则 |
| 毕业论文（小四·1.5倍） | 主标题二号、一级小三、正文小四 1.5 倍 |
| 党政机关公文（GB/T 9704） | 正文三号仿宋、固定行距 30 磅、标题二号小标宋 |
| 期刊投稿（五号·紧凑） | 正文五号、1.25 倍行距、标题黑体 |

> 预设里用到的 `仿宋_GB2312` / `方正小标宋简体` 等字体需要本机已安装，
> 没装的话 Word 会回退显示，文档本身不会出错 —— 换成你有的字体名即可。

---

## 三、快速开始

### 1. 图形界面（推荐）

```bash
pip install -r requirements.txt
python app.py            # 不带参数即启动 GUI
```

拖拽/选择文件 → 选输出方式 →「开始排版」。

### 2. 命令行

```bash
# 另存为「原名_已排版.docx」
python app.py a.docx b.docx

# 批量处理整个文件夹，结果输出到 out/
python app.py ./论文目录 -o out

# 直接覆盖原文件，并先备份
python app.py a.docx --inplace --backup

# 递归处理子文件夹；不自动补页码
python app.py ./目录 -r --no-page-number
```

常用参数：

| 参数 | 说明 |
|---|---|
| `-o, --outdir` | 输出文件夹 |
| `--inplace` | 直接覆盖原文件 |
| `--backup` | 处理前备份为 `原名.bak.docx` |
| `--no-page-number` | 缺少页码时不自动添加 |
| `--plain-title` | 首段没用标题样式时，也尝试识别为文章主标题 |
| `-r, --recursive` | 递归处理子文件夹 |
| `--gui` | 强制启动图形界面 |

### 3. 打包成 exe（Windows）

前置条件（打包机需要，生成的 exe 分发给他人时不需要）：

- Windows + Python 3.8+，安装时勾选 **Add python.exe to PATH**
- 能连外网装包（或配置国内镜像）

```bat
build.bat                 :: 双击即可，或：pyinstaller app.spec --noconfirm --clean
```

产物：`dist\WordFormat.exe`，单文件、双击即用，拷到任何 Windows 电脑都能跑，无需安装 Python。
（Linux/macOS 可跑 `./build.sh`，生成对应平台的可执行文件。）

> PyInstaller **不能跨平台生成 exe**，Windows 版 exe 必须在 Windows 上打包。

不想打包也能用：双击 `run.bat`（自动装依赖并直接启动 GUI），或 `python app.py`。

#### 双击 .bat / .exe 一闪而过？

按顺序排查：

| 现象 | 原因 | 解决 |
|---|---|---|
| 双击 `build.bat` 立刻关闭 | **批处理文件是 UTF-8 编码或含 `chcp 65001`**，cmd 按字节读脚本会直接中断；另一可能是没装 Python / 没勾 PATH | 仓库里的 `build.bat` 已是**纯 ASCII + CRLF**，重新下载覆盖即可；确认命令行 `python --version` 有输出 |
| `python` 弹出微软商店后消失 | Win10/11 的 `python` 是商店别名 | `build.bat` 已依次尝试 `py -3` → `python` → `python3`，仍不行就手动装 Python 并勾选 PATH |
| 双击 `WordFormat.exe` 无反应 | GUI 版没有控制台，出错看不见 | 同目录下会生成 **`crash.log`**；或改用 `pyinstaller app_console.spec` 打一份带控制台的 `WordFormat_Debug.exe` |
| 杀软报毒 / exe 被删 | PyInstaller 单文件 exe 常被误报 | 把项目目录加入杀软白名单，或改用 `--onedir` 模式 |

环境自检（把输出发给我就能定位）：

```bat
python app.py --selftest
```

国内镜像加速装包：

```bat
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

---

## 四、标题层级判定：为什么不会"错一位"

### 问题是怎么产生的

Word 里作者会用不同的样式排同一份逻辑结构，例如：

```
主标题  -> Heading 1
一、    -> Heading 2
（一）  -> Heading 3
1.      -> Heading 4
```

写死 `Heading1→一级、Heading2→二级、Heading3→三级` 的代码，会把主标题排成"一、"级（黑体四号居中），
把"一、"排成"（一）"级（宋体加粗缩进）——**整体错一位**。

### 本工具的做法

1. **收集**：扫描全文，取每个标题段落的"原始大纲级别"
   （依次读 `outlineLvl` → 样式的 `outlineLvl` → `basedOn` 继承链 → 样式名 `Heading N` / `标题 N` / `Title` → styleId）。
   完全没有标题样式的纯文本文档，则按编号推断：`一、`→1，`（一）`→2，`1.`→3。
2. **识别主标题**：最浅那一层**只有 1 个**、且是**全文第一个非空段落**、且**不带编号**（不是"一、/（一）/1."开头）时，判定为文章主标题。
3. **重映射**：把剩下的层级按"从浅到深"依次对齐到 1 / 2 / 3 级，而不是写死映射。

于是下面这些写法都能得到正确结果：

| 文档写法 | 判定结果 |
|---|---|
| Heading1(主标题) + Heading2(一、) + Heading3(（一）) + Heading4(1.) | 主标题 + 一级 + 二级 + 三级 ✅ |
| Title(主标题) + Heading1(一、) + Heading2(（一）) | 主标题 + 一级 + 二级 ✅ |
| 开头是正文 + Heading1(一、) + Heading2(（一）) + Heading3(1.) | 一级 + 二级 + 三级（无主标题）✅ |
| Heading1(一、) + Heading1(二、)（多个同级） | 两个都是一级，不误判主标题 ✅ |
| 纯文本，只有 `一、` `（一）` `1.` 编号 | 一级 + 二级 + 三级 ✅ |

若首段确实没用任何标题样式、你也想让它当主标题，勾选「首段没有标题样式时，也尝试识别为文章主标题」（CLI：`--plain-title`）。

---

## 五、目录结构

```
word-formatter/
├── formatter.py        # 核心排版库（可单独 import 复用）
├── profile.py          # 排版方案：JSON 导入/导出、内置预设
├── app.py              # 入口：无参数启动 GUI，带参数走 CLI
├── app.spec            # PyInstaller 打包配置（GUI，无控制台）
├── app_console.spec    # 调试用打包配置（带控制台，报错可见）
├── build.bat           # Windows 打包脚本 -> dist\WordFormat.exe
├── run.bat             # 不打包，直接装依赖并启动 GUI
├── build.sh            # Linux/macOS 打包脚本
├── requirements.txt
├── tests/
│   ├── make_sample.py  # 生成 5 种典型写法的测试样本（含脚注/页码/题注/表格）
│   ├── verify.py       # 对排版结果做 XML 级断言校验
│   ├── test_profile.py # 自定义 JSON 方案测试
│   └── test_gui.py     # GUI 冒烟测试（无 tkinter 时用假 tkinter）
└── README.md
```

### 复用核心库

```python
from formatter import format_document

format_document("in.docx", "out.docx",
                add_page_number=True,     # 缺页码时自动补
                format_caption=True,      # 识别并居中图表题注
                detect_plain_title=False, # 首段无样式时也尝试识别为主标题
                log=print)
```

---

## 六、测试

```bash
python tests/make_sample.py    # 生成 5 个覆盖不同标题写法的样本（含脚注、页码、题注、表格）
python tests/verify.py         # XML 级断言：字号/字体/行距/缩进/对齐/字色/脚注/页码/OOXML 顺序
python tests/test_profile.py   # JSON 方案：解析、往返、部分字段、错误处理、预设生效
python tests/test_quotes.py    # 引号语言：西文转半角、中文保全角、尾注/跨 run/嵌套/开关
python tests/test_footnote_refs.py  # 正文脚注序号：小四/上标/去空格/尾注/配置开关
python tests/test_gui.py       # GUI 冒烟测试（无 tkinter 时自动用假 tkinter）
python tests/stress.py         # 混合内容压力测试（页眉/图片/超链接/表格/空段落）
```

`verify.py` 会逐项断言排版结果，另外还可用 LibreOffice 实转 PDF 验证文档未被写坏：

```bash
soffice --headless --convert-to pdf --outdir out out/*.docx
```

---

## 七、已知限制

- 只处理 `.docx`。`.doc` 会尝试用 LibreOffice/WPS 或 Word（COM）自动转换，失败时会提示你先另存为 `.docx`。
- 只改**格式**，不改文字：标题原有的编号（如"一、"）保持原样，不会自动增删或重排编号。
- 图表标题按「图/表 + 数字」开头、或 Word 题注样式（`Caption`）识别；自定义前缀可在 `formatter._is_caption_para` 中扩展。
- 超过三级的标题统一按三级（宋体小四、首行缩进）排版。
- 处理前建议先备份；工具本身提供了 `--backup` / 界面勾选备份。
- 自定义方案里若填了本机没有的字体，Word 会按回退字体显示，文档不会损坏。

---

## 八、许可证

MIT License —— 可自由使用、修改、分发，包括商用。
