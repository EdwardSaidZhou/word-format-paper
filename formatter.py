# -*- coding: utf-8 -*-
"""
formatter.py —— Word(.docx) 学术论文/公文格式批量调整核心库

设计目标
--------
1. 正文 / 标题：段前段后 0 间距，1.5 倍行距
2. 正文：小四(12pt)、首行缩进 2 字符、宋体(中文，含包夹中文的引号) + Times New Roman(数字/西文)
3. 脚注区：小五(9pt)、不缩进、脚注分割线也不缩进、宋体 + Times New Roman、单倍行距
4. 页码：页面底部居中、小五、Times New Roman
5. 图表标题：五号(10.5pt)、居中
6. 全文字色：黑色
7. 标题层级（自动识别，不写死 Heading1/2/3 的层级，见 detect_headings）
   - 主标题        : 黑体 小三 居中（无编号）
   - 一级标题      : 黑体 四号 居中（一、二、三、四、五……）
   - 二级标题      : 宋体 小四 加粗 首行缩进 2 字符（（一）（二）（三）……）
   - 三级标题      : 宋体 小四 首行缩进 2 字符（1. 2. 3. 4. 5. ……）
   若文档开头没有主标题（第一个非空段落就是正文），则自动去掉"主标题"这一层，
   其余层级依次上移，永远不会出现"全部错一位"的问题。

依赖：python-docx
"""

from __future__ import annotations

import re
from copy import deepcopy

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from profile import (FormatProfile, ParaSpec, parse_color,
                     DEFAULT_CAPTION_PREFIXES)

_DEFAULT_PROFILE = None


def get_default_profile():
    """默认排版方案（懒加载，避免 import 期开销）。"""
    global _DEFAULT_PROFILE
    if _DEFAULT_PROFILE is None:
        _DEFAULT_PROFILE = FormatProfile.default()
    return _DEFAULT_PROFILE

# --------------------------------------------------------------------------- #
# 常量
# --------------------------------------------------------------------------- #

FONT_SONG = "宋体"
FONT_HEI = "黑体"
FONT_TNR = "Times New Roman"

# 中文字号 -> 磅值
PT_XIAOWU = 9.0      # 小五
PT_WUHAO = 10.5      # 五号
PT_XIAOSI = 12.0     # 小四
PT_SIHAO = 14.0      # 四号
PT_XIAOSAN = 15.0    # 小三

LINE_SINGLE = 1.0
LINE_ONE_HALF = 1.5

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# 需要单独指定中文字体的中文标点（Word 常把它们当西文字符渲染，必须单独处理）
CJK_PUNCT_CHARS = "“”‘’「」『』《》〈〉、。，；：！？——…"

# w:rPr 子元素顺序（OOXML schema 要求，顺序错会导致 Word 报"文档损坏"）
RPR_ORDER = [
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
    "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
    "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
    "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position",
    "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
    "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
    "w:eastAsianLayout", "w:specVanish", "w:oMath",
]

# w:pPr 子元素顺序
PPR_ORDER = [
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
]


# --------------------------------------------------------------------------- #
# 底层 XML 工具
# --------------------------------------------------------------------------- #

def _ordered_insert(parent, child, order):
    """按 OOXML schema 规定的顺序把 child 插入 parent（已存在则直接返回）。"""
    tag = child.tag
    for existing in parent:
        if existing.tag == tag:
            return existing
    try:
        idx = order.index(tag)
    except ValueError:
        parent.append(child)
        return child
    for i, existing in enumerate(parent):
        try:
            cur = order.index(existing.tag)
        except ValueError:
            continue
        if cur > idx:
            parent.insert(i, child)
            return child
    parent.append(child)
    return child


def _get_or_add(parent, tag, order):
    el = parent.find(qn(tag))
    if el is not None:
        return el
    el = OxmlElement(tag)
    return _ordered_insert(parent, el, order)


def _remove(parent, tag):
    el = parent.find(qn(tag))
    if el is not None:
        parent.remove(el)


def _set_bool(rPr, tag, value: bool):
    """value=True 时加 <w:b/>，False 时移除（或置 0）。"""
    if value:
        el = _get_or_add(rPr, tag, RPR_ORDER)
        el.set(qn("w:val"), "1")
    else:
        _remove(rPr, tag)


# --------------------------------------------------------------------------- #
# Run 级格式
# --------------------------------------------------------------------------- #

def set_run_format(run_el,
                   east_asia=FONT_SONG,
                   west=FONT_TNR,
                   size_pt=PT_XIAOSI,
                   bold=None,
                   italic=None,
                   color="000000"):
    """设置一个 <w:r> 元素的字体属性。"""
    rPr = run_el.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        run_el.insert(0, rPr)

    rFonts = _get_or_add(rPr, "w:rFonts", RPR_ORDER)
    rFonts.set(qn("w:ascii"), west)
    rFonts.set(qn("w:hAnsi"), west)
    rFonts.set(qn("w:eastAsia"), east_asia)
    rFonts.set(qn("w:cs"), west)
    rFonts.set(qn("w:hint"), "eastAsia")   # 关键：让 Word 用 eastAsia 字体渲染东亚字符

    if bold is not None:
        _set_bool(rPr, "w:b", bold)
        _set_bool(rPr, "w:bCs", bold)
    if italic is not None:
        _set_bool(rPr, "w:i", italic)
        _set_bool(rPr, "w:iCs", italic)

    color_el = _get_or_add(rPr, "w:color", RPR_ORDER)
    color_el.set(qn("w:val"), color)

    half = str(int(round(size_pt * 2)))
    sz = _get_or_add(rPr, "w:sz", RPR_ORDER)
    sz.set(qn("w:val"), half)
    szCs = _get_or_add(rPr, "w:szCs", RPR_ORDER)
    szCs.set(qn("w:val"), half)
    return run_el


def set_run_black(run_el):
    """只把颜色刷成黑色，不动其它属性（用于"全文字色黑色"的全量兜底）。"""
    rPr = run_el.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        run_el.insert(0, rPr)
    color_el = _get_or_add(rPr, "w:color", RPR_ORDER)
    color_el.set(qn("w:val"), "000000")
    _remove(rPr, "w:highlight")


def split_cjk_punctuation(paragraph_el, east_asia=FONT_SONG, size_pt=PT_XIAOSI,
                          chars=None):
    """
    把段落里"包夹中文的引号"等中文标点拆成独立的 run，
    并将其 ascii/hAnsi/eastAsia 全部设为中文字体，
    避免出现「中文用宋体、引号却是 Times New Roman」的割裂感。
    """
    punct = chars if chars is not None else CJK_PUNCT_CHARS
    runs = [r for r in paragraph_el.iter(qn("w:r"))]
    for r in runs:
        t_nodes = r.findall(qn("w:t"))
        if not t_nodes:
            continue
        text = "".join(t.text or "" for t in t_nodes)
        if not any(ch in punct for ch in text):
            continue

        # 拆成 (片段, 是否中文标点) 序列
        segments = []
        buf = []
        for ch in text:
            if ch in punct:
                if buf:
                    segments.append(("".join(buf), False))
                    buf = []
                segments.append((ch, True))
            else:
                buf.append(ch)
        if buf:
            segments.append(("".join(buf), False))
        if not segments:
            continue

        parent = r.getparent()
        if parent is None:
            continue
        idx = list(parent).index(r)

        new_runs = []
        for seg_text, is_punct in segments:
            nr = deepcopy(r)
            for t in nr.findall(qn("w:t")):
                nr.remove(t)
            t = OxmlElement("w:t")
            t.set(XML_SPACE, "preserve")
            t.text = seg_text
            nr.append(t)
            if is_punct:
                set_run_format(nr, east_asia=east_asia, west=east_asia,
                               size_pt=size_pt, color="000000")
            new_runs.append(nr)

        parent.remove(r)
        for i, nr in enumerate(new_runs):
            parent.insert(idx + i, nr)


# --------------------------------------------------------------------------- #
# 段落级格式
# --------------------------------------------------------------------------- #

def set_paragraph_format(paragraph_el,
                         align=None,
                         line=360,
                         line_rule="auto",
                         before_pt=0,
                         after_pt=0,
                         first_line_chars=None,
                         size_pt=PT_XIAOSI,
                         keep_indent=False,
                         left_pt=None,
                         hanging_chars=None):
    """
    段间距 / 行距 / 对齐 / 缩进。
      line, line_rule : w:line 与 w:lineRule（240=单倍，360=1.5 倍；rule 可为 auto/exact/atLeast）
      before_pt/after_pt : 段前段后，单位磅
      left_pt        : 左缩进（磅），用于脚注分割线等
      hanging_chars  : 悬挂缩进字符数
    """
    pPr = paragraph_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        paragraph_el.insert(0, pPr)

    spacing = _get_or_add(pPr, "w:spacing", PPR_ORDER)
    spacing.set(qn("w:before"), str(int(round(float(before_pt or 0) * 20))))
    spacing.set(qn("w:after"), str(int(round(float(after_pt or 0) * 20))))
    spacing.set(qn("w:line"), str(int(line)))
    spacing.set(qn("w:lineRule"), line_rule or "auto")
    _remove(spacing, "w:beforeAutospacing")
    _remove(spacing, "w:afterAutospacing")

    if left_pt is not None:
        ind = _get_or_add(pPr, "w:ind", PPR_ORDER)
        ind.set(qn("w:left"), str(int(round(float(left_pt) * 20))))

    ind = pPr.find(qn("w:ind"))
    if first_line_chars:
        ind = _get_or_add(pPr, "w:ind", PPR_ORDER)
        ind.set(qn("w:firstLineChars"), str(int(first_line_chars * 100)))
        ind.set(qn("w:firstLine"), str(int(first_line_chars * size_pt * 20)))
    elif not keep_indent and ind is not None:
        ind.attrib.pop(qn("w:firstLine"), None)
        ind.attrib.pop(qn("w:firstLineChars"), None)
        ind.attrib.pop(qn("w:hanging"), None)
        ind.attrib.pop(qn("w:hangingChars"), None)

    if hanging_chars:
        ind = _get_or_add(pPr, "w:ind", PPR_ORDER)
        ind.set(qn("w:hangingChars"), str(int(hanging_chars * 100)))
        ind.set(qn("w:hanging"), str(int(hanging_chars * size_pt * 20)))

    if align:
        jc = _get_or_add(pPr, "w:jc", PPR_ORDER)
        jc.set(qn("w:val"), align)


def format_paragraph_runs(paragraph_el,
                          east_asia=FONT_SONG,
                          west=FONT_TNR,
                          size_pt=PT_XIAOSI,
                          bold=None,
                          black_only=False,
                          italic=None,
                          color="000000"):
    """把一个段落（含超链接/修订标记里的 run）的所有文字统一格式。"""
    for r in paragraph_el.iter(qn("w:r")):
        if r.find(qn("w:fldChar")) is not None or r.find(qn("w:instrText")) is not None:
            continue
        if black_only:
            set_run_black(r)
        else:
            set_run_format(r, east_asia=east_asia, west=west,
                           size_pt=size_pt, bold=bold, italic=italic, color=color)


def apply_para_spec(p_el, spec: ParaSpec, keep_indent=False, force_align=None):
    """把一份 ParaSpec 完整应用到段落（段落属性 + 全部 run 属性）。"""
    set_paragraph_format(
        p_el,
        align=(force_align if force_align is not None else spec.align),
        line=spec.line,
        line_rule=spec.line_rule,
        before_pt=spec.space_before,
        after_pt=spec.space_after,
        first_line_chars=spec.first_line_indent_chars,
        size_pt=spec.size_pt,
        keep_indent=keep_indent,
    )
    format_paragraph_runs(p_el,
                          east_asia=spec.font.east_asia,
                          west=spec.font.west,
                          size_pt=spec.size_pt,
                          bold=spec.bold,
                          italic=spec.italic,
                          color=spec.color)


# --------------------------------------------------------------------------- #
# 引号智能归一化：引号内是西文时改用半角引号
# --------------------------------------------------------------------------- #

FQ_LD = "\u201c"   # “
FQ_RD = "\u201d"   # ”
FQ_LS = "\u2018"   # ‘
FQ_RS = "\u2019"   # ’
SQ_D = '"'          # 半角直双引号
SQ_S = "'"          # 半角直单引号
FULL_QUOTES = (FQ_LD, FQ_RD, FQ_LS, FQ_RS)

# CJK / 全角符号：命中就说明这段不是纯西文
CJK_RE = re.compile(
    "[\u2e80-\u2eff\u3000-\u303f\u3040-\u30ff\u3400-\u4dbf"
    "\u4e00-\u9fff\uf900-\ufaff\ufe30-\ufe4f\uff00-\uffef]")
LATIN_RE = re.compile(r"[A-Za-z]")


def find_quote_pairs(text):
    """
    找出所有配对的引号。返回 [(left, right, is_double, inner)]，按 left 排序。
    只认全角引号（半角 " ' 无法判断开合，且本来就是目标形态）。
    """
    pairs = []
    stack_d, stack_s = [], []
    for i, ch in enumerate(text):
        if ch == FQ_LD:
            stack_d.append(i)
        elif ch == FQ_RD:
            if stack_d:
                pairs.append((stack_d.pop(), i, True, ""))
        elif ch == FQ_LS:
            stack_s.append(i)
        elif ch == FQ_RS:
            if stack_s:
                pairs.append((stack_s.pop(), i, False, ""))
    pairs.sort(key=lambda x: x[0])
    return [(l, r, d, text[l + 1:r]) for l, r, d, _ in pairs]


def _looks_opening_quote(text, i):
    """
    判断半角直引号（" 或 '）在位置 i 是开引号还是闭引号。
    半角引号没有方向，只能靠前后文推断。
    """
    prev = text[i - 1] if i > 0 else ""
    nxt = text[i + 1] if i + 1 < len(text) else ""
    closing_after = (not nxt) or nxt.isspace() or nxt in ".,;:!?)]}）】》”’"

    if not prev:                                   # 行首
        return not closing_after
    if prev.isspace() or prev in "([{（【《“‘":    # 前面是空白或开括号
        return not closing_after
    if prev in ".,;:!?)]}）】》”’":                # 前面是标点
        return not closing_after
    if prev.isdigit():                             # 5' 6" 这类单位，不参与
        return False
    return False                                   # 前面是字母 -> 闭引号


def find_mixed_quote_pairs(text):
    """
    同时识别全角与半角引号并配对。返回 [(left, right, is_double, inner)]。
    全角引号自带方向；半角引号用 _looks_opening_quote 推断。
    """
    pairs = []
    stack_d, stack_s = [], []
    for i, ch in enumerate(text):
        if ch == FQ_LD:
            stack_d.append(i)
        elif ch == FQ_RD:
            if stack_d:
                pairs.append((stack_d.pop(), i, True))
        elif ch == FQ_LS:
            stack_s.append(i)
        elif ch == FQ_RS:
            if stack_s:
                pairs.append((stack_s.pop(), i, False))
        elif ch in (SQ_D, SQ_S):
            is_double = (ch == SQ_D)
            stack = stack_d if is_double else stack_s
            if _looks_opening_quote(text, i):
                stack.append(i)
            elif stack:
                pairs.append((stack.pop(), i, is_double))
    pairs.sort(key=lambda x: x[0])
    return [(l, r, d, text[l + 1:r]) for l, r, d in pairs]


def plan_smart_conversion(text, require_latin=True, convert_straight=True):
    """
    智能引号（smart quotes）规划。

    西文引文的引号统一成英文印刷体弯引号 “ ” / ‘ ’，
    并用西文字体渲染（视觉上就是半角）。

      · 原本是全角 “ ” -> 字符不变，只改字体
      · 原本是直引号 "  -> 转成 “ ”（convert_straight=True 时）

    返回 (char_map, font_positions)
      char_map       : {位置: 新字符}
      font_positions : 需要用西文字体单独渲染的位置集合
    """
    pairs = find_mixed_quote_pairs(text)
    char_map, font_pos = {}, set()
    stack = []          # [(left, right, converted)]
    for l, r, is_d, inner in pairs:
        while stack and stack[-1][1] < l:
            stack.pop()
        parent_ok = stack[-1][2] if stack else True
        if _is_western(inner, require_latin) and parent_ok:
            # 保留原文的双/单引号层级，只把直引号"扳弯"，不重排嵌套
            lc, rc = (FQ_LD, FQ_RD) if is_d else (FQ_LS, FQ_RS)
            if text[l] != lc:
                if convert_straight or text[l] not in (SQ_D, SQ_S):
                    char_map[l] = lc
            if text[r] != rc:
                if convert_straight or text[r] not in (SQ_D, SQ_S):
                    char_map[r] = rc
            font_pos.add(l)
            font_pos.add(r)
            converted = True
        else:
            converted = False
        stack.append((l, r, converted))
    return char_map, font_pos


def _is_western(inner, require_latin=True):
    """引号里的内容算不算"西文"：不含中日韩/全角字符。"""
    if not inner.strip():
        return False
    if CJK_RE.search(inner):
        return False
    if require_latin and not LATIN_RE.search(inner):
        return False
    return True


def plan_quote_conversion(text, require_latin=True):
    """
    返回需要转成半角的引号位置集合。
    嵌套规则：内层只有在"外层也判定为西文"时才转换，
    避免出现 “中文'英文'中文” 这种全角套半角的怪组合。
    """
    pairs = find_quote_pairs(text)
    positions = set()
    stack = []          # [(left, right, converted)]
    for l, r, is_d, inner in pairs:
        while stack and stack[-1][1] < l:
            stack.pop()
        parent_ok = stack[-1][2] if stack else True
        do = _is_western(inner, require_latin) and parent_ok
        if do:
            positions.add(l)
            positions.add(r)
        stack.append((l, r, do))
    return positions


def _collect_text_nodes(p_el):
    """收集段落内所有 w:t，返回 (全文, [(t节点, 起始, 结束)])。"""
    nodes, buf, pos = [], [], 0
    for t in p_el.iter(qn("w:t")):
        txt = t.text or ""
        nodes.append((t, pos, pos + len(txt)))
        buf.append(txt)
        pos += len(txt)
    return "".join(buf), nodes


def _replace_chars(nodes, char_map):
    """
    按全局位置替换字符（1:1 替换，不改变文本长度）。
    返回受影响的 run 列表，供调用方修字体。
    """
    touched = []
    for t, start, end in nodes:
        old = t.text or ""
        new = "".join(char_map.get(start + i, ch) for i, ch in enumerate(old))
        if new != old:
            t.text = new
            t.set(XML_SPACE, "preserve")
            run = t.getparent()
            if run is not None and run.tag == qn("w:r"):
                touched.append(run)
    return touched


def _set_west_font_keep_cjk(run, west_font, size_pt, color="000000"):
    """
    只把西文字体（ascii/hAnsi/cs）改掉，中文字体 eastAsia 保持原样。
    用于"全角引号 -> 半角引号"之后：字符已是 ASCII，必须按西文字体渲染。
    """
    rPr = run.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        run.insert(0, rPr)
    rf = _get_or_add(rPr, "w:rFonts", RPR_ORDER)
    east = rf.get(qn("w:eastAsia")) or FONT_SONG
    rf.set(qn("w:ascii"), west_font)
    rf.set(qn("w:hAnsi"), west_font)
    rf.set(qn("w:cs"), west_font)
    rf.set(qn("w:eastAsia"), east)
    rf.set(qn("w:hint"), "eastAsia")
    color_el = _get_or_add(rPr, "w:color", RPR_ORDER)
    color_el.set(qn("w:val"), color)
    half = str(int(round(size_pt * 2)))
    for tag in ("w:sz", "w:szCs"):
        el = _get_or_add(rPr, tag, RPR_ORDER)
        el.set(qn("w:val"), half)


def _isolate_and_refont(p_el, positions, west_font, size_pt, color="000000"):
    """把指定位置的字符拆成独立 run，并把中西文字体都设为西文字体。"""
    if not positions:
        return
    _full, nodes = _collect_text_nodes(p_el)
    for t, start, end in list(nodes):
        old = t.text or ""
        if not old:
            continue
        hits = {i for i in range(start, end) if i in positions}
        if not hits:
            continue
        run = t.getparent()
        if run is None or run.tag != qn("w:r"):
            continue
        if len(hits) == len(old):
            # 整个文本节点都是目标字符，直接改 run 字体
            set_run_format(run, east_asia=west_font, west=west_font,
                           size_pt=size_pt, color=color)
            continue
        # 拆成若干 run
        parent = run.getparent()
        if parent is None:
            continue
        idx = list(parent).index(run)
        segs = []
        for i, ch in enumerate(old):
            hit = (start + i) in hits
            if segs and segs[-1][1] == hit:
                segs[-1][0] += ch
            else:
                segs.append([ch, hit])
        parent.remove(run)
        for k, (seg_text, is_hit) in enumerate(segs):
            nr = deepcopy(run)
            for old_t in nr.findall(qn("w:t")):
                nr.remove(old_t)
            nt = OxmlElement("w:t")
            nt.set(XML_SPACE, "preserve")
            nt.text = seg_text
            nr.append(nt)
            if is_hit:
                set_run_format(nr, east_asia=west_font, west=west_font,
                               size_pt=size_pt, color=color)
            parent.insert(idx + k, nr)


def normalize_western_quotes(p_el, profile=None, scope="body"):
    """
    引号内是西文时，把全角引号转成半角。

      style="straight" : “Concept” -> "Concept"   （真正的半角字符）
      style="curly"    : 保留弯引号字符，但改用西文字体渲染（视觉上也是半角）

    必须在 split_cjk_punctuation 之后调用，否则会被中文字体覆盖回去。
    """
    if profile is None:
        profile = get_default_profile()
    q = profile.quotes
    if not q.normalize_western or q.scope == "none":
        return 0
    if q.scope not in ("all", scope):
        return 0

    text, nodes = _collect_text_nodes(p_el)
    # 注意：不能只在"有全角引号"时才处理。
    # 原文已经是半角直引号的段落（如英文文献条目）同样需要按西文字体渲染，
    # 否则 Word 会用中文字体把它画成全角。
    if not any(ch in FULL_QUOTES or ch in (SQ_D, SQ_S) for ch in text):
        return 0

    if q.style in ("smart", "curly"):
        # 英文印刷体弯引号：字符已是 “ ” 就只改字体；是直引号 " 则先扳弯
        char_map, font_pos = plan_smart_conversion(
            text, q.require_latin, convert_straight=(q.style == "smart"))
        if not char_map and not font_pos:
            return 0
        if char_map:
            _replace_chars(nodes, char_map)
        _isolate_and_refont(p_el, font_pos,
                            west_font=_quote_west_font(profile, scope),
                            size_pt=_quote_size(profile, scope),
                            color=parse_color(profile.color, "000000"))
        return len(font_pos)

    if q.style == "keep":
        return 0

    # straight：替换成真正的半角字符
    positions = plan_quote_conversion(text, q.require_latin)
    if not positions:
        return 0
    char_map = {}
    for i in positions:
        ch = text[i]
        if ch in (FQ_LD, FQ_RD):
            char_map[i] = SQ_D
        elif ch in (FQ_LS, FQ_RS):
            char_map[i] = SQ_S
    if not char_map:
        return 0
    touched = _replace_chars(nodes, char_map)
    # 字符已变成 ASCII 引号，必须按西文字体渲染，否则会显示成宋体的半角引号
    for run in touched:
        _set_west_font_keep_cjk(run,
                                west_font=_quote_west_font(profile, scope),
                                size_pt=_quote_size(profile, scope),
                                color=parse_color(profile.color, "000000"))
    return len(char_map)


def _quote_west_font(profile, scope):
    return (profile.footnote.font.west if scope == "footnote"
            else profile.body.font.west)


def _quote_size(profile, scope):
    return (profile.footnote.size_pt if scope == "footnote"
            else profile.body.size_pt)


# --------------------------------------------------------------------------- #
# 标题层级识别（本工具的核心：绝不写死 Heading1/2/3 -> 一级/二级/三级）
# --------------------------------------------------------------------------- #

NUM_LEVEL1 = re.compile(r"^\s*[一二三四五六七八九十百]+[、\.]")            # 一、 二、
NUM_LEVEL2 = re.compile(r"^\s*[（(]\s*[一二三四五六七八九十百]+\s*[)）]")  # （一） (二)
NUM_LEVEL3 = re.compile(r"^\s*\d{1,3}\s*[\.、]\s*\S")                     # 1. 2. 3.
NUM_LEVEL3B = re.compile(r"^\s*[（(]\s*\d{1,3}\s*[)）]")                  # （1） (2)

HEADING_STYLE_ID = re.compile(r"^heading\s*(\d+)$", re.I)          # Heading1
TITLE_STYLE_ID = re.compile(r"^title$", re.I)                      # Title
CN_HEADING_STYLE_ID = re.compile(r"^(\d+)$")                       # 中文版内置 "1"/"2"/"3"
CAPTION_STYLE_ID = re.compile(r"^caption$|^题注$", re.I)


def _style_outline_level(doc, style_id):
    """
    返回某 styleId 的大纲级别（1 基）。Title -> 0（比一级还浅，专指主标题）。
    查不到返回 None。
    """
    if not style_id:
        return None
    sid = str(style_id).strip()

    if TITLE_STYLE_ID.match(sid):
        return 0

    styles_el = doc.styles.element
    for style in styles_el.findall(qn("w:style")):
        if style.get(qn("w:styleId")) != sid:
            continue
        # 1) 优先读样式自带的 outlineLvl
        pPr = style.find(qn("w:pPr"))
        if pPr is not None:
            ol = pPr.find(qn("w:outlineLvl"))
            if ol is not None and ol.get(qn("w:val")) is not None:
                return int(ol.get(qn("w:val"))) + 1
        # 2) 其次读 basedOn 链
        based = style.find(qn("w:basedOn"))
        if based is not None:
            lv = _style_outline_level(doc, based.get(qn("w:val")))
            if lv is not None:
                return lv
        # 3) 最后看名字
        name_el = style.find(qn("w:name"))
        if name_el is not None and name_el.get(qn("w:val")):
            nm = name_el.get(qn("w:val"))
            m = re.match(r"^heading\s*(\d+)$", nm, re.I)
            if m:
                return int(m.group(1))
            if re.match(r"^title$", nm, re.I):
                return 0
            m = re.match(r"^标题\s*(\d+)$", nm)
            if m:
                return int(m.group(1))
        return None

    # 样式表里没有（可能是隐式样式），按 styleId 猜
    m = HEADING_STYLE_ID.match(sid)
    if m:
        return int(m.group(1))
    m = re.match(r"^标题\s*(\d+)$", sid)
    if m:
        return int(m.group(1))
    if CN_HEADING_STYLE_ID.match(sid):
        return int(sid)
    return None


_PATTERN_CACHE = {}


def _compiled(pattern):
    rx = _PATTERN_CACHE.get(pattern)
    if rx is None:
        rx = re.compile(pattern)
        _PATTERN_CACHE[pattern] = rx
    return rx


def _numbering_level(text, detection=None):
    """纯文本编号推断层级：返回 1/2/3 或 None。"""
    det = detection or get_default_profile().detection
    if _compiled(det.level2_pattern).match(text):
        return 2
    if _compiled(det.level1_pattern).match(text):
        return 1
    if _compiled(det.level3_pattern).match(text):
        return 3
    if _compiled(det.level3_alt_pattern).match(text):
        return 3
    return None


def _looks_like_heading_text(text, max_len=60):
    t = text.strip()
    if not t or len(t) > max_len:
        return False
    if t.endswith(("。", "；", "，", "：", ":", ";", "！", "？", "”", "）", ")")):
        return False
    return True


def iter_body_paragraphs(doc):
    """正文段落（含表格、文本框内的段落），文档顺序。"""
    return list(doc.element.body.iter(qn("w:p")))


def _in_table(p_el):
    parent = p_el.getparent()
    while parent is not None:
        if parent.tag == qn("w:tbl"):
            return True
        parent = parent.getparent()
    return False


def detect_headings(doc, detect_plain_title=None, profile=None):
    if profile is None:
        profile = get_default_profile()
    det = profile.detection
    if detect_plain_title is None:
        detect_plain_title = det.detect_plain_title
    """
    识别文档标题层级。返回 (heading_map, title_el)

      heading_map : {段落元素: level}，1/2/3 分别对应「一、」「（一）」「1.」三级
      title_el    : 文章主标题段落元素（没有则为 None）

    ── 为什么不会"错一位" ────────────────────────────────
    常见的错误实现是写死 "Heading1 -> 一级、Heading2 -> 二级、Heading3 -> 三级"。
    但很多文档的主标题本身就用 Heading1，"一、"用 Heading2，"（一）"用 Heading3，
    于是主标题被排成"一、"级、真正的"一、"被排成"（一）"级 —— 全局错一位。

    本实现的做法是"按文档中实际出现的层级重新对齐"，并且区分两条线索：
      · style_lv：样式/大纲级别（Title=0，Heading N = N）
      · num_lv  ：文本编号推断的语义级别（一、=1，（一）=2，1. =3）
    1) 主标题判定只看 style_lv：最浅的【样式】层级若只有一个、位于文首且不带编号，
       它就是主标题；
    2) 其余标题先按 style_lv 排序（有样式的），没样式但带编号的，用"已有标题的
       编号->样式对应关系"折算成等价级别，再统一从浅到深映射为 1/2/3 级。
    这样 Title+Heading1+Heading2、Heading1+Heading2+Heading3、Heading1 主标题 +
    纯编号正文标题、乃至完全无样式的纯文本文档，都能各自正确对齐。
    """
    candidates = []  # dict: order, p, style_lv, num_lv, text, by_style

    paragraphs = list(iter_body_paragraphs(doc))
    first_nonempty = None
    for p in paragraphs:
        if (p.find(qn("w:r")) is not None or p.find(qn("w:hyperlink")) is not None) \
                and _para_text(p).strip():
            first_nonempty = p
            break

    order = 0
    for p in paragraphs:
        text = _para_text(p).strip()
        if not text:
            continue
        if det.skip_table_paragraphs and _in_table(p):
            continue                       # 表格里的内容不参与标题识别
        if _is_caption_para(doc, p, text, profile):
            continue                       # 图表题注不是标题

        style_lv = None
        pPr = p.find(qn("w:pPr"))
        if pPr is not None:
            ol = pPr.find(qn("w:outlineLvl"))
            if ol is not None and ol.get(qn("w:val")) is not None:
                style_lv = int(ol.get(qn("w:val"))) + 1
            if style_lv is None:
                ps = pPr.find(qn("w:pStyle"))
                if ps is not None:
                    style_lv = _style_outline_level(doc, ps.get(qn("w:val")))

        num_lv = _numbering_level(text)

        if style_lv is None:
            # 没有标题样式：只认"短、独立、带编号"的段落，避免把正文长句当标题
            if num_lv is None or not _looks_like_heading_text(
                    text, max_len=det.max_plain_heading_length):
                continue
        candidates.append({"order": order, "p": p, "style_lv": style_lv,
                           "num_lv": num_lv, "text": text,
                           "by_style": style_lv is not None})
        order += 1

    if not candidates:
        return {}, None

    # ---------- 1) 主标题判定（只看样式层级） ----------
    styled = [c for c in candidates if c["by_style"]]
    title_el = None
    if styled and det.auto_detect_title:
        min_style = min(c["style_lv"] for c in styled)
        same_min = [c for c in styled if c["style_lv"] == min_style]
        first_styled = styled[0]
        if (first_styled["p"] is first_nonempty
                and first_styled["style_lv"] == min_style
                and first_styled["num_lv"] is None
                and len(first_styled["text"]) <= det.max_title_length
                and (len(same_min) == 1 or min_style == 0)):
            title_el = first_styled["p"]

    if (title_el is None and detect_plain_title and first_nonempty is not None
            and first_nonempty not in [c["p"] for c in candidates]
            and not _in_table(first_nonempty)):
        text0 = _para_text(first_nonempty).strip()
        if text0 and len(text0) <= det.max_title_length \
                and _numbering_level(text0, det) is None \
                and not text0.endswith(("。", "；", "，")) and candidates:
            title_el = first_nonempty

    rest = [c for c in candidates if c["p"] is not title_el]

    # ---------- 2) 折算统一排序键 ----------
    #   用已有"带编号的样式标题"建立 编号级别 -> 样式级别 的对应关系
    num2style = {}
    for c in rest:
        if c["by_style"] and c["num_lv"]:
            num2style.setdefault(c["num_lv"], c["style_lv"])
    base_style = min([c["style_lv"] for c in rest if c["by_style"]], default=None)

    for c in rest:
        if c["by_style"]:
            c["key"] = c["style_lv"]
        elif c["num_lv"] in num2style:
            c["key"] = num2style[c["num_lv"]]
        else:
            base = base_style if base_style is not None else 1
            c["key"] = base + c["num_lv"] - 1

    # ---------- 3) 从浅到深重映射为 1/2/3 ----------
    used = sorted({c["key"] for c in rest})
    mapping = {k: min(i + 1, 3) for i, k in enumerate(used)}

    heading_map = {c["p"]: mapping[c["key"]] for c in rest}
    return heading_map, title_el


def _para_text(p_el):
    parts = []
    for t in p_el.iter(qn("w:t")):
        parts.append(t.text or "")
    return "".join(parts)


def _is_caption_para(doc, p_el, text, profile=None):
    """判断是否为图/表题注（前缀可在配置里自定义）。"""
    if profile is None:
        profile = get_default_profile()
    cap = profile.caption
    pPr = p_el.find(qn("w:pPr"))
    if pPr is not None:
        ps = pPr.find(qn("w:pStyle"))
        if ps is not None and CAPTION_STYLE_ID.match(str(ps.get(qn("w:val")))):
            return True
    prefixes = [x for x in (cap.prefixes or DEFAULT_CAPTION_PREFIXES) if x]
    if not prefixes:
        return False
    alt = "|".join(re.escape(x) for x in prefixes)
    if re.match(r"^\s*(" + alt + r")\s*[\d\-．.、]", text):
        return len(text) <= cap.max_length
    return False


# --------------------------------------------------------------------------- #
# 样式应用
# --------------------------------------------------------------------------- #

def _finish_paragraph(p_el, spec, profile, scope="body"):
    """段落收尾：先处理中文标点字体，再做西文引号半角化（顺序不能反）。"""
    if profile.cjk_punct_enabled:
        split_cjk_punctuation(p_el, east_asia=spec.font.east_asia,
                              size_pt=spec.size_pt,
                              chars=profile.cjk_punct_chars)
    normalize_western_quotes(p_el, profile, scope=scope)


def apply_title(p_el, profile=None):
    """文章主标题（默认：黑体 小三 居中）。"""
    if profile is None:
        profile = get_default_profile()
    spec = profile.title
    apply_para_spec(p_el, spec)
    _finish_paragraph(p_el, spec, profile, scope="body")


def apply_heading_level(p_el, level, profile=None):
    """
    level 1 -> 一级标题（一、二、三…）
    level 2 -> 二级标题（（一）（二）…）
    level 3 -> 三级标题（1. 2. 3. …），更深的层级也按 3 处理
    """
    if profile is None:
        profile = get_default_profile()
    spec = {1: profile.heading1, 2: profile.heading2}.get(level, profile.heading3)
    apply_para_spec(p_el, spec)
    _finish_paragraph(p_el, spec, profile, scope="body")


def apply_body(p_el, keep_indent=False, profile=None):
    """正文（默认：小四、首行缩进 2 字符、宋体 + Times New Roman、1.5 倍行距）。"""
    if profile is None:
        profile = get_default_profile()
    spec = profile.body
    apply_para_spec(p_el, spec, keep_indent=keep_indent)
    _finish_paragraph(p_el, spec, profile, scope="body")


def apply_caption(p_el, profile=None):
    """图表题注（默认：五号 居中）。"""
    if profile is None:
        profile = get_default_profile()
    spec = profile.caption
    apply_para_spec(p_el, spec)
    _finish_paragraph(p_el, spec, profile, scope="body")


# --------------------------------------------------------------------------- #
# 脚注 / 尾注
# --------------------------------------------------------------------------- #

def _find_part(doc, suffix):
    """按 partname 后缀查找 XML part（自动跳过图片等无 element 的 part）。"""
    for part in doc.part.package.iter_parts():
        if str(part.partname).endswith(suffix) and hasattr(part, "element"):
            return part
    return None


def _get_or_create_style(doc, style_id, name, style_type="paragraph"):
    styles_el = doc.styles.element
    for style in styles_el.findall(qn("w:style")):
        if style.get(qn("w:styleId")) == style_id:
            return style
    style = OxmlElement("w:style")
    style.set(qn("w:type"), style_type)
    style.set(qn("w:styleId"), style_id)
    nm = OxmlElement("w:name")
    nm.set(qn("w:val"), name)
    style.append(nm)
    styles_el.append(style)
    return style


def _style_set_footnote_defaults(style_el, spec=None):
    """脚注/尾注文本样式：默认小五、不缩进、单倍行距、宋体 + TNR。"""
    if spec is None:
        spec = get_default_profile().footnote
    size_pt = spec.size_pt
    pPr = style_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        # w:pPr 必须在 w:name 之后、w:rPr 之前
        rPr = style_el.find(qn("w:rPr"))
        if rPr is not None:
            rPr.addprevious(pPr)
        else:
            style_el.append(pPr)
    spacing = _get_or_add(pPr, "w:spacing", PPR_ORDER)
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), str(int(spec.line)))
    spacing.set(qn("w:lineRule"), spec.line_rule)
    ind = _get_or_add(pPr, "w:ind", PPR_ORDER)
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:hanging"), "0")
    ind.set(qn("w:left"), str(int(round(float(spec.separator_indent or 0) * 20))))

    rPr = style_el.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        style_el.append(rPr)
    rFonts = _get_or_add(rPr, "w:rFonts", RPR_ORDER)
    rFonts.set(qn("w:ascii"), spec.font.west)
    rFonts.set(qn("w:hAnsi"), spec.font.west)
    rFonts.set(qn("w:eastAsia"), spec.font.east_asia)
    rFonts.set(qn("w:cs"), spec.font.west)
    rFonts.set(qn("w:hint"), "eastAsia")
    half = str(int(round(size_pt * 2)))
    for tag in ("w:sz", "w:szCs"):
        el = _get_or_add(rPr, tag, RPR_ORDER)
        el.set(qn("w:val"), half)
    color = _get_or_add(rPr, "w:color", RPR_ORDER)
    color.set(qn("w:val"), parse_color(spec.color, "000000"))


def format_footnote_styles(doc, profile=None):
    """脚注/尾注"样式"层面的排版（styles.xml），保存前调用。"""
    if profile is None:
        profile = get_default_profile()
    for sid, nm in (("FootnoteText", "footnote text"),
                    ("EndnoteText", "endnote text")):
        st = _get_or_create_style(doc, sid, nm)
        _style_set_footnote_defaults(st, profile.footnote)


def format_notes_xml(root, profile=None):
    """
    处理 w:footnotes / w:endnotes 根元素：
    脚注正文默认小五、宋体 + TNR、不缩进、单倍行距；分割线段落顶格。
    """
    if profile is None:
        profile = get_default_profile()
    spec = profile.footnote
    for fn in list(root.findall(qn("w:footnote"))) + list(root.findall(qn("w:endnote"))):
        try:
            fid_int = int(fn.get(qn("w:id")))
        except (TypeError, ValueError):
            fid_int = 1
        # 注意：w:separator 嵌在 w:p/w:r 下，必须递归查找
        is_separator = (fn.find(f".//{qn('w:separator')}") is not None
                        or fn.find(f".//{qn('w:continuationSeparator')}") is not None)
        for p in fn.findall(qn("w:p")):
            if is_separator or fid_int <= 0:
                _fix_separator_paragraph(p, profile)
            else:
                apply_para_spec(p, spec)
                if profile.cjk_punct_enabled:
                    split_cjk_punctuation(p, east_asia=spec.font.east_asia,
                                          size_pt=spec.size_pt,
                                          chars=profile.cjk_punct_chars)
                normalize_western_quotes(p, profile, scope="footnote")
                _strip_reference_runs(p, profile)
    for r in root.iter(qn("w:r")):
        set_run_black(r)
    return root


def postprocess_notes(path, profile=None):
    """
    python-docx 不认识 footnotes/endnotes part，保存后直接对 docx 压缩包里的
    word/footnotes.xml、word/endnotes.xml 做二次处理。
    """
    import zipfile
    from lxml import etree
    from docx.oxml.parser import parse_xml

    targets = ("word/footnotes.xml", "word/endnotes.xml")
    if not zipfile.is_zipfile(path):
        return False
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        items = {n: zin.read(n) for n in names}
    changed = False
    for name in targets:
        if name not in items:
            continue
        try:
            root = parse_xml(items[name])
            format_notes_xml(root, profile)
        except Exception:
            continue
        body = etree.tostring(root, encoding="UTF-8", xml_declaration=False)
        items[name] = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                       + body)
        changed = True
    if changed:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                zout.writestr(n, items[n])
    return changed


REF_PLACEHOLDER = "\u0001"        # 在文本流里代表一个脚注/尾注引用标记
SPACE_CHARS = " \u00a0\u3000"          # 半角空格 / 不换行空格 / 全角空格
REF_TAGS = (qn("w:footnoteReference"), qn("w:endnoteReference"))


def _iter_flow_tokens(p_el):
    """按文档顺序 yield (kind, 元素)："t" = w:t，"ref" = 脚注/尾注引用标记。"""
    for el in p_el.iter():
        tag = el.tag
        if tag == qn("w:t"):
            yield "t", el
        elif tag in REF_TAGS:
            yield "ref", el


def trim_spaces_around_refs(p_el):
    """
    去掉脚注序号与正文之间的空格。

    例：  "……意识形态。 40 41 ……"  ->  "……意识形态。40 41……"
         （序号 40 前的空格删掉，40 与 41 之间的空格保留，否则会变成 "4041"）

    只动空格，不动其它字符；跨 run 也能处理。
    """
    nodes = [(kind, el, (el.text or "") if kind == "t" else REF_PLACEHOLDER)
             for kind, el in _iter_flow_tokens(p_el)]
    if not nodes:
        return 0

    pos_map = {}
    chars = []
    for ni, (_k, _el, text) in enumerate(nodes):
        for k, ch in enumerate(text):
            pos_map[len(chars)] = (ni, k)
            chars.append(ch)
    full = "".join(chars)

    def is_ref(i):
        return 0 <= i < len(full) and full[i] == REF_PLACEHOLDER

    drop = set()
    # 按"连续空格串"整体处理，串的位置是 [s, e)
    i = 0
    while i < len(full):
        if full[i] not in SPACE_CHARS:
            i += 1
            continue
        s0 = i
        while i < len(full) and full[i] in SPACE_CHARS:
            i += 1
        e0 = i                       # 空格串 = full[s0:e0]
        left_ref = is_ref(s0 - 1)
        right_ref = is_ref(e0)
        if not (left_ref or right_ref):
            continue                 # 与序号无关的空格，一律不动
        if left_ref and right_ref:
            # 两个序号之间：保留 1 个空格，否则 "40 41" 会变成 "4041"
            drop.update(pos_map[k] for k in range(s0 + 1, e0))
        else:
            # 序号与正文之间：全部删掉
            drop.update(pos_map[k] for k in range(s0, e0))
    if not drop:
        return 0
    for ni, (kind, el, text) in enumerate(nodes):
        if kind != "t":
            continue
        new_text = "".join(ch for k, ch in enumerate(text)
                           if (ni, k) not in drop)
        if new_text != text:
            el.text = new_text
            el.set(XML_SPACE, "preserve")
    return len(drop)


def format_footnote_references(doc, profile=None):
    """
    正文中的脚注/尾注序号：默认小四（跟随正文字号）、上标、黑色，
    并去掉序号与正文之间的空格。
    """
    if profile is None:
        profile = get_default_profile()
    spec = profile.footnote

    if spec.reference_trim_spaces:
        for p in doc.element.body.iter(qn("w:p")):
            try:
                trim_spaces_around_refs(p)
            except Exception:      # pragma: no cover
                pass

    # 序号字号：没单独指定就跟随正文（默认小四）
    size_pt = spec.reference_size_pt or profile.body.size_pt
    for r in doc.element.body.iter(qn("w:r")):
        if r.find(qn("w:footnoteReference")) is not None \
                or r.find(qn("w:endnoteReference")) is not None:
            set_run_format(r, east_asia=spec.font.east_asia, west=spec.font.west,
                           size_pt=size_pt, bold=False,
                           color=parse_color(spec.color, "000000"))
            if spec.reference_superscript:
                rPr = r.find(qn("w:rPr"))
                if rPr is None:
                    rPr = OxmlElement("w:rPr")
                    r.insert(0, rPr)
                va = _get_or_add(rPr, "w:vertAlign", RPR_ORDER)
                va.set(qn("w:val"), "superscript")
            else:
                rPr = r.find(qn("w:rPr"))
                if rPr is not None:
                    _remove(rPr, "w:vertAlign")


def _strip_reference_runs(p_el, profile=None):
    """脚注段落里的引用标记：小五 + 上标，不按正文处理。"""
    if profile is None:
        profile = get_default_profile()
    spec = profile.footnote
    for r in p_el.iter(qn("w:r")):
        if r.find(qn("w:footnoteRef")) is not None \
                or r.find(qn("w:endnoteRef")) is not None:
            set_run_format(r, east_asia=spec.font.east_asia, west=spec.font.west,
                           size_pt=spec.size_pt, bold=False,
                           color=parse_color(spec.color, "000000"))
            if spec.reference_superscript:
                rPr = r.find(qn("w:rPr"))
                if rPr is None:
                    rPr = OxmlElement("w:rPr")
                    r.insert(0, rPr)
                va = _get_or_add(rPr, "w:vertAlign", RPR_ORDER)
                va.set(qn("w:val"), "superscript")
            else:
                rPr = r.find(qn("w:rPr"))
                if rPr is not None:
                    _remove(rPr, "w:vertAlign")


def _fix_separator_paragraph(p_el, profile=None):
    """脚注分割线所在段落：默认顶格（不缩进），缩进量可在配置里改。"""
    if profile is None:
        profile = get_default_profile()
    spec = profile.footnote
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)
    ind = _get_or_add(pPr, "w:ind", PPR_ORDER)
    for attr in ("w:firstLine", "w:firstLineChars", "w:hanging",
                 "w:hangingChars", "w:left", "w:leftChars", "w:right",
                 "w:rightChars"):
        ind.attrib.pop(qn(attr), None)
    ind.set(qn("w:left"), str(int(round(float(spec.separator_indent or 0) * 20))))
    ind.set(qn("w:firstLine"), "0")
    spacing = _get_or_add(pPr, "w:spacing", PPR_ORDER)
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), str(int(spec.line)))
    spacing.set(qn("w:lineRule"), spec.line_rule)


# --------------------------------------------------------------------------- #
# 页码
# --------------------------------------------------------------------------- #

def _iter_footer_parts(doc):
    for part in list(doc.part.package.iter_parts()):
        if "/word/footer" in str(part.partname).lower() and hasattr(part, "element"):
            yield part


def format_page_number(doc, add_if_missing=True, profile=None):
    """
    页码：默认页面底部居中、小五、Times New Roman。
    直接遍历 footer part，避免 python-docx 访问 section.footer 时"隐式创建空页脚"的副作用。
    """
    if profile is None:
        profile = get_default_profile()
    spec = profile.page_number
    page_found = False
    for part in _iter_footer_parts(doc):
        root = part.element
        for t in root.iter(qn("w:instrText")):
            if "PAGE" in (t.text or "").upper():
                page_found = True
        for p in root.iter(qn("w:p")):
            has_field = p.find(f".//{qn('w:instrText')}") is not None
            if not has_field and not _para_text(p).strip():
                continue
            apply_para_spec(p, spec, force_align=spec.align or "center")
            if has_field and spec.prefix:
                _prepend_text(p, spec.prefix, spec)
            if has_field and spec.suffix:
                _append_text(p, spec.suffix, spec)

    if not page_found and add_if_missing and spec.add_if_missing:
        try:
            target = (doc.sections[0].header if spec.position == "header"
                      else doc.sections[0].footer)
            target.is_linked_to_previous = False   # 确保第一节拥有独立页脚定义
            _insert_page_number(target, profile)
            page_found = True
        except Exception:
            pass
    return page_found


def _prepend_text(p_el, text, spec):
    """在页码域之前加前缀（如 "第 "）。"""
    first = p_el.find(qn("w:r"))
    if first is None:
        return
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    r.insert(0, rPr)
    t = OxmlElement("w:t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    set_run_format(r, east_asia=spec.font.east_asia, west=spec.font.west,
                   size_pt=spec.size_pt, bold=spec.bold,
                   color=parse_color(spec.color, "000000"))
    first.addprevious(r)


def _append_text(p_el, text, spec):
    """在页码域之后加后缀（如 " 页"）。"""
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    r.insert(0, rPr)
    t = OxmlElement("w:t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    set_run_format(r, east_asia=spec.font.east_asia, west=spec.font.west,
                   size_pt=spec.size_pt, bold=spec.bold,
                   color=parse_color(spec.color, "000000"))
    p_el.append(r)


def _insert_page_number(footer, profile=None):
    """在页脚/页眉插入页码域（默认居中、小五、Times New Roman）。"""
    if profile is None:
        profile = get_default_profile()
    spec = profile.page_number
    try:
        paragraphs = footer.paragraphs
        if paragraphs:
            p = paragraphs[0]._p
        else:
            p = OxmlElement("w:p")
            footer._element.append(p)
    except Exception:
        return
    # 清空
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)
    set_paragraph_format(p, align=(spec.align or "center"),
                         line=spec.line, line_rule=spec.line_rule,
                         before_pt=spec.space_before, after_pt=spec.space_after,
                         first_line_chars=None, size_pt=spec.size_pt)

    def _run(inner=None):
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        r.insert(0, rPr)
        set_run_format(r, east_asia=spec.font.east_asia, west=spec.font.west,
                       size_pt=spec.size_pt, bold=spec.bold,
                       color=parse_color(spec.color, "000000"))
        if inner is not None:
            r.append(inner)
        return r

    def _fld(t):
        f = OxmlElement("w:fldChar")
        f.set(qn("w:fldCharType"), t)
        return f

    def _instr(txt):
        i = OxmlElement("w:instrText")
        i.set(XML_SPACE, "preserve")
        i.text = txt
        return i

    if spec.prefix:
        _prepend_text(p, spec.prefix, spec)
    p.append(_run(_fld("begin")))
    p.append(_run(_instr(" PAGE ")))
    p.append(_run(_fld("separate")))
    t = OxmlElement("w:t")
    t.text = "1"
    p.append(_run(t))
    p.append(_run(_fld("end")))
    if spec.suffix:
        _append_text(p, spec.suffix, spec)


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #

def format_document(src_path, dst_path=None,
                    add_page_number=None,
                    format_caption=True,
                    detect_plain_title=None,
                    profile=None,
                    log=None):
    """
    按 profile 执行格式化。返回统计信息 dict。

      profile : FormatProfile 实例 / JSON 文件路径 / dict；None 则用默认方案
    """
    if log is None:
        def log(m):
            pass

    profile = resolve_profile(profile)
    if add_page_number is None:
        add_page_number = profile.page_number.add_if_missing
    if detect_plain_title is None:
        detect_plain_title = profile.detection.detect_plain_title

    doc = Document(src_path)
    stats = {"title": 0, "h1": 0, "h2": 0, "h3": 0, "body": 0,
             "caption": 0, "footnote": 0, "page_number": False,
             "profile": profile.name}

    heading_map, title_el = detect_headings(
        doc, detect_plain_title=detect_plain_title, profile=profile)

    body_spec = profile.body
    for p_el in iter_body_paragraphs(doc):
        text = _para_text(p_el).strip()
        if not text:
            # 空段落也统一行距/段间距，避免行距参差
            set_paragraph_format(p_el, line=body_spec.line,
                                 line_rule=body_spec.line_rule,
                                 before_pt=body_spec.space_before,
                                 after_pt=body_spec.space_after,
                                 first_line_chars=None, size_pt=body_spec.size_pt)
            continue

        if p_el is title_el:
            apply_title(p_el, profile)
            stats["title"] += 1
            continue

        if p_el in heading_map:
            lv = heading_map[p_el]
            apply_heading_level(p_el, lv, profile)
            stats[f"h{lv}"] += 1
            continue

        if format_caption and _is_caption_para(doc, p_el, text, profile):
            apply_caption(p_el, profile)
            stats["caption"] += 1
            continue

        # 自动编号/项目符号段落：保留原有缩进，不追加首行缩进
        pPr = p_el.find(qn("w:pPr"))
        keep = pPr is not None and pPr.find(qn("w:numPr")) is not None
        apply_body(p_el, keep_indent=keep, profile=profile)
        stats["body"] += 1

    # 表格单元格内的段落若没有文字 run，也要刷行距
    for tbl in doc.element.body.iter(qn("w:tbl")):
        for p in tbl.iter(qn("w:p")):
            if _para_text(p).strip() == "":
                set_paragraph_format(p, line=body_spec.line,
                                     line_rule=body_spec.line_rule,
                                     before_pt=body_spec.space_before,
                                     after_pt=body_spec.space_after,
                                     first_line_chars=None,
                                     size_pt=body_spec.size_pt)

    log(f"主标题 {stats['title']} 个；一级 {stats['h1']} 个；"
        f"二级 {stats['h2']} 个；三级 {stats['h3']} 个；"
        f"正文 {stats['body']} 段；图表题注 {stats['caption']} 个")

    # 脚注 / 尾注：样式定义 + 正文引用标记（保存前），脚注正文 XML（保存后）
    try:
        format_footnote_styles(doc, profile)
        format_footnote_references(doc, profile)
    except Exception as e:  # pragma: no cover
        log(f"[警告] 脚注样式处理跳过：{e}")

    # 页码
    try:
        stats["page_number"] = format_page_number(
            doc, add_if_missing=add_page_number, profile=profile)
    except Exception as e:  # pragma: no cover
        log(f"[警告] 页码处理跳过：{e}")

    # 兜底：全文字色（含页眉/页脚/脚注/表格/文本框）
    if profile.force_all_black:
        for r in doc.element.body.iter(qn("w:r")):
            set_run_black(r)
        for part in list(doc.part.package.iter_parts()):
            name = str(part.partname).lower()
            if ("/word/footer" in name or "/word/header" in name) \
                    and hasattr(part, "element"):
                for r in part.element.iter(qn("w:r")):
                    set_run_black(r)

    out = dst_path or src_path
    doc.save(out)

    try:
        if postprocess_notes(out, profile):
            stats["footnote"] = 1
    except Exception as e:  # pragma: no cover
        log(f"[警告] 脚注正文处理跳过：{e}")
    return stats


def resolve_profile(profile=None):
    """把 None / 路径 / 字典 / FormatProfile 统一成 FormatProfile 实例。"""
    if profile is None:
        return get_default_profile()
    if isinstance(profile, FormatProfile):
        return profile
    if isinstance(profile, dict):
        return FormatProfile.from_dict(profile)
    return FormatProfile.load(str(profile))


def preview_profile(profile=None):
    """返回人可读的方案摘要（多行），不改动任何文件。"""
    profile = resolve_profile(profile)
    lines = [f"方案：{profile.name}", ""]
    rows = [
        ("正文", profile.body),
        ("主标题", profile.title),
        ("一级标题（一、）", profile.heading1),
        ("二级标题（（一））", profile.heading2),
        ("三级标题（1.）", profile.heading3),
        ("图表题注", profile.caption),
        ("脚注", profile.footnote),
        ("页码", profile.page_number),
    ]
    for label, spec in rows:
        lines.append(f"  {label:<16}{_spec_desc(spec)}")
    d = profile.detection
    lines.append("")
    lines.append(f"  自动识别主标题      : {'是' if d.auto_detect_title else '否'}")
    lines.append(f"  首段无样式也当主标题: {'是' if d.detect_plain_title else '否'}")
    lines.append(f"  中文标点用中文字体  : {'是' if profile.cjk_punct_enabled else '否'}")
    lines.append(f"  全文字色            : #{profile.color}")
    q = profile.quotes
    style_txt = {"straight": '半角直引号 " \'', "curly": "半角弯引号（西文字体）",
                 "keep": "不转换"}.get(q.style, q.style)
    scope_txt = {"all": "正文+脚注", "footnote": "仅脚注",
                 "body": "仅正文", "none": "关闭"}.get(q.scope, q.scope)
    lines.append(f"  西文引文引号        : "
                 f"{'开' if q.normalize_western else '关'}（{style_txt}·{scope_txt}）")
    return "\n".join(lines)


def _spec_desc(spec):
    from profile import _size_name, _line_name
    align = {"left": "左对齐", "center": "居中", "right": "右对齐",
             "both": "两端对齐"}.get(spec.align, "—")
    bold = "加粗" if spec.bold else ("不加粗" if spec.bold is False else "保持")
    indent = (f"首行缩进{spec.first_line_indent_chars:g}字符"
              if spec.first_line_indent_chars else "不缩进")
    return (f"{_size_name(spec.size_pt)} / {spec.font.east_asia}+{spec.font.west} / "
            f"{_line_name(spec.line, spec.line_rule)} / {align} / {indent} / {bold}")
