"""INTERX 양식 PPT 를 도형으로 그리는 엔진 (인치)"""
from __future__ import annotations

import copy
import datetime
import itertools
import math
import os
import re

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from check_ban import slide_texts

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(os.path.dirname(HERE), "assets")
TEMPLATE = os.path.join(ASSETS, "interx_template.pptx")
ICONS = os.path.join(ASSETS, "icons")
LOGOS = os.path.join(ASSETS, "logos")
FONT = "Noto Sans KR"
LINE = 1.25                # 1.25 또는 1.5 만

# 기본 양식 색
NAVY = "002060"        # 플랫폼 머리 띠, Data Lake
NAVY_D = "0E2841"
PANEL = "252E3F"       # 보고형 패널 머리
SLATE9 = "28303A"      # 본문 짙은 글자
SLATE6 = "546680"      # DB, 내부 모듈 상자
SLATE3 = "D2DAE5"      # Edge 머리 띠
SLATE4 = "A6B2C2"      # 묶음 영역 테두리
SLATE0 = "F8F9FA"      # 묶음 영역 바탕
GRAY = "7F7F7F"        # 서버 상자, 기본 연결선, Prop 칩
GRAY_L = "BFBFBF"      # 패널 테두리
TBL_HEAD = "AEAEAE"    # 표 머리칸
BLUE = "0035D5"        # AAS 칩 (SM, SMC, SML)
BLUE_L = "A8BEE2"      # Ent 칩
GREEN = "13C690"       # EAC 표시, ISH 연결
ORANGE = "EA580B"      # 회사 주황: 목차 번호만
ACCENT = NAVY          # 강조: 번호 원, 이름표, 지금 단계, 핵심 행과 수치, 주 흐름
RED = "FF0000"
TEAL = "299288"        # 패널 소제목, 점선
WHITE = "FFFFFF"
BLACK = "000000"

CHIP_FILL = {"AAS": BLUE, "SM": BLUE, "SMC": BLUE, "SML": BLUE, "Ent": BLUE_L,
             "Prop": GRAY, "Rel": GRAY, "Ref": GRAY, "File": GRAY}

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


# --- 기본 글자
def _rgb(h):
    return RGBColor.from_string(h)


def _font(run, size, bold=False, color=BLACK):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = _rgb(color)
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rpr, qn(tag))
        el.set("typeface", FONT)


_BOLD_NAME = ("SemiBold", "ExtraBold", "Bold", "Black", "Heavy")
_KEEP_FONT = ("Wingdings", "Symbol", "Webdings")  # 기호 글꼴. 바꾸면 기호가 깨짐


def normalize_fonts(path):
    """모든 글꼴 이름을 FONT 로 바꾼다 (굵기 이름은 b=1)"""
    import re
    import shutil
    import tempfile
    import zipfile
    # presentation.xml 기본 글꼴도 바꿈
    part = re.compile(r"ppt/((slides|slideLayouts|slideMasters|notesMasters|notesSlides|handoutMasters|theme|charts)/[^/]+|presentation)\.xml$")
    tmp = tempfile.mktemp(suffix=".pptx", dir=os.path.dirname(os.path.abspath(path)))
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if part.match(info.filename):
                root = etree.fromstring(data)
                for el in root.iter():
                    if not isinstance(el.tag, str):
                        continue
                    name = el.get("typeface")
                    if name is None or name.startswith("+"):
                        continue
                    local = el.tag.split("}")[-1]
                    if local in ("sym", "buFont") and name.startswith(_KEEP_FONT):
                        continue
                    if name == MATH_FONT:   # 수식 글꼴은 그대로
                        continue
                    if name != FONT:
                        el.set("typeface", FONT)
                        for k in ("panose", "pitchFamily", "charset"):
                            el.attrib.pop(k, None)
                        par = el.getparent()
                        if (local in ("latin", "ea", "cs") and any(w in name for w in _BOLD_NAME)
                                and par is not None and par.tag.split("}")[-1] in ("rPr", "defRPr", "endParaRPr")):
                            par.set("b", "1")
                    elif name == "" or local == "font":
                        el.set("typeface", FONT)
                for el in list(root.iter(qn("a:p"))):   # 돌면서 넣으면 다음 요소를 건너뜀
                    ppr = el.find(qn("a:pPr"))
                    if ppr is None:
                        el.insert(0, etree.Element(qn("a:pPr")))
                for el in list(root.iter()):  # 템플릿 마스터의 latinLnBrk=1 을 끔
                    if isinstance(el.tag, str) and el.get("latinLnBrk") == "1":
                        el.set("latinLnBrk", "0")
                    if isinstance(el.tag, str) and el.tag == qn("a:pPr") and el.getparent().tag == qn("a:p"):
                        el.set("latinLnBrk", "0")
                    if isinstance(el.tag, str) and el.tag == qn("a:pPr") and info.filename.startswith("ppt/slides/"):
                        ls = el.find(qn("a:lnSpc"))
                        pct = ls.find(qn("a:spcPct")) if ls is not None else None
                        if pct is None or pct.get("val") not in ("125000", "150000"):
                            if ls is not None:
                                el.remove(ls)
                            ls = etree.Element(qn("a:lnSpc"))
                            etree.SubElement(ls, qn("a:spcPct")).set("val", "125000")
                            el.insert(0, ls)
                    if isinstance(el.tag, str) and el.tag == qn("a:bodyPr") and info.filename.startswith("ppt/slides/"):
                        sp = el.getparent().getparent() if el.getparent() is not None else None
                        nv = sp.find(qn("p:nvSpPr")) if sp is not None else None
                        cnv = nv.find(qn("p:cNvSpPr")) if nv is not None else None
                        if cnv is not None and cnv.get("txBox") == "1":
                            el.set("anchor", "ctr")
                    if isinstance(el.tag, str) and el.tag == qn("a:rPr") and el.get("lang") != "ko-KR":
                        # lang 이 ko-KR 아니면 낱말 중간 줄 바꿈
                        el.set("lang", "ko-KR")
                        el.set("altLang", "en-US")
                for el in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}ea",
                                    "{http://schemas.openxmlformats.org/drawingml/2006/main}cs"):
                    if el.get("typeface") == "":
                        el.set("typeface", FONT)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(info, data)
    shutil.move(tmp, path)


def _set_text(shape, text, size, bold=False, color=BLACK, align="c", anchor="m", inset=None):
    """text: 문자열, 줄 목록, (글, 색) 튜플"""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if inset is not None:
        l, t, r, b = inset
        tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = map(Inches, (l, t, r, b))
    lines = text if isinstance(text, list) else str(text).split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        txt, col = (line if isinstance(line, tuple) else (line, color))
        r = p.add_run()
        r.text = txt
        _font(r, size, bold, col)


def _fill(shape, color, alpha=None):
    if color is None:
        shape.fill.background()
        return
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(color)
    if alpha is not None:
        clr = shape.fill._xPr.find(qn("a:solidFill"))[0]
        etree.SubElement(clr, qn("a:alpha")).set("val", str(int(alpha * 100000)))


def _line(shape, color, width=0.5, dash=None):
    if color is None:
        shape.line.fill.background()
        return
    shape.line.color.rgb = _rgb(color)
    shape.line.width = Pt(width)
    if dash:
        ln = shape.line._get_or_add_ln()
        etree.SubElement(ln, qn("a:prstDash")).set("val", dash)


def _shadow(shape):
    sppr = shape._element.spPr
    eff = sppr.find(qn("a:effectLst"))
    if eff is None:
        eff = etree.SubElement(sppr, qn("a:effectLst"))
    sh = etree.SubElement(eff, qn("a:outerShdw"), blurRad="50800", dist="38100", dir="2700000", algn="tl", rotWithShape="0")
    c = etree.SubElement(sh, qn("a:prstClr"), val="black")
    etree.SubElement(c, qn("a:alpha"), val="40000")


def _no_style_effects(shape):
    """테마 기본 테두리, 그림자 제거"""
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def _adj(shape, *vals):
    for i, v in enumerate(vals):
        shape.adjustments[i] = v


def box(s, x, y, w, h, text="", fill=WHITE, line=None, lw=0.5, size=7.26, bold=False,
        color=BLACK, align="c", anchor="m", geom="rect", radius=None, shadow=False, inset=(0.04, 0.02, 0.04, 0.02),
        dash=None, alpha=None, name=None):
    geo = {"rect": MSO_SHAPE.RECTANGLE, "round": MSO_SHAPE.ROUNDED_RECTANGLE, "round2": MSO_SHAPE.ROUND_2_SAME_RECTANGLE,
           "oval": MSO_SHAPE.OVAL, "diamond": MSO_SHAPE.DIAMOND, "homePlate": MSO_SHAPE.PENTAGON,
           "chevron": MSO_SHAPE.CHEVRON}[geom]
    sh = s.shapes.add_shape(geo, Inches(x), Inches(y), Inches(w), Inches(h))
    _no_style_effects(sh)
    if radius is not None:
        _adj(sh, radius)
    _fill(sh, fill, alpha)
    _line(sh, line, lw, dash)
    if shadow:
        _shadow(sh)
    if name:
        sh.name = name
    _set_text(sh, text, size, bold, color, align, anchor, inset)
    return sh


def text(s, x, y, w, h, t, size=7.26, bold=False, color=BLACK, align="l", anchor="m", line=None):
    """글상자. 세로 중간 맞춤, line 1.25 또는 1.5"""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    _set_text(tb, t, size, bold, color, align, anchor, (0, 0, 0, 0))
    if line is not None:
        if line not in (1.25, 1.5):
            raise ValueError("줄 간격은 1.25 또는 1.5 만")
        for p_ in tb.text_frame.paragraphs:
            p_.line_spacing = line
    return tb


# --- 보고형 부품
def panel(s, x, y, w, h, title):
    """보고형 패널, 몸통 반환"""
    body = box(s, x, y + 0.29, w, h - 0.29, fill=WHITE, line=GRAY_L, geom="round", radius=0)
    box(s, x, y, w, 0.29, title, fill=PANEL, line=GRAY_L, geom="round2", size=12, bold=True, color=WHITE)
    return body


def subhead(s, x, y, w, t, color="299288"):
    """패널 안 작은 제목 (10.5pt, 양옆 선)"""
    text(s, x, y, w, 0.24, t, size=10.5, bold=True, color=color, align="c", anchor="m")


def rule_list(s, x, y, w, title, items, pitch=0.29, title_color="299288"):
    """아이콘 + 굵은 글 목록. 바닥 y 반환"""
    if any(len(it) > 2 for it in items):
        pitch = max(pitch, 0.38)
    h = 0.27 + pitch * len(items) + 0.06
    box(s, x, y + 0.12, w, h - 0.12, fill=None, line=GRAY_L, lw=0.5)
    tw = min(w - 0.4, 0.16 * len(title) + 0.3)
    box(s, x + (w - tw) / 2, y, tw, 0.24, title, fill=WHITE, size=10.5, bold=True, color=title_color,
        inset=(0, 0, 0, 0))
    _reg(s).setdefault("list_titles", []).append(title)
    used = _reg(s).setdefault("icon_uses", {})
    for i, it in enumerate(items):
        ic, t = it[0], it[1]
        used.setdefault(ic, set()).add(t)
        desc = it[2] if len(it) > 2 else None
        ry = y + 0.27 + i * pitch
        ih = 0.3 if desc else 0.22
        icon(s, ic, x + 0.2, ry + 0.03 if desc else ry, h=ih)
        box(s, x + 0.66, ry + 0.03, 0.007, (0.32 if desc else 0.19), fill=GRAY_L)
        text(s, x + 0.76, ry, w - 0.86, 0.22 if desc else 0.25, t, size=9, bold=True, color=PANEL, anchor="m")
        if desc:
            text(s, x + 0.76, ry + 0.2, w - 0.86, 0.18, desc, size=7.26, color=SLATE9, anchor="m")
    return y + h


NUM_D, NUM_PT = 0.18, 8.0            # 목록, 카드, 캡처
NUM_D_LINE, NUM_PT_LINE = 0.15, 7.26  # 구성도 선 옆


def num_badge(s, x, y, n, color=ACCENT, d=NUM_D, size=NUM_PT, line=None):
    """주황 번호 원"""
    return box(s, x, y, d, d, str(n), fill=color, line=line, lw=1.0 if line else 0.5, geom="oval", size=size,
               bold=True, color=WHITE, inset=(0, 0, 0, 0))


def _lh(size):
    """글 한 줄 높이"""
    return size * 1.2 * 1.25 / 72


def _nl(text_, w, size):
    """줄 수 어림 (마침표, 쉼표 줄 바꿈 포함)"""
    if not text_:
        return 0
    return sum(max(1, math.ceil(_text_w(l, size) / max(w, 0.1))) for l in ko_breaks(text_, w, size))


def _td_h(title, desc, w, tsize=9, dsize=7.26):
    """제목 + 설명 높이"""
    ds = desc if isinstance(desc, list) else ([desc] if desc else [])
    h = _nl(title, w, tsize) * _lh(tsize)
    if ds:
        h += 2 / 72 + sum(_nl(d, w, dsize) for d in ds) * _lh(dsize)
    return h


def title_desc(s, x, y, w, h, title, desc, tsize=9, dsize=7.26, tcolor=BLACK, dcolor=SLATE9):
    """제목과 설명을 한 덩어리로 칸 세로 가운데에"""
    th = _td_h(title, desc, w, tsize, dsize)
    top = y + (h - th) / 2
    tb = box(s, x, top, w, th, title, fill=None, size=tsize, bold=True, color=tcolor, align="l", anchor="t",
             inset=(0, 0, 0, 0))
    ds = desc if isinstance(desc, list) else ([desc] if desc else [])
    if ds:
        tb.text_frame.paragraphs[0].space_after = Pt(2)
        for d in ds:
            p_ = tb.text_frame.add_paragraph()
            r_ = p_.add_run()
            r_.text = d
            _font(r_, dsize, False, dcolor)
    return top


def _unit(s, kind, h):
    """같은 장, 같은 종류 상자는 가장 큰 높이로"""
    u = _reg(s).setdefault("unit", {})
    u[kind] = max(u.get(kind, 0.0), h)


def _unit_of(s, kind):
    v = _reg(s)["unit"][kind]
    _reg(s).setdefault("unit_drawn", {}).setdefault(kind, set()).add(round(v, 3))
    return v


def summary(s, t):
    """머리 띠 아래 요약 한 줄 (12pt)"""
    return text(s, 0.57, 1.35, 12.17, 0.32, t, size=12, color=BLACK)


def notes(s, x, y, w, lines, size=7.26, bold=False):
    """'※' 메모 줄"""
    return text(s, x, y, w, _notes_h(lines, w, size), ["※ " + l for l in lines], size=size, bold=bold)


def _notes_h(lines, w, size):
    n = sum(max(1, math.ceil(_text_w("※ " + l, size) / w)) for l in lines)
    return max(0.15, size * 1.3 / 72) * n + 0.05


def big_arrow(s, x, y, w=0.38, h=0.33, direction="right"):
    """패널 사이 회색 큰 화살표"""
    rot = {"right": 0, "down": 90, "left": 180, "up": 270}[direction]
    sh = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    _no_style_effects(sh)
    _adj(sh, 0.48473, 0.52219)
    sh.rotation = rot
    sppr = sh._element.spPr
    for el in sppr.findall(qn("a:solidFill")):
        sppr.remove(el)
    grad = etree.fromstring(
        f'<a:gradFill xmlns:a="{_A}" flip="none" rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:schemeClr val="bg1"><a:lumMod val="50000"/></a:schemeClr></a:gs>'
        '<a:gs pos="50000"><a:schemeClr val="bg1"><a:lumMod val="75000"/></a:schemeClr></a:gs>'
        '<a:gs pos="100000"><a:schemeClr val="bg1"><a:lumMod val="95000"/></a:schemeClr></a:gs>'
        '</a:gsLst><a:lin ang="10800000" scaled="1"/></a:gradFill>')
    geom = sppr.find(qn("a:prstGeom"))
    geom.addnext(grad)
    _line(sh, None)
    return sh


def table(s, x, y, col_w, rows, size=7.26, row_h=0.24, header_fill=TBL_HEAD, header_color=WHITE,
          highlight=(), align=None):
    """rows[0] 머리행. highlight: 주황 행 (0부터)"""
    nr, nc = len(rows), len(col_w)
    head_h = min(row_h, 0.27)
    gf = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(sum(col_w)), Inches(head_h + row_h * (nr - 1)))
    _obst(s, gf)
    tbl = gf.table
    tblpr = tbl._tbl.tblPr
    sid = tblpr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblpr, qn("a:tableStyleId"))
    sid.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"
    for j, w in enumerate(col_w):
        tbl.columns[j].width = Inches(w)
    for i in range(nr):
        tbl.rows[i].height = Inches(head_h if i == 0 else row_h)
        for j in range(nc):
            c = tbl.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.06)
            c.margin_top = c.margin_bottom = Inches(0.01)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            head = i == 0
            c.fill.solid()
            c.fill.fore_color.rgb = _rgb(header_fill if head else WHITE)
            col = header_color if head else (ACCENT if (i - 1) in highlight else BLACK)
            p = c.text_frame.paragraphs[0]
            a = (align[j] if align else ("c" if head else "l"))
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[a]
            r = p.add_run()
            r.text = str(rows[i][j])
            _font(r, size, bold=head or (i - 1) in highlight, color=col)
            tcpr = c._tc.get_or_add_tcPr()
            for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                ln = etree.SubElement(tcpr, qn(side), w="6350")
                sf = etree.SubElement(ln, qn("a:solidFill"))
                etree.SubElement(sf, qn("a:srgbClr"), val="D9D9D9" if not head else header_fill)
            # 테두리는 채움 앞에 와야 함
            fill_el = tcpr.find(qn("a:solidFill"))
            tcpr.remove(fill_el)
            tcpr.append(fill_el)
    return gf


# --- 구성도 부품
def frame(s, x, y, w, h, title, kind="platform", sub=None):
    """큰 묶음. kind: platform, edge, source"""
    _wall(s, (x, y, x + w, y + h))
    _reg(s)["obst"].append(((x, y, x + w, y + 0.27), 0))
    if kind == "platform":
        body = box(s, x, y + 0.27, w, h - 0.27, fill=WHITE, line=BLACK, lw=0.5, shadow=True)
        lines = [(title, WHITE)] + ([(sub, "FFC9A3")] if sub else [])
        box(s, x, y, w, 0.24, lines, fill=NAVY, size=7.26 if not sub else 6.35, bold=True, color=WHITE,
            inset=(0.02, 0, 0.02, 0))
    elif kind == "edge":
        body = box(s, x, y + 0.27, w, h - 0.27, fill=WHITE, line="E2E8F0", lw=1, geom="round", radius=0, shadow=True)
        lines = [(title, BLACK)] + ([(sub, ACCENT)] if sub else [])
        box(s, x, y, w, 0.24, lines, fill=SLATE3, size=6.35, bold=True, inset=(0.02, 0, 0.02, 0))
    else:
        body = box(s, x, y + 0.29, w, h - 0.29, fill=WHITE, line=BLACK, lw=0.5, shadow=True)
        box(s, x, y, w, 0.25, title, fill=GRAY, line=BLACK, geom="round", radius=0.06, size=8.16, bold=True,
            color=WHITE, shadow=True)
    return body


def server(s, x, y, w, h, t, fill=GRAY, size=8.16):
    """서버 상자 (회색)"""
    return _obst(s, box(s, x, y, w, h, t, fill=fill, size=size, bold=True, color=WHITE))


def db(s, x, y, w, h, t, size=7.26, line=None):
    """모듈, DB 상자 (남회색)"""
    return _obst(s, box(s, x, y, w, h, t, fill=SLATE6, line=line, lw=0.75, size=size, bold=True, color=WHITE))


def area(s, x, y, w, h, line=SLATE4):
    """서버 아래 모듈 묶음 영역"""
    a = box(s, x, y, w, h, fill=SLATE0, line=line, lw=0.25, geom="round", radius=0)
    _wall(s, a)
    return a


def zone(s, x, y, w, h, color=ACCENT, alpha=0.06, above=None):
    """옅은 색 구역, above 도형 바로 위에 배치"""
    z = box(s, x, y, w, h, fill=color, alpha=alpha)
    if above is None:
        send_to_back(s, z)
    else:
        above._element.addnext(z._element)
    return z


def focus(shape, color=ACCENT, width=1.5):
    """강조 상자에 주황 테두리"""
    _before_draw(shape)
    _line(shape, color, width)
    return shape


def more(s, x, y, color=GRAY):
    """'더 있음' 세로 줄임표"""
    return _obst(s, text(s, x - 0.1, y, 0.2, 0.22, "⋮", size=10, bold=True, color=color, align="c"))


def legend(s, x, y, items, pitch=0.17, kind="oval"):
    """종류 범례. items=[(색, 이름)]"""
    for i, (c, t) in enumerate(items):
        yy = y + i * pitch
        if kind == "oval":
            box(s, x, yy + 0.02, 0.1, 0.1, fill=None, line=c, lw=1.25, geom="oval")
        else:
            box(s, x, yy + 0.065, 0.16, 0.025, fill=c)
        text(s, x + 0.2, yy, _text_w(t, 6.35) + 0.1, 0.2, t, size=6.35, color=SLATE9, align="l")
    w = 0.2 + max(_text_w(t, 6.35) + 0.1 for _, t in items)
    _reg(s)["obst"].append(((x, y, x + w, y + pitch * (len(items) - 1) + 0.2), 0))


def tag(s, x, y, t="EAC", fill=GREEN):
    """연결선 위 작은 초록 표시 (EAC 등)"""
    t_ = box(s, x, y, 0.23, 0.16, t, fill=fill, size=5.44, bold=True, color=WHITE, inset=(0, 0, 0, 0))
    _reg(s)["obst"].append((_bb(t_), -1))  # 선 위 표시라 선 지남 검사에서 뺌
    return t_


def step(s, x, y, n, label=None, color=ACCENT, label_w=None, label_color=BLACK):
    """번호 원 (0.15 in) 과 오른쪽 설명"""
    c = num_badge(s, x, y, n, color, d=NUM_D_LINE, size=NUM_PT_LINE)
    if label_w is None:
        label_w = _text_w(label or "", 7.26) + 0.12
    if label:
        text(s, x + 0.19, y - 0.03, label_w, 0.22, label, size=7.26, bold=True, color=label_color, anchor="m")
        _reg(s).setdefault("rows", []).append((x, x + 0.19 + label_w, y + 0.075))
    _reg(s)["obst"].append(((x, y - 0.03, x + 0.19 + (label_w if label else 0), y + 0.19), 0))
    return c


def caption(s, x, y, w, t, color=RED, size=12):
    """도식 옆 빨간 설명 문구"""
    lines = t if isinstance(t, list) else [t]
    n = sum(max(1, math.ceil(_text_w(l, size) * 1.05 / w)) for l in lines)
    return text(s, x, y, w, n * size * 1.25 / 72 + 0.06, t, size=size, bold=True, color=color)


def label(s, x, y, w, t, size=6.35, color=SLATE9, bold=False, align="c"):
    """아이콘 아래 이름, 선 옆 작은 글자"""
    _reg(s).setdefault("rows", []).append((x, x + w, y + size * 1.2 / 72 / 2))
    return _obst(s, text(s, x, y, w, 0.2, t, size=size, bold=bold, color=color, align=align))


_BAD_ICONS = {"binary": "숫자가 빽빽한 그림이라 0.3 in 로 줄이면 뭉개짐. spreadsheet, doc_id_green, chart_distribution 등"}


def _icon_path(key):
    """assets/icons, assets/logos 에서 이름으로 찾음"""
    if key in _BAD_ICONS:
        raise ValueError(f"아이콘 '{key}' 는 쓰지 않음: {_BAD_ICONS[key]}")
    for d in (ICONS, LOGOS):
        hit = next((os.path.join(d, f) for f in os.listdir(d) if os.path.splitext(f)[0] == key), None)
        if hit:
            return hit
    raise ValueError(f"아이콘 '{key}' 없음: assets/icons/_catalog.png, assets/logos/_catalog.png 에서 고름")


def icon(s, key, x, y, w=None, h=None, name=None, lw=0.8):
    """아이콘 그림. name 을 주면 아래에 이름"""
    path = _icon_path(key)
    pic = s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w) if w else None, Inches(h) if h else None)
    if name:
        cx = x + pic.width / 914400 / 2
        label(s, cx - lw / 2, y + pic.height / 914400 + 0.03, lw, name)
    return _obst(s, pic)


# --- AAS 칩 트리
def chip(s, x, y, kind, value, value_w=1.46, h=0.17, value_line=GRAY, value_color=BLACK, size=8.0):
    """AAS 칩 + 값 상자. (칩, 값상자) 반환"""
    kw = 0.34 if len(kind) > 2 else 0.32
    k = box(s, x, y, kw, h, kind, fill=CHIP_FILL.get(kind, GRAY), size=size, bold=True, color=WHITE,
            geom="round", radius=0.14483, inset=(0, 0, 0, 0))
    v = box(s, x + kw + 0.03, y, value_w, h, value, fill=WHITE, line=value_line, lw=0.75, size=size, bold=True,
            color=value_color, align="l", geom="round", radius=0.06108, inset=(0.08, 0, 0.04, 0))
    return k, v


def chip_tree(s, x, y, rows, indent=0.31, pitch=0.255, value_w=1.46, size=8.0):
    """rows=[(깊이, 종류, 값)] 칩 트리"""
    stack, out = {}, []
    for i, (depth, kind, value) in enumerate(rows):
        vw = value_w(depth, kind) if callable(value_w) else value_w
        k, v = chip(s, x + depth * indent, y + i * pitch, kind, value, vw, size=size)
        if depth > 0 and depth - 1 in stack:
            connect(s, stack[depth - 1], "bottom", k, "left", color=GRAY, width=0.5, head=None)
        stack[depth] = k
        for d in [d for d in stack if d > depth]:
            del stack[d]
        out.append((k, v))
    return out


# --- 연결선
_SITES = {"rect": {"top": 0, "left": 1, "bottom": 2, "right": 3},
          "ellipse": {"top": 0, "left": 2, "bottom": 4, "right": 6}}


def _pt(shape, side):
    x, y, w, h = (v / 914400 for v in (shape.left, shape.top, shape.width, shape.height))
    return {"top": (x + w / 2, y), "bottom": (x + w / 2, y + h), "left": (x, y + h / 2), "right": (x + w, y + h / 2)}[side]


def _site(shape, side):
    g = shape._element.find(".//" + qn("a:prstGeom"))
    kind = "ellipse" if g is not None and g.get("prst") == "ellipse" else "rect"
    return _SITES[kind][side]


def connect(s, a, side_a, b, side_b, color=GRAY, width=1.5, head="end", dash=None, mid=0.5, text=None):
    """붙은 연결선. text 는 선 이름"""
    _before_draw(s)
    c = _connect(s, a, side_a, b, side_b, color, width, head, dash, mid)
    R = _reg(s)
    pts = _route(a, side_a, b, side_b, mid, check=False)
    R["segs"].extend(zip(pts, pts[1:]))
    R["links"].append((pts, a.shape_id, b.shape_id, f"{a.name} -> {b.name}"))
    for sh, sd in ((a, side_a), (b, side_b)):
        R["used"][(sh.shape_id, sd)] = R["used"].get((sh.shape_id, sd), 0) + 1
    if text:
        line_label(s, a, side_a, b, side_b, text, mid)
    return c


def line_label(s, a, side_a, b, side_b, t, mid=0.5, color=SLATE9):
    """선 모양 기준 선 이름 자리 선택"""
    (x1, y1), (x2, y2) = _pt(a, side_a), _pt(b, side_b)
    lw = _text_w(t, 6.35) + 0.08
    hz = ("left", "right")
    def above(cx, cy):
        return label(s, cx - lw / 2, cy - 0.19, lw, t, color=color, bold=True)
    def right_of(vx, cy):
        return label(s, vx + 0.05, cy - 0.09, lw, t, color=color, bold=True, align="l")
    if side_a in hz and side_b in hz:
        if abs(y1 - y2) < 0.05:
            return above((x1 + x2) / 2, y1)
        return right_of(x1 + (x2 - x1) * mid, (y1 + y2) / 2)
    if side_a not in hz and side_b not in hz:
        if abs(x1 - x2) < 0.05:
            return right_of(x1, (y1 + y2) / 2)
        return above((x1 + x2) / 2, y1 + (y2 - y1) * mid)
    vx, (ya, yb) = (x2, (y1, y2)) if side_a in hz else (x1, (y1, y2))
    return right_of(vx, (ya + yb) / 2)


def _connect(s, a, side_a, b, side_b, color=GRAY, width=1.5, head="end", dash=None, mid=0.5, htype="triangle"):
    """a.side_a 에서 b.side_b 로 붙는 연결선"""
    (x1, y1), (x2, y2) = _pt(a, side_a), _pt(b, side_b)
    ha, hb = side_a in ("left", "right"), side_b in ("left", "right")
    swap = False
    if ha != hb and not ha:  # 세로에서 가로로 가는 선은 역방향으로 그림
        a, side_a, b, side_b = b, side_b, a, side_a
        (x1, y1), (x2, y2) = (x2, y2), (x1, y1)
        ha, hb = True, False
        swap = True
    straight = abs(x1 - x2) < 0.005 or abs(y1 - y2) < 0.005
    if straight and ha == hb:
        prst, rot = "straightConnector1", 0
    elif ha and hb:
        prst, rot = "bentConnector3", 0
    elif not ha and not hb:
        prst, rot = "bentConnector3", 90
    else:
        prst, rot = "bentConnector2", 0
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(0), Inches(0), Inches(1), Inches(1))
    sp = c._element
    xfrm = sp.spPr.find(qn("a:xfrm"))
    for k in ("flipH", "flipV", "rot"):
        xfrm.attrib.pop(k, None)
    dx, dy = x2 - x1, y2 - y1
    if rot == 0:
        ox, oy, w, h = min(x1, x2), min(y1, y2), abs(dx), abs(dy)
        if dx < 0:
            xfrm.set("flipH", "1")
        if dy < 0:
            xfrm.set("flipV", "1")
    else:  # 90도 돌린 ㄷ 자: 가로 = 화면 세로
        w, h = abs(dy), abs(dx)
        cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
        ox, oy = cx - w / 2, cy - h / 2
        xfrm.set("rot", "5400000")
        if dx > 0:
            xfrm.set("flipV", "1")
        if dy < 0:
            xfrm.set("flipH", "1")
    xfrm.find(qn("a:off")).set("x", str(int(Inches(ox))))
    xfrm.find(qn("a:off")).set("y", str(int(Inches(oy))))
    xfrm.find(qn("a:ext")).set("cx", str(int(Inches(w))))
    xfrm.find(qn("a:ext")).set("cy", str(int(Inches(h))))
    geom = sp.spPr.find(qn("a:prstGeom"))
    geom.set("prst", prst)
    av = geom.find(qn("a:avLst"))
    if prst == "bentConnector3" and abs(mid - 0.5) > 1e-6:
        etree.SubElement(av, qn("a:gd"), name="adj1", fmla=f"val {int(mid * 100000)}")
    nv = sp.find(qn("p:nvCxnSpPr")).find(qn("p:cNvCxnSpPr"))
    etree.SubElement(nv, qn("a:stCxn"), id=str(a.shape_id), idx=str(_site(a, side_a)))
    etree.SubElement(nv, qn("a:endCxn"), id=str(b.shape_id), idx=str(_site(b, side_b)))
    st = sp.find(qn("p:style"))
    if st is not None:
        sp.remove(st)
    ln = sp.spPr.find(qn("a:ln"))
    if ln is None:
        ln = etree.SubElement(sp.spPr, qn("a:ln"))
    ln.set("w", str(int(Pt(width))))
    sf = etree.SubElement(ln, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr"), val=color)
    if dash:
        etree.SubElement(ln, qn("a:prstDash"), val=dash)
    if swap and head in ("end", "start"):
        head = "start" if head == "end" else "end"
    if head in ("start", "both"):
        etree.SubElement(ln, qn("a:headEnd"), type=htype)
    if head in ("end", "both"):
        etree.SubElement(ln, qn("a:tailEnd"), type=htype)
    return c


def send_to_back(s, shape):
    tree = s.shapes._spTree
    tree.remove(shape._element)
    tree.insert(2, shape._element)


# --- 자동 배치 (기본 양식에서 잰 틀)
FOOTER_Y = 7.215    # 꼬리말 선. 이 선이 바닥
FOOTER_GAP = 0.30   # 바닥 선 위 여백
REPORT_TOP, REPORT_BOTTOM = 1.86, round(FOOTER_Y - FOOTER_GAP, 3)
REPORT_LEFT, REPORT_RIGHT = 0.29, 13.09
REPORT_WIDTHS = (3.94, 3.21, 3.60)
PAD = 0.29                                   # 패널 안 좌우 여백
BODY_TOP, BODY_BOTTOM = 2.42, round(REPORT_BOTTOM - 0.2, 3)   # 패널 안 위아래 여백
DOC_TOP, DOC_BOTTOM = 2.21, min(6.81, REPORT_BOTTOM)


# --- 범용 장 부품
LINE_L = "D9D9D9"      # 표 칸선, 카드 테두리, 눈금선
ROW_BG = "F2F2F2"      # 목록 행 바탕
PALETTE = (NAVY, SLATE6, TEAL, GRAY, SLATE4)   # 차트 계열 색, 주황 없음


def mpl_style():
    """matplotlib 그림을 덱 색 및 글꼴로"""
    import matplotlib as mpl
    from cycler import cycler
    from matplotlib import font_manager as fm
    # 가변 글꼴만 있으면 가장 가는 굵기로 그려져 OS 한글 글꼴로
    ok = any(f.name == FONT and f.weight in (400, "normal", "regular") for f in fm.fontManager.ttflist)
    have = {f.name for f in fm.fontManager.ttflist}
    fam = [FONT] if ok else [n for n in ("Malgun Gothic", "Apple SD Gothic Neo", "AppleGothic") if n in have] or [FONT]
    mpl.rcParams.update({"axes.prop_cycle": cycler(color=["#" + c for c in PALETTE]), "font.family": fam,
                         "axes.unicode_minus": False, "axes.spines.top": False, "axes.spines.right": False,
                         "axes.edgecolor": "#" + GRAY, "axes.grid": True, "axes.axisbelow": True,
                         "grid.color": "#E5E7EB", "grid.linewidth": 0.6, "figure.dpi": 200})


def _default_mpl(path):
    """matplotlib 기본 주황, 파랑 칸이 그림의 0.5% 넘음"""
    from PIL import Image as _Im
    with _Im.open(path) as im:
        px = im.convert("RGB").resize((200, 200), _Im.NEAREST).getdata()
    hits = sum(1 for p in px if p in ((255, 127, 14), (31, 119, 180)))
    return hits > len(px) * 0.005


def card(s, x, y, w, h, title=None, title_color=BLUE):
    """흰 카드, 그림자, 파란 제목"""
    body = box(s, x, y, w, h, fill=WHITE, line=LINE_L, lw=0.75, shadow=True)
    _wall(s, body)
    if title:
        _obst(s, text(s, x + 0.1, y + 0.08, w - 0.2, 0.22, title, size=9, bold=True, color=title_color, align="c"))
    return body


def section(s, x, y, w, title, rows, row_h=0.3, highlight=(), size=9):
    """회색 머리 띠 + 행 목록. 바닥 y 반환"""
    box(s, x, y, w, 0.27, title, fill=TBL_HEAD, size=9, bold=True, color=WHITE, align="l", inset=(0.1, 0, 0.05, 0))
    for i, r in enumerate(rows):
        hl = i in highlight
        box(s, x, y + 0.27 + i * row_h + 0.02, w, row_h - 0.02, r, fill=ROW_BG, size=size, bold=hl,
            color=ACCENT if hl else BLACK, align="l", inset=(0.1, 0, 0.05, 0))
    return y + 0.27 + len(rows) * row_h


def _chart_font(ch):
    """차트 한글(ea) 글꼴까지 FONT 로"""
    for rpr in list(ch._chartSpace.iter(qn("a:defRPr"))) + list(ch._chartSpace.iter(qn("a:rPr"))):
        for tag_ in ("a:latin", "a:ea", "a:cs"):
            el = rpr.find(qn(tag_))
            if el is None:
                el = etree.SubElement(rpr, qn(tag_))
            el.set("typeface", FONT)


def chart(s, x, y, w, h, kind, categories, series, highlight=(), number_format="0", legend=None, colors=None,
          vmax=None, labels=True):
    """편집 가능한 차트. kind: col, bar, stack, line"""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_MARKER_STYLE
    types = {"col": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
             "stack": XL_CHART_TYPE.COLUMN_STACKED, "line": XL_CHART_TYPE.LINE_MARKERS}
    if kind not in types:
        raise ValueError(f"chart kind '{kind}' 없음: {list(types)}")
    if not categories or not series:
        raise ValueError("chart: 항목이나 계열이 없음")
    if kind == "bar":  # 가로 막대는 거꾸로 넣어야 첫 항목이 위
        n_ = len(categories)
        categories = list(reversed(categories))
        series = {k: list(reversed(v)) for k, v in series.items()}
        highlight = tuple(n_ - 1 - j for j in highlight)
    cd = CategoryChartData()
    cd.categories = categories
    for name, vals in series.items():
        cd.add_series(name, vals)
    gf = s.shapes.add_chart(types[kind], Inches(x), Inches(y), Inches(w), Inches(h), cd)
    ch = gf.chart
    ch.has_title = False
    ch.font.size = Pt(7.26)
    ch.font.name = FONT
    ch.font.color.rgb = _rgb(SLATE9)
    ch.has_legend = (len(series) > 1) if legend is None else legend
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(7.26)
    va, ca = ch.value_axis, ch.category_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = _rgb(LINE_L)
    va.major_gridlines.format.line.width = Pt(0.5)
    va.format.line.fill.background()
    va.tick_labels.number_format = number_format
    va.tick_labels.number_format_is_linked = False
    va.tick_labels.font.size = Pt(7.26)
    if kind != "line":
        va.minimum_scale = 0
    if vmax is not None:
        va.maximum_scale = vmax
    ca.format.line.color.rgb = _rgb(GRAY_L)
    ca.has_major_gridlines = False

    ca.tick_labels.font.size = Pt(7.26)
    plot = ch.plots[0]
    if kind in ("col", "bar", "stack"):
        plot.gap_width = 80
        if kind == "stack":
            plot.overlap = 100
    if labels and kind in ("col", "bar"):
        va.tick_label_position = __import__("pptx.enum.chart", fromlist=["XL_TICK_LABEL_POSITION"]).XL_TICK_LABEL_POSITION.NONE
        va.has_major_gridlines = False
    if labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(7.26)
        dl.font.bold = True
        dl.number_format = number_format
        dl.number_format_is_linked = False
        dl.position = (XL_LABEL_POSITION.CENTER if kind == "stack" else
                       XL_LABEL_POSITION.ABOVE if kind == "line" else XL_LABEL_POSITION.OUTSIDE_END)
    pal = colors or PALETTE
    for i, ser in enumerate(plot.series):
        c = pal[i % len(pal)]
        if highlight and len(series) == 1 and kind != "line":   # 강조 막대만 남색, 나머지 회색
            c = SLATE4
        if kind == "line":
            ser.format.line.color.rgb = _rgb(c)
            ser.format.line.width = Pt(1.5)
            ser.smooth = False
            ser.marker.style = XL_MARKER_STYLE.CIRCLE
            ser.marker.size = 5
            ser.marker.format.fill.solid()
            ser.marker.format.fill.fore_color.rgb = _rgb(c)
            ser.marker.format.line.color.rgb = _rgb(WHITE)
        else:
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = _rgb(c)
            if len(series) == 1:
                for j in highlight:
                    pt = ser.points[j]
                    pt.format.fill.solid()
                    pt.format.fill.fore_color.rgb = _rgb(ACCENT)
    _chart_font(ch)
    _obst(s, gf)
    return gf


def _axis_title(ax, t):
    ax.has_title = True
    ax.axis_title.text_frame.text = t
    for r in ax.axis_title.text_frame.paragraphs[0].runs:
        r.font.size, r.font.bold, r.font.name = Pt(7.26), False, FONT
        r.font.color.rgb = _rgb(SLATE9)


def scatter_chart(s, x, y, w, h, series, x_title, y_title, x_format="0.0", y_format="0.0"):
    """편집 가능한 산점도. series={이름: [(x, y), ...]}, 축 이름에 단위"""
    from pptx.chart.data import XyChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
    cd = XyChartData()
    for name, pts in series.items():
        ser = cd.add_series(name)
        for px_, py_ in pts:
            ser.add_data_point(px_, py_)
    gf = s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    ch = gf.chart
    ch.has_title = False
    ch.font.size, ch.font.name = Pt(7.26), FONT
    ch.has_legend = len(series) > 1
    if ch.has_legend:
        ch.legend.position, ch.legend.include_in_layout = XL_LEGEND_POSITION.BOTTOM, False
    for ax, t, f in ((ch.category_axis, x_title, x_format), (ch.value_axis, y_title, y_format)):
        ax.tick_labels.number_format, ax.tick_labels.number_format_is_linked = f, False
        ax.tick_labels.font.size = Pt(7.26)
        ax.format.line.color.rgb = _rgb(GRAY_L)
        ax.has_major_gridlines = ax is ch.value_axis
        _axis_title(ax, t)
    ch.value_axis.major_gridlines.format.line.color.rgb = _rgb(LINE_L)
    from pptx.enum.chart import XL_AXIS_CROSSES
    xs = [p_[0] for pts in series.values() for p_ in pts]
    ch.category_axis.minimum_scale, ch.category_axis.maximum_scale = math.floor(min(xs)), math.ceil(max(xs))
    ch.category_axis.crosses = ch.value_axis.crosses = XL_AXIS_CROSSES.MINIMUM   # 축은 왼쪽, 아래
    for i, ser in enumerate(ch.plots[0].series):
        ser.format.line.fill.background()
        ser.marker.style, ser.marker.size = XL_MARKER_STYLE.CIRCLE, 4
        ser.marker.format.fill.solid()
        ser.marker.format.fill.fore_color.rgb = _rgb(PALETTE[i % len(PALETTE)])
        ser.marker.format.line.fill.background()
    _chart_font(ch)
    _obst(s, gf)
    return gf


def hist_bins(values, edges, unit=""):
    """히스토그램 구간 이름과 건수. edges 는 경계 목록"""
    cats = [f"{a:g} ~ {b:g}{unit}" for a, b in zip(edges, edges[1:])]
    counts = [sum(1 for v in values if a <= v < b or (k == len(edges) - 2 and v == b))
              for k, (a, b) in enumerate(zip(edges, edges[1:]))]
    return cats, counts


def screenshot(s, path, x, y, w, h, marks=(), color=RED):
    """캡처를 칸에 넣고 marks 영역 표시"""
    pic = s.shapes.add_picture(path, Inches(x), Inches(y), width=Inches(w))
    if pic.height > Inches(h):
        r = Inches(h) / pic.height
        pic.width, pic.height = int(pic.width * r), Inches(h)
    pic.left = Inches(x) + (Inches(w) - pic.width) // 2
    pic.line.color.rgb = _rgb(GRAY_L)
    pic.line.width = Pt(0.75)
    _wall(s, pic)
    return marks_on(s, pic, marks, color)


def marks_on(s, pic, marks, color=RED):
    """그림 위 영역 네모 및 번호"""
    px, py, pw, ph = pic.left / _E, pic.top / _E, pic.width / _E, pic.height / _E
    marks = _snap_marks(marks, pw, ph)
    placed = []
    for m in marks:
        if len(m) == 5:
            x0, y0, x1, y1, n = m
            box(s, px + x0 * pw, py + y0 * ph, (x1 - x0) * pw, (y1 - y0) * ph, fill=None, line=color, lw=1.5,
                geom="round", radius=0.04)
            cx, cy = px + x0 * pw - 0.06, py + y0 * ph - 0.06
        else:
            fx, fy, n = m
            cx, cy = px + fx * pw, py + fy * ph
        if any(abs(cx - qx) < NUM_D + 0.02 and abs(cy - qy) < NUM_D + 0.02 for qx, qy in placed):
            _warn(s, f"캡처 번호 {n} 이 다른 번호와 겹침: 두 네모 사이를 띄움")
        placed.append((cx, cy))
        _obst(s, num_badge(s, cx - NUM_D / 2, cy - NUM_D / 2, n, color, line=WHITE))
    return pic


def snap_marks(marks, pw, ph, tol=0.2):
    """영역 변 맞추기 (공개 이름)"""
    return _snap_marks(marks, pw, ph, tol)


def _snap_marks(marks, pw, ph, tol=0.2):
    """가까운 변 (0.2 in 이내) 정렬"""
    boxes = [list(m) for m in marks]
    for k, size in ((0, pw), (2, pw), (1, ph), (3, ph)):
        idx = sorted((i for i, m in enumerate(boxes) if len(m) == 5), key=lambda i: boxes[i][k])
        span = (lambda j: (boxes[j][2] - boxes[j][0]) * pw) if k in (0, 2) else (lambda j: (boxes[j][3] - boxes[j][1]) * ph)
        group = []
        for i in idx + [None]:
            if i is not None and group and (boxes[i][k] - boxes[group[0]][k]) * size <= min(
                    tol, 0.3 * min(span(i), span(group[0]))):
                group.append(i)
                continue
            if len(group) > 1:
                v = sum(boxes[j][k] for j in group) / len(group)
                for j in group:
                    boxes[j][k] = v
            group = [i] if i is not None else []
    return [tuple(m) for m in boxes]


def kpi(s, x, y, w, h, value, unit, label_, sub=None, color=ACCENT):
    """큰 주황 숫자 핵심 수치. 덱의 1/5 장까지"""
    need = _text_w(str(value), 20) + (_text_w(" " + unit, 9) if unit else 0) + 0.25
    R = _reg(s)
    R["kpi_items"] = R.get("kpi_items", 0) + 1     # 직접 부른 kpi 도 남발 검사에 셈
    if need > w or _text_w(label_, 7.26) + 0.1 > w:
        _warn(s, f"핵심 수치 '{value}' 가 칸({w:.2f} in)보다 넓음: 한 줄의 수치를 줄임")
    bw = min(w, max(1.05, need, _text_w(label_, 7.26) + 0.2))
    if sub and w - bw <= 0.4:
        _warn(s, f"핵심 수치 '{value}' 의 설명 '{sub}' 이 들어갈 자리 없음: 수치를 줄이거나 설명을 결론 줄로")
    if not sub:
        x = x + (w - bw) / 2
    body = _obst(s, box(s, x, y, bw, h, fill=None))
    box(s, x + bw * 0.3, y + h * 0.58, bw * 0.4, 0.02, fill=color)
    tb = s.shapes.add_textbox(Inches(x), Inches(y + 0.02), Inches(bw), Inches(h * 0.55))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_ = tf.paragraphs[0]
    p_.alignment = PP_ALIGN.CENTER
    r1 = p_.add_run()
    r1.text = str(value)
    _font(r1, 20, True, color)
    if unit:
        r2 = p_.add_run()
        r2.text = " " + unit
        _font(r2, 9, True, color)
    text(s, x, y + h * 0.64, bw, 0.2, label_, size=7.26, bold=True, color=SLATE9, align="c")
    if sub and w - bw > 0.4:
        text(s, x + bw + 0.12, y + h / 2 - 0.1, w - bw - 0.12, 0.2, sub, size=7.26, color=SLATE9, anchor="m")
    return body


def icon_cards(s, x, y, w, items, card=0.46):
    """흰 둥근 카드 아이콘 줄. 바닥 y 반환"""
    n = len(items)
    gap = (w - card * n) / max(n + 1, 2)
    for i, (ic, t) in enumerate(items):
        cx = x + gap + i * (card + gap)
        box(s, cx, y, card, card + 0.12, fill=WHITE, line=LINE_L, lw=0.5, geom="round", radius=0.1, shadow=True)
        icon(s, ic, cx + 0.08, y + 0.05, w=card - 0.16)
        text(s, cx - 0.1, y + card - 0.03, card + 0.2, 0.14, t, size=5.44, bold=True, color=SLATE9, align="c")
    return y + card + 0.12


def zoom_wedge(s, src, x1, ty0, ty1):
    """확대 표시, 펼친 내용보다 먼저 그림"""
    x0, y0, _, y1 = _bb(src)
    x0 = _bb(src)[2]
    fb = s.shapes.build_freeform(Inches(x0), Inches(y0))
    fb.add_line_segments([(Inches(x1), Inches(ty0)), (Inches(x1), Inches(ty1)), (Inches(x0), Inches(y1))], close=True)
    sh = fb.convert_to_shape()
    _no_style_effects(sh)
    sppr = sh._element.spPr
    for el in sppr.findall(qn("a:solidFill")):
        sppr.remove(el)
    grad = etree.fromstring(
        f'<a:gradFill xmlns:a="{_A}" flip="none" rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/><a:alpha val="90000"/></a:schemeClr></a:gs>'
        '<a:gs pos="100000"><a:schemeClr val="bg1"><a:alpha val="0"/></a:schemeClr></a:gs>'
        '</a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>')
    sppr.find(qn("a:custGeom")).addnext(grad)
    _line(sh, None)
    return sh


def flow_arrow(s, x, y, w=0.29, h=0.33, direction="right"):
    """파란 흐름 화살표"""
    rot = {"down": 0, "left": 90, "up": 180, "right": 270}[direction]
    if direction in ("left", "right"):  # 회전 뒤 w x h 가 되게 회전 전 크기로
        x, y, w, h = x + (w - h) / 2, y + (h - w) / 2, h, w
    sh = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    _no_style_effects(sh)
    sh.rotation = rot
    sppr = sh._element.spPr
    for el in sppr.findall(qn("a:solidFill")):
        sppr.remove(el)
    grad = etree.fromstring(
        f'<a:gradFill xmlns:a="{_A}"><a:gsLst>'
        '<a:gs pos="0"><a:schemeClr val="bg1"/></a:gs>'
        '<a:gs pos="85000"><a:schemeClr val="tx2"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr></a:gs>'
        '<a:gs pos="100000"><a:schemeClr val="tx2"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:gs>'
        '</a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>')
    sppr.find(qn("a:prstGeom")).addnext(grad)
    _line(sh, None)
    return sh


def step_badge(s, gx0, gx1, cy, icon_key, label_, top=None):
    """패널 틈에 배지 + 흐름 화살표 + 이름"""
    cx = (gx0 + gx1) / 2
    bw = min(0.46, gx1 - gx0 - 0.1)
    icon(s, icon_key, cx - bw / 2, cy - bw - 0.12, w=bw)
    flow_arrow(s, cx - 0.17, cy - 0.06, w=0.34, h=0.26, direction="right")
    lw = gx1 - gx0 - 0.04
    lines, cur = [], ""
    for wd in label_.split():
        cand = f"{cur} {wd}".strip()
        if cur and _text_w(cand, 8) * 1.1 > lw - 0.04:
            lines.append(cur)
            cur = wd
        else:
            cur = cand
    lines.append(cur)
    text(s, gx0 + 0.02, cy + 0.26, lw, 0.16 * len(lines) + 0.04, lines, size=8, bold=True, color=SLATE9, align="c")


def gantt(s, x, y, w, periods, tasks, row_h=0.3, label_w=2.0, today=None, today_label="기준일"):
    """일정 표. kind: done, now, plan, mile"""
    if not periods or not tasks:
        raise ValueError("gantt: 기간이나 항목이 없음")
    pw = (w - label_w) / len(periods)
    hh = 0.27
    box(s, x, y, label_w, hh, "항목", fill=TBL_HEAD, line=WHITE, lw=0.5, size=7.26, bold=True, color=WHITE)
    for i, pd in enumerate(periods):
        box(s, x + label_w + i * pw, y, pw, hh, pd, fill=TBL_HEAD, line=WHITE, lw=0.5, size=7.26, bold=True,
            color=WHITE)
    body_h = row_h * len(tasks)
    for i in range(1, len(periods)):
        box(s, x + label_w + i * pw - 0.0035, y + hh, 0.007, body_h, fill=LINE_L)
    fills = {"done": GRAY_L, "now": ACCENT, "plan": SLATE6}
    for r, t in enumerate(tasks):
        name, a, b, kind = t[:4]
        note = t[4] if len(t) > 4 else ""
        ry = y + hh + r * row_h
        box(s, x, ry, label_w, row_h, name, fill=None, size=7.26, bold=kind == "now", align="l",
            color=ACCENT if kind == "now" else BLACK, inset=(0.08, 0, 0.04, 0))
        box(s, x, ry + row_h - 0.0035, w, 0.007, fill=LINE_L)
        if kind == "mile":
            cx = x + label_w + a * pw
            box(s, cx - 0.09, ry + row_h / 2 - 0.09, 0.18, 0.18, fill=ACCENT, geom="diamond")
            if note:  # 장 밖으로 나가지 않게 마름모 왼쪽에
                nw = _text_w(note, 6.35) + 0.1
                right_room = x + w - (cx + 0.12)
                nx, al = (cx + 0.12, "l") if right_room >= nw else (cx - 0.12 - nw, "r")
                text(s, nx, ry + row_h / 2 - 0.1, nw, 0.2, note, size=6.35, bold=True, color=ACCENT, align=al,
                     anchor="m")
        else:
            if kind not in fills:
                raise ValueError(f"gantt 종류 '{kind}' 없음: done, now, plan, mile")
            box(s, x + label_w + a * pw + 0.03, ry + 0.07, max((b - a) * pw - 0.06, 0.05), row_h - 0.14, note,
                fill=fills[kind], geom="round", radius=0.5, size=6.35, bold=True, color=WHITE, inset=(0.04, 0, 0.04, 0))
    if today is not None:
        tx = x + label_w + today * pw
        box(s, tx, y + hh, 0.001, body_h, fill=None, line=RED, lw=1, dash="dash")
        text(s, tx - 0.3, y + hh + body_h + 0.02, 0.6, 0.18, today_label, size=6.35, bold=True, color=RED, align="c")
    return y + hh + body_h


class _Cols(list):
    """report_panels 결과: 칸 목록 + gaps"""


def banner(s, t, tags=None, current=None, y=REPORT_TOP, x=None, w=None, h=0.4):
    """핵심 문장 띠. 띠 아래 y 반환"""
    x = REPORT_LEFT if x is None else x
    w = (REPORT_RIGHT - REPORT_LEFT) if w is None else w
    _obst(s, box(s, x, y, w, h, fill="F2F2F2"))
    text(s, x + 0.2, y, w - 0.4, h, t, size=10.5, bold=True, color=BLACK, align="c")
    if tags:
        tw_ = 0.5
        for k, tg in enumerate(reversed(tags)):
            i = len(tags) - 1 - k
            on = i == current
            box(s, x + w - 0.15 - (k + 1) * (tw_ + 0.04), y + (h - 0.2) / 2, tw_, 0.2, tg,
                fill=ACCENT if on else WHITE, line=ACCENT if on else GRAY_L, lw=0.75, size=6.35, bold=True,
                color=WHITE if on else SLATE9, geom="round", radius=0.3, inset=(0, 0, 0, 0))
    return y + h + 0.12


def report_panels(s, titles, widths=None, arrows=False, hub_gap=None, steps=None, top=REPORT_TOP, footnote=None):
    """보고형 패널과 칸 (Col) 목록"""
    n = len(titles)
    if n == 0:
        raise ValueError("패널 제목이 없음")
    if widths is None:
        widths = REPORT_WIDTHS if n == 3 else [(REPORT_RIGHT - REPORT_LEFT - 0.67 * (n - 1)) / n] * n
    if not arrows and not steps and hub_gap is None and n > 1:  # 화살표 없으면 넓은 틈이 빈 띠로 보임
        k = (REPORT_RIGHT - REPORT_LEFT - 0.3 * (n - 1)) / sum(widths)
        widths = [w_ * k for w_ in widths]
    room = REPORT_RIGHT - REPORT_LEFT - sum(widths)
    gaps = [room / max(n - 1, 1)] * max(n - 1, 0)
    if hub_gap is not None:
        if not (0 <= hub_gap < n - 1):
            raise ValueError(f"hub_gap={hub_gap} 는 틈 번호(0 ~ {n - 2})가 아님")
        if n > 2:
            gaps = [0.67] * (n - 1)
            gaps[hub_gap] = room - 0.67 * (n - 2)
    bottom = REPORT_BOTTOM - (0.22 if footnote else 0)
    if footnote:  # 패널 아래 각주 한 줄
        text(s, REPORT_LEFT, bottom + 0.04, REPORT_RIGHT - REPORT_LEFT, 0.17, footnote, size=6.35, color=SLATE9)
    mid = (top + 0.29 + bottom) / 2
    cols, x = _Cols(), REPORT_LEFT
    cols.gaps = []
    for i, (t, w) in enumerate(zip(titles, widths)):
        panel(s, x, top, w, bottom - top, t)
        _reg(s).setdefault("panel_titles", set()).add(t)
        cols.append(Col(s, x + PAD, w - 2 * PAD, top=top + (BODY_TOP - REPORT_TOP), bottom=bottom - (REPORT_BOTTOM - BODY_BOTTOM)))
        if i < n - 1:
            cols.gaps.append((x + w, x + w + gaps[i]))
            if steps and i < len(steps) and steps[i]:
                step_badge(s, x + w, x + w + gaps[i], mid, *steps[i])
            elif arrows and i != hub_gap:
                big_arrow(s, x + w + gaps[i] / 2 - 0.19, mid - 0.165)
            x += w + gaps[i]
    return cols


def _em(ch):
    """굵은 글자 폭 (em). check.py 와 같은 표"""
    if ord(ch) >= 0x1100:
        return 1.0
    if ch == " ":
        return 0.29
    if ch.isupper():
        return 0.68
    if ch.isdigit():
        return 0.62
    if ch.islower():
        return 0.54
    return 0.48


def _text_w(t, size):
    """글자 폭 어림 (인치)"""
    return sum(_em(c) for c in t) * size / 72


_NOT_END = ("e.g.", "i.e.", "vs.", "etc.", "No.", "Inc.", "Co.", "Ltd.", "Fig.", "p.")


def ko_breaks(text, width, size):
    """문장 끝, 긴 문장은 쉼표 뒤에서 줄 바꿈"""
    import re
    out = []
    parts = re.split(r"(?<=[.!?])(?<!\.\.)(?<![.][.][.])\s+(?=\S)", text)
    sents = []
    for p_ in parts:   # 약어로 끝난 조각은 다음과 붙임
        if sents and sents[-1].endswith(_NOT_END):
            sents[-1] += " " + p_
        else:
            sents.append(p_)
    for snt in sents:
        rest = snt
        while width > 0.3 and _text_w(rest, size) > width:
            cut = None
            for m in re.finditer(r",\s+", rest):
                head = rest[:m.start() + 1]
                if head.count("(") > head.count(")"):   # 괄호 안 쉼표에서는 끊지 않음
                    continue
                wh = _text_w(head, size)
                if wh <= width and wh >= width * 0.45:
                    cut = m
            if cut is None:
                break
            out.append(rest[:cut.start() + 1])
            rest = rest[cut.end():]
        out.append(rest)
    return out


# --- 수식 (PowerPoint 수식 개체 OMML)
MATH_FONT = "Cambria Math"
_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_SYM = {"times": "×", "pm": "±", "le": "≤", "ge": "≥", "approx": "≈", "neq": "≠", "Delta": "Δ", "sigma": "σ",
        "mu": "μ", "alpha": "α", "beta": "β", "sum": "∑", "%": "%", ",": " "}


def _m(tag, *kids, **attrs):
    el = etree.Element(f"{{{_M}}}{tag}", nsmap={"m": _M})
    for k, v in attrs.items():
        el.set(f"{{{_M}}}{k}", v)
    el.extend(kids)
    return el


def _mi(ch):
    """라틴 한 글자를 수학 기울임 글자로"""
    if "A" <= ch <= "Z":
        return chr(0x1D434 + ord(ch) - 65)
    if "a" <= ch <= "z":
        return "\u210E" if ch == "h" else chr(0x1D44E + ord(ch) - 97)
    return ch


def _mr(t, sz, upright=False):
    r = _m("r")
    if upright:
        r.append(_m("rPr", _m("sty", val="p")))
    rpr = etree.SubElement(r, qn("a:rPr"), lang="en-US", sz=str(round(sz * 100)))
    if not upright:
        rpr.set("i", "1")
    etree.SubElement(etree.SubElement(rpr, qn("a:solidFill")), qn("a:srgbClr"), val=BLACK)
    etree.SubElement(rpr, qn("a:latin"), typeface=MATH_FONT)
    etree.SubElement(rpr, qn("a:ea"), typeface=FONT)
    etree.SubElement(r, f"{{{_M}}}t").text = t
    return r


def _hangul(c):
    return "가" <= c <= "힣"


def _group(src, i):
    """{...} 또는 한 글자, 앞 공백 무시. (내용, 다음 위치)"""
    while i < len(src) and src[i] == " ":
        i += 1
    if i >= len(src):
        raise ValueError(f"식 끝에 인자 없음: {src}")
    if src[i] == "\\":
        name = re.match(r"\\([A-Za-z]+|.)", src[i:]).group(0)
        return name, i + len(name)
    if src[i] != "{":
        return src[i], i + 1
    depth = 0
    for j in range(i, len(src)):
        depth += {"{": 1, "}": -1}.get(src[j], 0)
        if depth == 0:
            return src[i + 1:j], j + 1
    raise ValueError(f"식 괄호 짝이 맞지 않음: {src}")


def _tokens(src):
    """식 조각 (종류, 값), 종류 cmd sub sup grp word ko sym"""
    i = 0
    while i < len(src):
        c = src[i]
        if c == "\\":
            m_ = re.match(r"\\([A-Za-z]+|.)", src[i:])
            if not m_:
                raise ValueError(f"식 끝의 \\: {src}")
            name = m_.group(1)
            i += len(m_.group(0))
            args = []
            for _ in range({"frac": 2, "hat": 1, "bar": 1, "text": 1}.get(name, 0)):
                a, i = _group(src, i)
                args.append(a)
            if name not in ("frac", "hat", "bar", "text") and name not in _SYM:
                raise ValueError(f"식에 모르는 명령 \\{name}: {src}")
            yield "cmd", (name, args)
        elif c in "_^":
            a, i = _group(src, i + 1)
            yield ("sub" if c == "_" else "sup"), a
        elif c == "{":
            a, i = _group(src, i)
            yield "grp", a
        elif c == "}":
            raise ValueError(f"식에 짝 없는 }}: {src}")
        elif c == " ":
            i += 1
        elif c.isascii() and c.isalpha():
            j = i
            while j < len(src) and src[j].isascii() and src[j].isalpha():
                j += 1
            yield "word", src[i:j]
            i = j
        elif _hangul(c):
            j = i
            while j < len(src) and (_hangul(src[j]) or (src[j] == " " and j + 1 < len(src) and _hangul(src[j + 1]))):
                j += 1
            yield "ko", src[i:j]
            i = j
        else:
            j = i + 1
            while j < len(src) and not (src[j].isascii() and src[j].isalpha()) and not _hangul(src[j]) \
                    and src[j] not in "\\_^ {}":
                j += 1
            yield "sym", src[i:j]
            i = j


def _omml(src, sz):
    """작은 LaTeX 표기, 한 글자 변수만 기울임"""
    out = []
    toks = list(_tokens(src))
    k = 0
    while k < len(toks):
        kind, val = toks[k]
        k += 1
        if kind == "cmd" and val[0] == "sum":   # 큰 연산자, 아래 위 한계
            lim = {"sub": [], "sup": []}
            while k < len(toks) and toks[k][0] in lim:
                lim[toks[k][0]] = _omml(toks[k][1], sz)
                k += 1
            out.append(_m("nary", _m("naryPr", _m("chr", val="∑"), _m("limLoc", val="undOvr"),
                                     *([_m("subHide", val="1")] if not lim["sub"] else []),
                                     *([_m("supHide", val="1")] if not lim["sup"] else [])),
                          _m("sub", *lim["sub"]), _m("sup", *lim["sup"]), _m("e")))
        elif kind == "cmd":
            name, args = val
            if name == "frac":
                out.append(_m("f", _m("num", *_omml(args[0], sz)), _m("den", *_omml(args[1], sz))))
            elif name in ("hat", "bar"):
                out.append(_m("acc", _m("accPr", _m("chr", val="\u0302" if name == "hat" else "\u0305")),
                              _m("e", *_omml(args[0], sz))))
            elif name == "text":
                out.append(_mr(args[0], sz, True))
            else:
                out.append(_mr(_SYM[name], sz, True))
        elif kind in ("sub", "sup"):
            base = out.pop() if out else _mr("", sz, True)
            out.append(_m("sSub" if kind == "sub" else "sSup", _m("e", base), _m(kind, *_omml(val, sz))))
        elif kind == "grp":
            out.append(_m("d", _m("dPr", _m("begChr", val=""), _m("endChr", val="")), _m("e", *_omml(val, sz))))
        elif kind == "word":
            out.append(_mr(_mi(val), sz) if len(val) == 1 else _mr(val, sz, True))
        elif kind == "ko":
            out.append(_mr(val, sz, True))
        else:
            out.append(_mr(val.replace("-", "\u2212"), sz, True))
    for i, el in enumerate(out):   # 합 기호 뒤 한 항을 본체로
        if el.tag == f"{{{_M}}}nary" and i + 1 < len(out):
            el.find(f"{{{_M}}}e").append(out.pop(i + 1))
    return out


def _linear(src):
    """수식을 못 그리는 프로그램용 한 줄 글"""
    op = lambda t: f" {t} " if t in "=+-×±≤≥≈≠" else t
    out = []
    for kind, val in _tokens(src):
        if kind == "cmd":
            name, args = val
            out.append(f"({_linear(args[0])}) / ({_linear(args[1])})" if name == "frac"
                       else _linear(args[0]) if args else op(_SYM[name]))
        elif kind in ("sub", "sup"):
            t = _linear(val)
            out.append(("_" if kind == "sub" else "^") + (f"({t})" if len(t) > 1 else t))
        elif kind == "grp":
            out.append(f"({_linear(val)})")
        else:
            out.append(op(val) if kind == "sym" else val)
    return re.sub(r" +", " ", "".join(out)).strip()


def _math(src, sz, display):
    """a14:m 하나. display 면 가운데 수식 문단"""
    om = _m("oMath", *_omml(src, sz))
    if display:
        om = _m("oMathPara", _m("oMathParaPr", _m("jc", val="centerGroup")), om)
    el = etree.Element(f"{{{_A14}}}m", nsmap={"a14": _A14})
    el.append(om)
    el.set("lin", _linear(src))
    return el


def _alt_content(tb):
    """수식 글상자를 mc:AlternateContent 로 감쌈. Fallback 은 한 줄 글"""
    sp = tb._element
    fb = copy.deepcopy(sp)
    for el in list(fb.iter(f"{{{_A14}}}m")):
        r = etree.Element(qn("a:r"))
        etree.SubElement(r, qn("a:rPr"), lang="en-US", sz="1050")
        etree.SubElement(r, qn("a:t")).text = el.get("lin")
        el.getparent().replace(el, r)
    for el in sp.iter(f"{{{_A14}}}m"):
        del el.attrib["lin"]
    parent = sp.getparent()
    idx = parent.index(sp)
    ac = etree.Element(f"{{{_MC}}}AlternateContent", nsmap={"mc": _MC})
    ch = etree.SubElement(ac, f"{{{_MC}}}Choice", nsmap={"a14": _A14}, Requires="a14")
    parent.remove(sp)
    ch.append(sp)
    etree.SubElement(ac, f"{{{_MC}}}Fallback").append(fb)
    parent.insert(idx, ac)


class Col:
    """패널 안 세로 칸, done() 에서 높이 분배"""
    MIN_GAP, MAX_GAP, MAX_ROW = 0.12, 0.50, 0.32

    def __init__(self, s, x, w, top=BODY_TOP, bottom=BODY_BOTTOM):
        self.s, self.x, self.w, self.top, self.bottom = s, x, w, top, bottom
        self.items = []  # (높이, 그리기, 행 수, 행 추가 한도)
        _reg(s).setdefault("cols", []).append(self)

    def _add(self, h, draw, rows=0, room=0.0):
        self.items.append((h, draw, rows, room))
        return self

    def __getitem__(self, i):
        """i 번째 블록이 만든 도형"""
        if not hasattr(self, "out"):
            raise ValueError("done() 을 먼저 부른다")
        return self.out[i]

    def box(self, t, h=0.3, **kw):
        kw = {"fill": SLATE6, "bold": True, "color": WHITE, **kw}
        return self._add(h, lambda y, r: box(self.s, self.x, y, self.w, h, t, **kw))

    def arrow(self):
        return self._add(0.33, lambda y, r: big_arrow(self.s, self.x + self.w / 2 - 0.19, y, direction="down"))

    def chip_tree(self, rows, pitch=0.255, **kw):
        """value_w 없으면 값 상자 오른쪽 끝을 맞춤"""
        if not rows:
            raise ValueError("chip_tree 행이 없음")
        if "value_w" not in kw:
            ind = kw.get("indent", 0.31)
            kw["value_w"] = lambda d, k, w=self.w: w - d * ind - (0.34 if len(k) > 2 else 0.32) - 0.03
        return self._add(pitch * (len(rows) - 1) + 0.17,
                         lambda y, r: chip_tree(self.s, self.x, y, rows, pitch=pitch + r, **kw),
                         rows=len(rows) - 1, room=0.05)

    def table(self, rows, ratio=None, row_h=0.26, **kw):
        if not rows or not rows[0]:
            raise ValueError("표 행이 없음")
        ratio = ratio or [1] * len(rows[0])
        cw = [self.w * k / sum(ratio) for k in ratio]
        return self._add(min(row_h, 0.27) + row_h * (len(rows) - 1),
                         lambda y, r: table(self.s, self.x, y, cw, rows, row_h=row_h + r, **kw),
                         rows=len(rows) - 1, room=max(self.MAX_ROW - row_h, 0))

    def chart(self, kind, categories, series, h=1.9, title=None, **kw):
        """차트 블록. title 은 파란 9pt 제목"""
        th = 0.26 if title else 0

        def draw(y, r):
            if title:
                text(self.s, self.x, y, self.w, 0.22, title, size=9, bold=True, color=BLUE, align="c")
            return chart(self.s, self.x, y + th, self.w, h + r, kind, categories, series, **kw)
        return self._add(h + th, draw, rows=1, room=1.2)

    def scatter(self, series, x_title, y_title, h=2.2, **kw):
        """산점도 블록. 축 이름에 단위 (예: '목표 두께 (µm)')"""
        return self._add(h, lambda y, r: scatter_chart(self.s, self.x, y, self.w, h + r, series, x_title, y_title,
                                                       **kw), rows=1, room=1.2)

    def hist(self, values, edges, name="건수", unit="", count_unit="건", h=1.9, **kw):
        """히스토그램 블록 (편집 가능한 세로 막대)"""
        cats, counts = hist_bins(values, edges, unit)
        return self.chart("col", cats, {name: counts}, h=h, number_format=f'#,##0"{count_unit}"', **kw)

    def icon_cards(self, items, card=0.46):
        """아이콘 카드 줄"""
        return self._add(card + 0.12, lambda y, r: icon_cards(self.s, self.x, y, self.w, items, card=card))

    def kpis(self, items, h=0.85, gap=0.15):
        """핵심 수치 한 줄. items=[(값, 단위, 이름)]"""
        R = _reg(self.s)
        R["kpi_rows"] = R.get("kpi_rows", 0) + 1
        n = len(items)
        cw = (self.w - gap * (n - 1)) / n

        def draw(y, r):
            return [kpi(self.s, self.x + i * (cw + gap), y, cw, h, *it) for i, it in enumerate(items)]
        return self._add(h, draw)

    def section(self, title, rows, row_h=0.3, **kw):
        """회색 머리 띠 목록"""
        return self._add(0.27 + row_h * len(rows), lambda y, r: section(self.s, self.x, y, self.w, title, rows,
                                                                       row_h=row_h + r, **kw),
                         rows=len(rows), room=max(0.36 - row_h, 0))

    def steps(self, items, color=ACCENT):
        """번호 설명 목록. items=[(번호, 제목, 설명)]"""
        w = self.w - NUM_D - 0.1
        for _, ttl, d in items:
            _unit(self.s, "steps", _td_h(ttl, d, w))
        gap = 0.14

        def draw(y, r):
            U = _unit_of(self.s, "steps")
            yy = y
            for n, ttl, d in items:
                top = title_desc(self.s, self.x + NUM_D + 0.1, yy, w, U, ttl, d, tcolor=SLATE9)
                num_badge(self.s, self.x, top + (_lh(9) - NUM_D) / 2, n, color)
                yy += U + gap
        return self._add(lambda: len(items) * _unit_of(self.s, "steps") + gap * (len(items) - 1), draw)

    def gantt(self, periods, tasks, row_h=0.3, **kw):
        return self._add(0.27 + row_h * len(tasks) + (0.22 if kw.get("today") is not None else 0),
                         lambda y, r: gantt(self.s, self.x, y, self.w, periods, tasks, row_h=row_h, **kw))

    def image(self, path, h, marks=()):
        """그림 블록. h 는 최대 높이"""
        from PIL import Image as _Im
        with _Im.open(path) as im:
            iw, ih = im.size
        real = min(h, self.w * ih / iw)
        if _default_mpl(path):
            _warn(self.s, f"그림 {os.path.basename(path)} 이 matplotlib 기본 색 (주황, 파랑): 그리기 전 mpl_style()")
        pad = 0.12 if marks else 0.0  # 번호 원이 그림 위로 0.17 in 나옴
        return self._add(real + pad, lambda y, r: screenshot(self.s, path, self.x, y + pad, self.w, real, marks))

    def rule_list(self, title, items, pitch=0.29):
        """items=[(아이콘, 글[, 설명])]"""
        kind = "rule_desc" if any(len(it) > 2 for it in items) else "rule"
        if kind == "rule_desc":
            pitch = max(pitch, 0.38)
        _unit(self.s, kind, pitch)
        return self._add(lambda: 0.33 + _unit_of(self.s, kind) * len(items),
                         lambda y, r: rule_list(self.s, self.x, y, self.w, title, items, pitch=_unit_of(self.s, kind)))

    def tab_box(self, label, lines, color=ACCENT):
        """왼쪽 위 이름표 묶음. lines=[글, ...]"""
        last = (label.split() or [""])[-1]   # 결함, 포함 같은 두 글자 명사는 통과
        if not last or re.search(r"[가-힣]{2}(함|됨|듦)$|(했음|있음|없음|하기|한다|된다|이다)$", last):
            _warn(self.s, f"이름표 '{label}' 는 명사로 (예: '결론', '예측 대상')")
        w = self.w - 0.32
        body_txt = lambda: sum(_nl("• " + ln, w, 8) for ln in lines) * _lh(8)
        _unit(self.s, "tab_box", 0.14 + body_txt() + 0.12)

        def draw(y, r):
            U = _unit_of(self.s, "tab_box")
            lw_ = _text_w(label, 9) + 0.3
            body = box(self.s, self.x, y + 0.11, self.w, U, fill=None, line=GRAY_L, lw=0.75)
            _wall(self.s, body)
            box(self.s, self.x, y, lw_, 0.22, label, fill=color, size=9, bold=True, color=WHITE, inset=(0, 0, 0, 0))
            text(self.s, self.x + 0.16, y + 0.25, w, U - 0.2, ["• " + ln for ln in lines], size=8, color=BLACK)
            return body
        _reg(self.s).setdefault("list_titles", []).append(label)
        _reg(self.s).setdefault("tab_labels", []).append(label)
        return self._add(lambda: 0.11 + _unit_of(self.s, "tab_box"), draw)

    def num_cards(self, items, color=ACCENT):
        """번호 카드. items=[(제목, 설명)]"""
        w = self.w - 0.12 - NUM_D - 0.1 - 0.1
        for ttl, d in items:
            _unit(self.s, "num_cards", _td_h(ttl, d, w) + 0.2)
        gap = 0.12

        def draw(y, r):
            U = _unit_of(self.s, "num_cards")
            yy = y
            for n, (ttl, d) in enumerate(items, 1):
                box(self.s, self.x, yy, self.w, U, fill=WHITE, line=GRAY_L, lw=0.75, geom="round", radius=0.08)
                top = title_desc(self.s, self.x + 0.12 + NUM_D + 0.1, yy, w, U, ttl, d)
                num_badge(self.s, self.x + 0.12, top + (_lh(9) - NUM_D) / 2, n, color)
                yy += U + gap
        return self._add(lambda: len(items) * _unit_of(self.s, "num_cards") + gap * (len(items) - 1), draw)

    def chevrons(self, items, current=None, h=0.32):
        """단계 화살표 줄. current 는 주황"""
        n = len(items)

        w0 = (self.w + 0.12 * (n - 1)) / n
        for it in items:
            if _text_w(it, 8) > w0 - h - 0.24:
                _warn(self.s, f"단계 화살표 '{it}' 가 칸보다 김: 단계 수를 줄이거나 글을 짧게")

        def draw(y, r):
            w_ = w0
            for k, it in enumerate(items):
                on = k == current
                sh = box(self.s, self.x + k * (w_ - 0.12), y, w_, h, it, fill=ACCENT if on else SLATE3, size=8,
                         bold=True, color=WHITE if on else BLACK, geom="homePlate" if k == 0 else "chevron",
                         inset=(0.14, 0, 0.1, 0))
            _reg(self.s)["obst"].append(((self.x, y, self.x + self.w, y + h), 0))   # 화살표끼리 일부러 겹침
        return self._add(h, draw)

    def formula(self, expr, where=None, name=None):
        """논문식 수식 상자 (PowerPoint 수식 개체)"""
        exprs = [expr] if isinstance(expr, str) else list(expr)
        where = list(where or [])
        for x_ in where:
            if ":" not in x_:
                raise ValueError(f"formula where 는 '기호: 설명 (단위)': {x_}")
        w = self.w - 0.3
        h = 0.16 + (_lh(9) if name else 0) + sum(_lh(12) * (2.5 if "\\sum_" in e or "\\sum^" in e else 1.9 if "\\frac" in e else 1.2)
                                                       for e in exprs) \
            + ((1 + sum(_nl(_linear(x_), w, 7.26) + (0.35 if "_" in x_.split(":")[0] else 0) for x_ in where))
               * _lh(7.26) + 0.06 if where else 0)   # 첨자 기호 줄은 높음

        def draw(y, r):
            body = box(self.s, self.x, y, self.w, h, fill="F7F8FB", line=GRAY_L, lw=0.75)
            tb = text(self.s, self.x + 0.15, y + 0.08, w, h - 0.16, "")
            tf = tb.text_frame
            rows = ([(name, None, 9, True, "l")] if name else []) + [("", e, 12, False, "c") for e in exprs]
            if where:
                rows += [("여기서", None, 7.26, True, "l")]
                rows += [(": " + x_.split(":", 1)[1].strip(), x_.split(":", 1)[0], 7.26, False, "l") for x_ in where]
            for k, (txt, math_, sz, b_, al) in enumerate(rows):
                p_ = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p_.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER}[al]
                r_ = p_.runs[0] if p_.runs else p_.add_run()
                r_.text = txt
                _font(r_, sz, b_, SLATE9)
                if math_ is not None:
                    if not txt:
                        p_._p.remove(r_._r)
                    p_._p.insert(1 if p_._p.find(qn("a:pPr")) is not None else 0, _math(math_, sz, display=not txt))
            _alt_content(tb)
            return body
        return self._add(h, draw)

    def notes(self, lines, size=7.26):
        return self._add(_notes_h(lines, self.w, size), lambda y, r: notes(self.s, self.x, y, self.w, lines, size=size))

    def done(self):
        if getattr(self, "_done", False):
            raise ValueError("done() 은 칸마다 한 번만 부른다 (두 번 부르면 도형이 겹침)")
        self._done = True
        if not self.items:
            _warn(self.s, f"패널 칸이 비었음 (x {self.x:.2f}). 블록을 넣거나 패널을 뺌")
            return self.top
        self.items = [((h() if callable(h) else h), d_, n_, r_) for h, d_, n_, r_ in self.items]
        need = sum(i[0] for i in self.items)
        gaps = len(self.items) - 1
        free = self.bottom - self.top - need - self.MIN_GAP * gaps
        if free < -1e-6:
            raise ValueError(f"패널 칸(x {self.x:.2f})이 {-free:.2f} in 넘침: 블록을 줄이거나 표 행 높이를 낮춤")
        # 남는 높이를 늘림 한도 있는 블록에 고르게
        grow = [0.0] * len(self.items)
        open_ = [k for k, it in enumerate(self.items) if it[2] and it[3] > 0]
        while open_ and free > 1e-6:
            nrows = sum(self.items[k][2] for k in open_)
            step_ = min(free / nrows, min(self.items[k][3] - grow[k] for k in open_))
            for k in open_:
                grow[k] += step_
            free -= step_ * nrows
            open_ = [k for k in open_ if self.items[k][3] - grow[k] > 1e-6]
        gap = self.MIN_GAP + min(free / gaps, self.MAX_GAP - self.MIN_GAP) if gaps else 0.0
        y = self.top
        self.out = []
        for (h, draw, n, _), g_ in zip(self.items, grow):
            self.out.append(draw(y, g_))
            y += h + g_ * n + gap
        y -= gap
        if self.bottom - y > 0.3:
            _warn(self.s, f"패널 칸 아래 {self.bottom - y + 0.2:.2f} in 빔 (x {self.x:.2f}). 블록을 더하거나 패널을 줄임")
        return y


def doc_columns(right="source"):
    """설계문서형 열 (x, 너비)"""
    c = {"source": (0.75, 0.91), "edge": (2.10, 4.13), "top": DOC_TOP, "bottom": DOC_BOTTOM}
    if right == "source":
        c.update(platform=(6.71, 4.67), right=(11.65, 0.91))
    elif right == "side":
        raise ValueError("right='side' (오른쪽 설명 열) 는 없앰: right=None 으로 도식을 끝까지 쓰고 설명은 note_by")
    else:
        c["platform"] = (6.71, 5.85)
    return c


# --- 자동 연결 (선 방향, 번호 자리)
_E = 914400
_NORM = {"right": (1, 0), "left": (-1, 0), "bottom": (0, 1), "top": (0, -1)}


def _reg(s):
    p = s.part
    r = getattr(p, "_ixreg", None)
    if r is None:
        r = {"obst": [], "walls": [], "segs": [], "used": {}, "pending": [], "links": []}
        p._ixreg = r
    return r


def _warn(s, msg):
    """경고를 모은다. strict 저장이면 멈춤"""
    _reg(s).setdefault("warn", []).append(msg)
    print("주의: " + msg)


ALLOWED_PT = (5.44, 6.35, 7.26, 8.0, 8.16, 9.0, 9.07, 10.0, 10.5, 11.0, 12.0, 14.0, 16.32, 20.0)


def _snap_pt(pt):
    return min(ALLOWED_PT, key=lambda a: abs(a - pt))


def _bb(sh):
    x, y = sh.left / _E, sh.top / _E
    return (x, y, x + sh.width / _E, y + sh.height / _E)


def _col_box(s, key, shape, maxw, pic=None):
    """같은 열 상자 기록 (_settle 이 폭 맞춤)"""
    R = _reg(s)
    R.setdefault("colbox", {}).setdefault(key, []).append((shape, maxw, pic))
    if R.get("settled"):
        _settle(s, only=key)


def _settle(s, only=None):
    """같은 열 상자 폭을 최대 폭으로 통일"""
    R = _reg(s)
    for key, items in R.get("colbox", {}).items():
        if only is not None and key != only:
            continue
        W = min(max(sh.width for sh, _, _ in items), min(int(Inches(m)) for _, m, _ in items))
        for sh, _, pic in items:
            if sh.width >= W:
                continue
            old = _bb(sh)
            d = W - sh.width
            sh.left, sh.width = sh.left - d // 2, W
            new = _bb(sh)
            R["obst"] = [((new if (bb == old and sid == sh.shape_id) else bb), sid) for bb, sid in R["obst"]]
            if pic is not None:
                oldp = _bb(pic)
                pic.left = pic.left + d - d // 2
                R["obst"] = [((_bb(pic) if (bb == oldp and sid == pic.shape_id) else bb), sid) for bb, sid in R["obst"]]
    R["settled"] = True


def _obst(s, sh):
    _reg(s)["obst"].append((_bb(sh), sh.shape_id))
    return sh


def _wall(s, r):
    _reg(s)["walls"].append(r if isinstance(r, tuple) else _bb(r))


def _route(a, sa, b, sb, mid, check=True):
    """connect 선의 꺾임점. 못 그리면 None"""
    (x1, y1), (x2, y2) = _pt(a, sa), _pt(b, sb)
    ha, hb = sa in ("left", "right"), sb in ("left", "right")
    if ha and hb:
        if abs(y1 - y2) < 0.005:
            pts = [(x1, y1), (x2, y2)]
        else:
            xm = x1 + (x2 - x1) * mid
            pts = [(x1, y1), (xm, y1), (xm, y2), (x2, y2)]
    elif not ha and not hb:
        if abs(x1 - x2) < 0.005:
            pts = [(x1, y1), (x2, y2)]
        else:
            ym = y1 + (y2 - y1) * mid
            pts = [(x1, y1), (x1, ym), (x2, ym), (x2, y2)]
    else:
        pts = [(x1, y1), (x2, y1), (x2, y2)] if ha else [(x1, y1), (x1, y2), (x2, y2)]
    for (px, py), (qx, qy), (nx, ny), sign in ((pts[0], pts[1], _NORM[sa], 1), (pts[-2], pts[-1], _NORM[sb], -1)):
        dx, dy = qx - px, qy - py
        if check and ((nx and abs(dy) > 0.005) or (ny and abs(dx) > 0.005) or sign * (dx * nx + dy * ny) < 0.12):
            return None
    return pts


def _seg_hits(p, q, r, pad=0.03):
    """구간 p-q 가 사각형 r 을 지나는가"""
    x0, y0, x1, y1 = r[0] + pad, r[1] + pad, r[2] - pad, r[3] - pad
    if x0 >= x1 or y0 >= y1:
        return False
    sx0, sx1 = sorted((p[0], q[0]))
    sy0, sy1 = sorted((p[1], q[1]))
    return sx0 <= x1 and sx1 >= x0 and sy0 <= y1 and sy1 >= y0


def _cross(p, q, u, v):
    """엇갈림 1, 같은 줄 겹침 2"""
    h1, h2 = abs(p[1] - q[1]) < 1e-9, abs(u[1] - v[1]) < 1e-9
    if h1 == h2:
        k = 1 if h1 else 0
        if abs(p[k] - u[k]) > 0.03:
            return 0
        a0, a1 = sorted((p[1 - k], q[1 - k]))
        b0, b1 = sorted((u[1 - k], v[1 - k]))
        return 2 if min(a1, b1) - max(a0, b0) > 0.03 else 0
    hz, vt = ((p, q), (u, v)) if h1 else ((u, v), (p, q))
    hx0, hx1 = sorted((hz[0][0], hz[1][0]))
    vy0, vy1 = sorted((vt[0][1], vt[1][1]))
    return 1 if hx0 + 0.01 < vt[0][0] < hx1 - 0.01 and vy0 + 0.01 < hz[0][1] < vy1 - 0.01 else 0


def _pairs(a, b, strict=True):
    """선을 이을 변 조합 후보"""
    ax0, ay0, ax1, ay1 = _bb(a)
    bx0, by0, bx1, by1 = _bb(b)
    oy = min(ay1, by1) - max(ay0, by0) > 0.05
    ox = min(ax1, bx1) - max(ax0, bx0) > 0.05
    hz, vt = ("left", "right"), ("top", "bottom")
    for sa in _NORM:
        for sb in _NORM:
            if strict and oy and not ox and not (sa in hz and sb in hz):
                continue
            if strict and ox and not oy and not (sa in vt and sb in vt):
                continue
            yield sa, sb


def _before_draw(s):
    if not _reg(s).get("settled"):
        _settle(s)


def link(s, a, b, color=GRAY, width=1.5, head="end", dash=None, text=None, step=None, step_color=None, cond=None,
         cut=False):
    """a 에서 b 로 선 (변과 꺾임 자동)"""
    _before_draw(s)
    R = _reg(s)
    obst = [r for r, _ in R["obst"]] + [_bb(a), _bb(b)]
    ia, ib = a.shape_id, b.shape_id
    best = None
    for strict in (True, False):
        for sa, sb in _pairs(a, b, strict):
            mixed = (sa in ("left", "right")) != (sb in ("left", "right"))
            for mid in ((0.5,) if mixed else (0.5, 0.35, 0.65, 0.25, 0.75)):
                pts = _route(a, sa, b, sb, mid)
                if not pts:
                    continue
                segs = list(zip(pts, pts[1:]))
                if any(_seg_hits(p, q, r) for p, q in segs for r in obst):
                    continue
                cr = [_cross(p, q, u, v) for p, q in segs for u, v in R["segs"]]
                if 2 in cr:
                    continue
                length = sum(abs(q[0] - p[0]) + abs(q[1] - p[1]) for p, q in segs)
                sc = ((len(pts) - 2) + 3 * (R["used"].get((ia, sa), 0) + R["used"].get((ib, sb), 0))
                      + 1.5 * sum(cr) + 0.08 * length + abs(mid - 0.5))
                if best is None or sc < best[0]:
                    best = (sc, sa, sb, mid, pts)
                if len(pts) == 2:
                    break
        if best:
            break
    if best is None:
        _warn(s, f"상자를 피하는 선을 못 찾음 ({a.name} -> {b.name}). 상자 자리를 바꾸거나 connect 로 변을 정함")
        sa, sb, mid = "right", "left", 0.5
        pts = _route(a, sa, b, sb, mid, check=False)
    else:
        _, sa, sb, mid, pts = best
    c = _connect(s, a, sa, b, sb, color, width, head, dash, mid)
    R["used"][(ia, sa)] = R["used"].get((ia, sa), 0) + 1
    R["used"][(ib, sb)] = R["used"].get((ib, sb), 0) + 1
    R["segs"].extend(zip(pts, pts[1:]))
    R["links"].append((pts, ia, ib, f"{a.name} -> {b.name}"))
    if cut:  # 끊긴 연결: 가장 긴 구간에 빨간 ✕
        p_, q_ = max(zip(pts, pts[1:]), key=lambda pq: abs(pq[1][0] - pq[0][0]) + abs(pq[1][1] - pq[0][1]))
        mx, my = (p_[0] + q_[0]) / 2, (p_[1] + q_[1]) / 2
        x_ = box(s, mx - 0.13, my - 0.13, 0.26, 0.26, "✕", fill=WHITE, line=RED, lw=1.25, geom="oval", size=10,
                 bold=True, color=RED, inset=(0, 0, 0, 0))
        R["obst"].append((_bb(x_), -1))
    if step:
        R["pending"].append(("step", (pts, a, b), step[1], step[0],
                             step_color or (color if color != GRAY else ACCENT), cond))
    if text:
        R["pending"].append(("text", (pts, a, b), text, None, SLATE9, None))
    return c


def fan(s, hub, items, color=GRAY, width=1.5, head="end", direction="in", trunk=None, dash=None, step=None,
        cond=None, htype="triangle"):
    """여러 상자를 한 줄기로 hub 에 연결"""
    _before_draw(s)
    R = _reg(s)
    hx0, hy0, hx1, hy1 = _bb(hub)
    left = all(_bb(i)[2] <= hx0 for i in items)
    right = all(_bb(i)[0] >= hx1 for i in items)
    if not (left or right):
        raise ValueError("fan: items 가 hub 의 한쪽(왼쪽이나 오른쪽)에 모여 있지 않음. link 로 하나씩 이음")
    si, sh = ("right", "left") if left else ("left", "right")
    edge_i = max(_bb(i)[2] for i in items) if left else min(_bb(i)[0] for i in items)
    tx = trunk if trunk is not None else (edge_i + (hx0 if left else hx1)) / 2
    hp = _pt(hub, sh)
    ys = []
    for it in items:
        ip = _pt(it, si)
        a, sa, b, sb = (it, si, hub, sh) if direction == "in" else (hub, sh, it, si)
        x1, x2 = _pt(a, sa)[0], _pt(b, sb)[0]
        mid = (tx - x1) / (x2 - x1) if abs(x2 - x1) > 1e-6 else 0.5
        _connect(s, a, sa, b, sb, color, width, head, dash, mid, htype)
        R["segs"].extend([((ip[0], ip[1]), (tx, ip[1])), ((tx, hp[1]), (hp[0], hp[1]))])
        R["links"].append(([ip, (tx, ip[1]), (tx, hp[1]), hp], it.shape_id, hub.shape_id, f"fan {it.name}"))
        ys.append(ip[1])
    lo, hi = min(ys + [hp[1]]), max(ys + [hp[1]])
    R["segs"].append(((tx, lo), (tx, hi)))
    R["used"][(hub.shape_id, sh)] = R["used"].get((hub.shape_id, sh), 0) + 1
    if step:
        R["pending"].append(("step", ([(tx, lo), (tx, hi)], hub, hub), step[1], step[0],
                             color if color != GRAY else ACCENT, cond))
    return tx


def row_anchors(s, gf, rows=None, side="right"):
    """표 행마다 보이지 않는 연결 자리"""
    x0, y, w = gf.left / _E, gf.top / _E, gf.width / _E
    heights = [r.height / _E for r in gf.table.rows]
    tops = [y + sum(heights[:i]) for i in range(len(heights))]
    pick = range(len(heights) - 1) if rows is None else rows
    out = []
    for i in pick:
        top, h = tops[i + 1], heights[i + 1]
        ax = x0 + w - 0.18 if side == "right" else x0
        out.append(box(s, ax, top + 0.025, 0.18, h - 0.05, fill=None, line=None, geom="round"))
    return out


def bundle(s, lefts, rights, gap, label, icon_key="link", y=None, left_color=TEAL, right_color=BLUE):
    """점선 다발: lefts → 아이콘 → rights"""
    if not lefts and not rights:
        raise ValueError("bundle: 이을 대상이 없음")
    R = _reg(s)
    cx = (gap[0] + gap[1]) / 2
    ys = [(_bb(a)[1] + _bb(a)[3]) / 2 for a in (lefts or rights)]
    cy = y if y is not None else sum(ys) / len(ys)
    ic = icon(s, icon_key, cx - 0.165, cy - 0.3, w=0.33)
    room = min(gap[1] - gap[0] - 0.1, 0.85)
    lines, cur = [], ""
    for wd in label.split():
        cand = f"{cur} {wd}".strip()
        if cur and _text_w(cand, 8) * 1.1 > room - 0.06:
            lines.append(cur)
            cur = wd
        else:
            cur = cand
    lines.append(cur)
    n = len(lines)
    lw = min(gap[1] - gap[0] - 0.1, max(_text_w(l, 8) for l in lines) * 1.1 + 0.12)
    _obst(s, text(s, cx - lw / 2, cy + 0.06, lw, 0.16 * n + 0.04, lines, size=8, bold=True, color=left_color,
                  align="c"))
    hub = box(s, cx - 0.22, cy - 0.39, 0.44, 0.49 + 0.16 * n, fill=None, line=None, geom="round")
    if lefts:
        fan(s, hub, lefts, color=left_color, width=0.75, head="both", dash="sysDot", htype="arrow")
    hy0, hy1 = _bb(hub)[1], _bb(hub)[3]
    for t in rights:
        ty = (_bb(t)[1] + _bb(t)[3]) / 2
        side = "top" if ty < hy0 else ("bottom" if ty > hy1 else "right")
        _connect(s, hub, side, t, "left", right_color, 0.75, "both", "sysDot", 0.5, "arrow")
        pts = _route(hub, side, t, "left", 0.5, check=False)
        R["segs"].extend(zip(pts, pts[1:]))
        R["links"].append((pts, hub.shape_id, t.shape_id, f"bundle -> {t.name}"))
    return hub


def _free(R, bb, walls=True):
    x0, y0, x1, y1 = bb
    if x0 < 0.3 or x1 > 13.05 or y0 < 1.95 or y1 > 6.98:
        return False
    if any(not (x1 <= r[0] or x0 >= r[2] or y1 <= r[1] or y0 >= r[3]) for r, _ in R["obst"]):
        return False
    if any(_seg_hits(p, q, bb, pad=-0.02) for p, q in R["segs"]):
        return False
    crossed = 0
    for w in R["walls"]:  # 묶음 테두리를 걸치지 않음
        inside = w[0] <= x0 and x1 <= w[2] and w[1] <= y0 and y1 <= w[3]
        outside = x1 <= w[0] or x0 >= w[2] or y1 <= w[1] or y0 >= w[3]
        if not (inside or outside):
            crossed += 1
    if walls:
        return crossed == 0
    if crossed > 1:
        return False
    for w in R["walls"]:
        near = not (x1 + 0.1 <= w[0] or x0 - 0.1 >= w[2] or y1 + 0.1 <= w[1] or y0 - 0.1 >= w[3])
        crosses = not ((w[0] <= x0 and x1 <= w[2] and w[1] <= y0 and y1 <= w[3])
                       or x1 <= w[0] or x0 >= w[2] or y1 <= w[1] or y0 >= w[3])
        inside = w[0] <= x0 and x1 <= w[2] and w[1] <= y0 and y1 <= w[3]
        if near and not crosses and not inside:
            return False
    return True


def _line_spots(pts, w, h, gap=0.05):
    """선 옆 후보 자리. 긴 구간 가운데부터"""
    segs = sorted(zip(pts, pts[1:]), key=lambda pq: -(abs(pq[1][0] - pq[0][0]) + abs(pq[1][1] - pq[0][1])))
    for p, q in segs:
        horiz = abs(p[1] - q[1]) < 1e-9
        for t in (0.5, 0.3, 0.7, 0.15, 0.85):
            mx, my = p[0] + (q[0] - p[0]) * t, p[1] + (q[1] - p[1]) * t
            if horiz:
                for yy in (my - gap - h, my + gap):
                    for xx in (mx - w / 2, mx + 0.05, mx - w - 0.05):
                        yield xx, yy
            else:
                for xx in (mx + gap, mx - gap - w):
                    yield xx, my - h / 2


def _box_spots(sh, w, h):
    x0, y0, x1, y1 = _bb(sh)
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    yield from ((cx + 0.06, y1 + 0.06), (x0 + 0.02, y1 + 0.06), (cx + 0.06, y0 - h - 0.03), (x0 + 0.02, y0 - h - 0.03),
                (x1 + 0.06, cy - h / 2), (x0 - w - 0.06, cy - h / 2))


def _near_spots(px, py, w, h, reach=1.2, step_=0.05):
    """점 둘레 후보를 가까운 순으로"""
    n = int(reach / step_)
    cand = []
    for i in range(-n, n + 1):
        for j in range(-n, n + 1):
            x, y = px + i * step_ - w / 2, py + j * step_ - h / 2
            dx = max(x - px, 0, px - (x + w))
            dy = max(y - py, 0, py - (y + h))
            cand.append((dx * dx + dy * dy, x, y))
    cand.sort()
    return ((x, y) for _, x, y in cand)


def layout_problems(s):
    """겹침, 선이 상자를 뚫음 등 문제 목록"""
    R = getattr(s.part, "_ixreg", None)
    if not R:
        return []
    out, ob = [], R["obst"]
    for i in range(len(ob)):
        for j in range(i + 1, len(ob)):
            if ob[i][1] > 0 and ob[i][1] == ob[j][1]:
                continue
            a, b = ob[i][0], ob[j][0]
            w = min(a[2], b[2]) - max(a[0], b[0])
            h = min(a[3], b[3]) - max(a[1], b[1])
            if w > 0.02 and h > 0.02:
                out.append(f"겹침 ({a[0]:.2f}, {a[1]:.2f}) 과 ({b[0]:.2f}, {b[1]:.2f})")
    for c in R.get("cols", []):
        if not getattr(c, "_done", False):
            out.append(f"패널 칸 done() 을 부르지 않음 (x {c.x:.2f}). 블록이 그려지지 않음")
    linked = {i for _, a_, b_, _ in R["links"] for i in (a_, b_)}
    for sid, t in R.get("grid_boxes", []):
        if sid not in linked:
            out.append(f"구성도 상자 '{t}' 에 이어진 선이 없음 (선을 잇거나 상자를 뺌)")
    titles = R.get("panel_titles", set())
    for t in R.get("list_titles", []):
        if t in titles:
            out.append(f"목록 제목 '{t}' 이 패널 제목과 같음")
    out.extend(R.get("warn", []))
    for kind, hs_ in R.get("unit_drawn", {}).items():
        if len(hs_) > 1:
            out.append(f"같은 종류 상자({kind}) 높이가 다름 {sorted(hs_)}: 모든 패널에 블록을 넣은 뒤 done() 을 부름")
    for ic, names in R.get("icon_uses", {}).items():
        if len(names) > 1:
            out.append(f"아이콘 '{ic}' 을 두 뜻으로 씀: {sorted(names)}")
    for pts, ia, ib, name in R["links"]:
        for p, q in zip(pts, pts[1:]):
            for r, sid in ob:
                if sid not in (ia, ib, -1) and _seg_hits(p, q, r):
                    out.append(f"선 {name} 이 ({r[0]:.2f}, {r[1]:.2f}) 상자를 지남")
    return out


def _nearest_free(R, pts, w, h, reach=1.2, step_=0.05, strict_only=False):
    """선에서 가장 가까운 빈 글상자 자리"""
    def seg_d(x0, y0, x1, y1):
        best = 9.0
        for (px, py), (qx, qy) in zip(pts, pts[1:]):
            ax0, ax1 = sorted((px, qx))
            ay0, ay1 = sorted((py, qy))
            dx = max(ax0 - x1, 0, x0 - ax1)
            dy = max(ay0 - y1, 0, y0 - ay1)
            best = min(best, (dx * dx + dy * dy) ** .5)
        return best
    xs = [p_[0] for p_ in pts]
    ys = [p_[1] for p_ in pts]
    n = int(reach / step_)
    cand = []
    for i in range(-n, n + 1):
        for j in range(-n, n + 1):
            x = min(xs) + (max(xs) - min(xs)) / 2 + i * step_ - w / 2
            y = min(ys) + (max(ys) - min(ys)) / 2 + j * step_ - h / 2
            cand.append((seg_d(x, y, x + w, y + h), x, y))
    own = set(zip(pts, pts[1:]))
    others = [sg for sg in R["segs"] if sg not in own]

    def near_other(x, y):  # 다른 선 옆은 그 선 번호로 읽혀 멀게 침
        return any(_seg_hits(p_, q_, (x, y, x + w, y + h), pad=-0.2) for p_, q_ in others)
    cand = [(d + (0.4 if near_other(x, y) else 0), x, y) for d, x, y in cand]
    cand.sort()
    strict = next(((d, x, y) for d, x, y in cand if _free(R, (x, y, x + w, y + h), True)), None)
    if strict_only:
        return (strict[1], strict[2]) if strict and strict[0] <= reach else None
    loose = next(((d, x, y) for d, x, y in cand if _free(R, (x, y, x + w, y + h), False)), None)
    if strict and (not loose or strict[0] <= loose[0] + 0.25):
        return strict[1], strict[2]
    return (loose[1], loose[2]) if loose else None


def _flush(s):
    """저장 직전 번호와 선 이름을 빈 자리에"""
    R = getattr(s.part, "_ixreg", None)
    if not R or not R["pending"]:
        return
    todo, R["pending"] = R["pending"], []
    todo.sort(key=lambda it: {"note": 1, "text": 2}.get(it[0], 0))
    for kind, where, t, n, color, cond in todo:
        if kind == "note":
            nw = max([_text_w(t, 7.26) * 1.05] + [_text_w("※ " + l, 6.35) for l in n]) + 0.1
            nh = 0.15 + 0.13 * len(n) + 0.03
            rows_ = R.get("rows", [])

            def aligned(c):
                mine = [c[1] + 0.075] + [c[1] + 0.16 + 0.13 * k + 0.055 for k in range(len(n))]
                for rx0, rx1, ry in rows_:
                    if rx0 < c[0] + nw + 0.8 and c[0] - 0.8 < rx1 and any(0.02 < abs(ry - m) < 0.13 for m in mine):
                        return False
                return True
            x0, y0, x1, y1 = _bb(where)
            cands = itertools.chain(_box_spots(where, nw, nh), _near_spots((x0 + x1) / 2, (y0 + y1) / 2, nw, nh, reach=2.5))
            # 묶음 테두리 안에서 작은 것부터 찾음
            frames = sorted((w_ for w_ in R["walls"] if w_[0] <= x0 and w_[1] <= y0 and w_[2] >= x1 and w_[3] >= y1),
                            key=lambda w_: (w_[2] - w_[0]) * (w_[3] - w_[1])) or [(0, 0, 13.33, 7.5)]
            cands = list(cands)
            sp = None
            for f_ in frames:
                sp = next((c for c in cands if f_[0] + 0.04 <= c[0] and f_[1] + 0.04 <= c[1]
                           and c[0] + nw <= f_[2] - 0.04 and c[1] + nh <= f_[3] - 0.04
                           and aligned(c) and _free(R, (c[0], c[1], c[0] + nw, c[1] + nh), True)), None)
                if sp:
                    break
            if sp is None:
                sp = next(_box_spots(where, nw, nh))
                _warn(s, f"설명 '{t}' 를 놓을 빈 자리가 없어 겹치게 둠")
            _obst(s, text(s, sp[0], sp[1], nw, 0.15, t, size=7.26, bold=True, color=color))
            R.setdefault("rows", []).append((sp[0], sp[0] + nw, sp[1] + 0.075))
            if n:
                _obst(s, text(s, sp[0], sp[1] + 0.16, nw, 0.13 * len(n) + 0.02, ["※ " + l for l in n], size=6.35,
                              color=SLATE9))
            continue
        is_step = kind in ("step", "boxstep")
        lw = _text_w(t or "", 7.26 if is_step else 6.35) + 0.12
        cw = max(lw, _text_w(cond, 6.35) + 0.1) if cond else 0
        w = max(lw, cw) + (0.19 if is_step else 0)
        h = (0.22 + (0.17 if cond else 0)) if is_step else 0.2

        def spots():
            if kind == "boxstep":
                x0, y0, x1, y1 = _bb(where)
                near = lambda: itertools.chain(_box_spots(where, w, h), _near_spots((x0 + x1) / 2, (y0 + y1) / 2, w, h))
                return itertools.chain(((sp, True) for sp in near()), ((sp, False) for sp in near()))
            pts, a, b = where
            return itertools.chain(((sp, True) for sp in _line_spots(pts, w, h)),
                                   ((sp, False) for sp in _line_spots(pts, w, h)))
        spot = None
        if kind != "boxstep":
            spot = next((sp for sp in _line_spots(where[0], w, h) if _free(R, (sp[0], sp[1], sp[0] + w, sp[1] + h), True)),
                        None)
            if spot is None:
                spot = _nearest_free(R, where[0], w, h, reach=0.6, strict_only=True)
        if spot is None:
            spot = next((sp for sp, strict in spots() if _free(R, (sp[0], sp[1], sp[0] + w, sp[1] + h), strict)), None)
        if spot is None and kind != "boxstep":
            pts, a_, b_ = where
            own = set(zip(pts, pts[1:]))
            others = [sg for sg in R["segs"] if sg not in own]
            for sh in (b_, a_):
                for sp in _box_spots(sh, w, h):
                    bb = (sp[0], sp[1], sp[0] + w, sp[1] + h)
                    if _free(R, bb, False) and not any(_seg_hits(p_, q_, bb, pad=-0.15) for p_, q_ in others):
                        spot = sp
                        break
                if spot:
                    break
        if spot is None and kind != "boxstep":
            pts = where[0]
            spot = _nearest_free(R, pts, w, h)
        if spot is None:
            spot = next(spots())[0]
            _warn(s, f"'{t}' 를 놓을 빈 자리가 없어 겹치게 둠")
        x, y = spot
        if is_step:
            step(s, x, y + 0.03, n, t, color=color, label_w=lw)
            if cond:
                label(s, x + 0.19, y + 0.22, cw, cond, color=RED, align="l")
        else:
            label(s, x, y, lw, t, color=color, bold=True)


class DocGrid:
    """설계문서형 묶음 안 자동 배치 (열, 줄)"""
    PAD = 0.25       # 묶음 안 좌우 여백
    INNER = 0.2      # area 안 좌우 여백

    def __init__(self, s, right=None, rows=4, box_h=0.24, top=None, bottom=DOC_BOTTOM):
        if top is None:
            top = _reg(s).get("doc_top", DOC_TOP)
        self.s, self.cols, self.n, self.bh, self.top, self.bottom = s, doc_columns(right), rows, box_h, top, bottom
        first = top + self._head(True, True)
        last = bottom - self._foot(True)
        self.ys = [first + (last - first) * i / (rows - 1) for i in range(rows)] if rows > 1 else [first]
        self._used = {}

    def _head(self, header, area):
        """묶음 위에서 첫 줄 상자 가운데까지"""
        return 0.27 + 0.27 + (0.45 if header else 0) + (0.15 if area else 0) + self.bh / 2

    def _foot(self, area):
        return self.bh / 2 + (0.15 if area else 0) + 0.25

    def y(self, r):
        return self.ys[r]

    def group(self, col, title, kind="platform", sub=None, rows=None, header=None, cols=1, area=True, badge=None):
        if col == "edge" and not getattr(self, "_has_source", False) and "source" in self.cols:
            # 원천 열이 없으면 edge 열을 넓힘
            sx, _ = self.cols["source"]
            ex, ew = self.cols["edge"]
            self.cols["edge"] = (sx, ex + ew - sx)
            self.cols.pop("source")
        x, w = self.cols[col]
        r0, r1 = rows or (0, self.n - 1)
        if not (0 <= r0 <= r1 < self.n) or cols < 1:
            raise ValueError(f"rows={rows}, cols={cols} 가 격자(줄 {self.n})를 벗어남")
        top = self.top if r0 == 0 else self.ys[r0] - self._head(header, area)
        bottom = self.bottom if r1 == self.n - 1 else self.ys[r1] + self._foot(area)
        for t0, b0 in self._used.get(col, []):
            if top < b0 + 0.15 and bottom > t0 - 0.15:
                raise ValueError(f"{col} 열 묶음 '{title}' 이 위아래 묶음과 겹치거나 0.15 in 보다 가까움: "
                                 "줄 사이를 한 줄 더 비우거나(rows 늘림), 아래 묶음을 header=None, area=False 로")
        self._used.setdefault(col, []).append((top, bottom))
        g_ = _Group(self, x, w, top, bottom, title, kind, sub, header, cols, use_area=area)
        if badge and g_.header is not None:
            hx, hy, _, hy1 = _bb(g_.header)
            hh = hy1 - hy
            bh_ = hh + 0.2
            pic = self.s.shapes.add_picture(_icon_path(badge), Inches(hx + 0.06), Inches(hy - 0.1), height=Inches(bh_))
            _reg(self.s)["obst"].append((_bb(pic), g_.header.shape_id))
        return g_

    def bar(self, col, t, fill=NAVY, tag_=None):
        """열 아래 가로 띠 (Data Lake)"""
        x, w = self.cols[col]
        b = _obst(self.s, box(self.s, x, self.bottom + 0.1, w, 0.24, t, fill=fill, size=7.26, bold=True, color=WHITE))
        if tag_:
            tag(self.s, x + w / 2 - 0.115, self.bottom + 0.1 - 0.17, tag_)
        return b

    def stack(self, col, title, sub, n=2, dots=True):
        """겹친 띠 + ⋮"""
        x, w = self.cols[col]
        stack_bars(self.s, x, self.bottom + 0.1, w, title, sub, n=n, pitch=0.32, dots=dots)

    def source(self, col, items, title="Data Source", icon_w=0.4):
        """Data Source 열. items=[(아이콘, 이름)]"""
        if not items:
            raise ValueError("Data Source 항목이 없음")
        if len(items) > 6:
            _warn(self.s, f"Data Source {len(items)}개는 이름이 겹침. 6개 이하로 두고 more() 로 줄임")
        if col == "source" and "source" not in self.cols:
            raise ValueError("source() 는 edge 묶음보다 먼저 부른다")
        self._has_source = self._has_source if hasattr(self, "_has_source") else False
        if col == "source":
            self._has_source = True
        x, w = self.cols[col]
        body = frame(self.s, x, self.top, w, self.bottom - self.top, title, kind="source")
        y0, y1 = self.top + 0.55, self.bottom - 0.55
        step_ = (y1 - y0) / len(items)
        icon_w = min(icon_w, step_ - 0.3)
        self.icons = [icon(self.s, ic, x + w / 2 - icon_w / 2, y0 + step_ * i + (step_ - icon_w - 0.2) / 2, w=icon_w,
                           name=n) for i, (ic, n) in enumerate(items)]
        return body


class _Group:
    def __init__(self, g, x, w, top, bottom, title, kind, sub, header, cols, use_area):
        self.g, self.s = g, g.s
        self.body = frame(self.s, x, top, w, bottom - top, title, kind=kind, sub=sub)
        ix, iw = x + g.PAD, w - 2 * g.PAD
        y = top + 0.54
        self.header = None
        if header:
            self.header = server(self.s, ix, y, iw, 0.27, header)
            y += 0.45
        self.area = area(self.s, ix, y, iw, bottom - 0.25 - y) if use_area else None
        self.ix, self.iw = (ix + g.INNER, iw - 2 * g.INNER) if use_area else (ix, iw)
        self.cw = self.iw / cols
        self.ncols = cols

    def box(self, r, c, t, kind="db", span=1, w=None, icon_=None):
        """줄 r, 칸 c 에 상자. kind: db, server"""
        kw_w = w is not None
        if not (0 <= r < self.g.n) or not (0 <= c and c + span <= self.ncols):
            raise ValueError(f"'{t}' 자리 (줄 {r}, 칸 {c}, span {span}) 가 묶음(줄 {self.g.n}, 칸 {self.ncols})을 벗어남")
        if w is None:  # 상자 사이 0.42: 선 세로 구간과 번호 자리
            base = min(1.4, self.cw - 0.42) if span == 1 else self.cw * span - 0.42
            need = _text_w(t, 7.26) * 1.1 + 0.2
            w = min(max(base, need), self.cw * span - 0.08)
            if need > self.cw * span - 0.08:
                _warn(self.s, f"'{t}' 가 칸보다 김 (칸 {self.cw * span:.2f} in). cols 를 줄이거나 묶음을 넓힘")
        bw = w
        cx = self.ix + self.cw * (c + span / 2)
        y = self.g.y(r) - self.g.bh / 2
        fn = db if kind == "db" else server
        if icon_:
            bw = min(bw, self.cw * span - 0.5)
            cx -= 0.2
        b = fn(self.s, cx - bw / 2, y, bw, self.g.bh, t)
        _reg(self.s).setdefault("grid_boxes", []).append((b.shape_id, t))
        pic = None
        if icon_:
            ih = self.g.bh + 0.12
            pic = icon(self.s, icon_, cx + bw / 2 + 0.08, y - 0.06, h=ih)
        if not kw_w:
            _col_box(self.s, (id(self), c, span, bool(icon_)), b, self.cw * span - (0.5 if icon_ else 0.42), pic)
        return b

    def sub(self, c, r0, r1, header=None, span=1, caption=None):
        """묶음 안 작은 묶음 (칸 c, 줄 r0 ~ r1)"""
        if not (0 <= r0 <= r1 < self.g.n) or not (0 <= c and c + span <= self.ncols):
            raise ValueError(f"sub 자리 (칸 {c}, span {span}, 줄 {r0} ~ {r1}) 가 묶음을 벗어남")
        g = self.g
        x0, w = self.ix + self.cw * c + 0.06, self.cw * span - 0.12
        out = _Sub(self, x0, w)
        top_r = r0
        if header:
            out.header = server(self.s, x0, g.y(r0) - g.bh / 2 - 0.015, w, g.bh + 0.03, header)
            top_r = r0 + 1
            if top_r > r1:
                return out
        ay0 = g.y(top_r) - g.bh / 2 - 0.12
        ay1 = g.y(r1) + g.bh / 2 + 0.12
        out.area = area(self.s, x0, ay0, w, ay1 - ay0)
        if caption:
            label(self.s, x0, g.y(top_r) - 0.1, w, caption)
        return out

    def zone(self, r0, r1, c0, c1, color=ACCENT, alpha=0.07):
        """줄, 칸 범위를 덮는 옅은 구역"""
        r0, r1, c0, c1 = min(r0, r1), max(r0, r1), min(c0, c1), max(c0, c1)
        y0 = self.g.y(r0) - self.g.bh / 2 - 0.12
        y1 = self.g.y(r1) + self.g.bh / 2 + 0.12
        x0, x1 = self.ix + self.cw * c0 + 0.05, self.ix + self.cw * (c1 + 1) - 0.05
        return zone(self.s, x0, y0, x1 - x0, y1 - y0, color=color, alpha=alpha, above=self.area or self.body)


class _Sub:
    """작은 묶음. box(줄, 글)"""
    def __init__(self, grp, x, w):
        self.grp, self.s, self.x, self.w, self.header, self.area = grp, grp.s, x, w, None, None

    def box(self, r, t, kind="db", w=None):
        g = self.grp.g
        bw = w or min(self.w - 0.24, max(1.0, _text_w(t, 7.26) * 1.1 + 0.2))
        fn = db if kind == "db" else server
        b = fn(self.s, self.x + (self.w - bw) / 2, g.y(r) - g.bh / 2, bw, g.bh, t)
        _reg(self.s).setdefault("grid_boxes", []).append((b.shape_id, t))
        if not w:
            _col_box(self.s, (id(self),), b, self.w - 0.24)
        return b


def note_by(s, shape, head, lines=()):
    """상자 옆 빈 자리에 설명 (head, lines)"""
    _reg(s)["pending"].append(("note", shape, head, list(lines), RED, None))


def step_by(s, shape, n, t=None, where="auto", color=ACCENT, cond=None, dx=0.0):
    """상자 옆에 번호 원 자동 배치"""
    _before_draw(s)
    if where == "auto":
        _reg(s)["pending"].append(("boxstep", shape, t, n, color, cond))
        return None
    E = 914400
    x, y, w, h = shape.left / E, shape.top / E, shape.width / E, shape.height / E
    lw = _text_w(t or "", 7.26) + 0.12
    px, py = {"below": (x + 0.02, y + h + 0.09), "above": (x + 0.02, y - 0.24),
              "below_r": (x + w / 2 + 0.06, y + h + 0.09), "above_r": (x + w / 2 + 0.06, y - 0.24),
              "right": (x + w + 0.08, y + h / 2 - 0.075), "left": (x - 0.27 - lw, y + h / 2 - 0.075)}[where]
    c = step(s, px + dx, py, n, t, color=color, label_w=lw)
    if cond:
        label(s, px + dx + 0.19, py + 0.19, max(lw, _text_w(cond, 6.35) + 0.1), cond, color=RED, align="l")
    return c


def stack_bars(s, x, y, w, title, sub, n=2, pitch=0.6, dots=True):
    """겹친 띠. 대수를 알면 dots=False"""
    for i in range(n):
        _obst(s, box(s, x, y + i * pitch, w, 0.24, [(title, BLACK), (sub, ACCENT)], fill=SLATE3, size=6.35, bold=True,
                     inset=(0.02, 0, 0.02, 0)))
    if dots:
        more(s, x + w / 2, y + n * pitch - 0.08)


# --- 덱
def _apply_ko_breaks(sl):
    """글 조각 하나인 문단에 ko_breaks 적용"""
    def frames():
        for sh in sl.shapes:
            if sh.is_placeholder:
                continue
            if sh.has_text_frame and sh.text_frame.word_wrap is not False:
                tf = sh.text_frame
                yield tf, sh.width / _E - (tf.margin_left + tf.margin_right) / _E
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for k, cell in enumerate(row.cells):
                        w = sh.table.columns[k].width / _E - (cell.margin_left + cell.margin_right) / _E
                        yield cell.text_frame, w
    for tf, w in frames():
        for para in tf.paragraphs:
            runs = para.runs
            if len(runs) != 1 or para._p.find(qn("a:br")) is not None:
                continue
            r = runs[0]
            txt = r.text
            sz = r.font.size.pt if r.font.size else 7.26
            ind = 0.15 if (para._p.pPr is not None and para._p.pPr.find(qn("a:buChar")) is not None) else 0
            lines = ko_breaks(txt, w - ind, sz)
            if len(lines) < 2:
                continue
            r.text = lines[0]
            prev = r._r
            for ln in lines[1:]:
                br = etree.SubElement(para._p, qn("a:br"))
                rpr = r._r.find(qn("a:rPr"))
                if rpr is not None:
                    br.append(copy.deepcopy(rpr))
                prev.addnext(br)
                nr = copy.deepcopy(r._r)
                nr.find(qn("a:t")).text = ln
                br.addnext(nr)
                prev = nr


PROFILE_DIR = os.path.join(os.environ.get("CLAUDE_CONFIG_DIR") or os.path.join(os.path.expanduser("~"), ".claude"),
                           "interx-deck", "profile")   # 개인 설정 (플러그인 밖)


def _template_fresh(template):
    """회사 양식 원본(template_source)이 더 새로우면 알림"""
    import json
    p = os.path.join(PROFILE_DIR, "profile.json")
    src = json.load(open(p, encoding="utf8")).get("template_source") if os.path.exists(p) else None
    if src and os.path.exists(src) and os.path.getmtime(src) > os.path.getmtime(template):
        print(f"주의: 회사 양식 원본이 템플릿보다 새로움. python -P {os.path.join(HERE, 'make_template.py')} "
              f'"{src}" "{template}" 로 다시 만듦')


def profile_author():
    """profile.json 작성자. 없으면 None"""
    import json
    p = os.path.join(PROFILE_DIR, "profile.json")
    if not os.path.exists(p):
        return None
    a = json.load(open(p, encoding="utf8")).get("author") or {}
    if not a.get("name"):
        return None
    return f"{a.get('team', '').strip()} / {a['name'].strip()} {a.get('title', '').strip()}".strip(" /")


def save_profile_author(team, name, title):
    """작성자를 개인 설정에 저장"""
    import json
    os.makedirs(PROFILE_DIR, exist_ok=True)
    p = os.path.join(PROFILE_DIR, "profile.json")
    d_ = json.load(open(p, encoding="utf8")) if os.path.exists(p) else {}
    d_["author"] = {"team": team, "name": name, "title": title}
    json.dump(d_, open(p, "w", encoding="utf8"), ensure_ascii=False, indent=2)
    return p


class Deck:
    """템플릿에 새 장 추가"""

    def __init__(self, author, template=TEMPLATE, parts_only=False, forbid=(), allow=()):
        """parts_only: 본문만. forbid: 덱에 나가면 안 되는 정규식
        - allow: 검사에서 뺄 정규식 (화면 문구 그대로 옮긴 글 등)"""
        self.parts_only = parts_only
        self.forbid = [re.compile(p) for p in forbid]
        self.allow = list(allow)
        for p in self.allow:   # 빈 글에 맞는 정규식은 모든 글을 지움
            if not p or re.compile(p).fullmatch(""):
                raise ValueError(f"allow 정규식 '{p}' 이 빈 글에도 맞음: 범위를 좁힘")
        if not author or not author.strip():
            raise ValueError("작성자 정보가 없음: 사용자에게 먼저 물어 Deck(author=...) 로 넘김")
        self.author = author.strip()
        _template_fresh(template)
        self.prs = Presentation(template)
        self.prs.core_properties.author = self.author
        self.prs.core_properties.last_modified_by = self.author
        if self.allow:   # 검사 도구가 읽음
            self.prs.core_properties.keywords = "ix-allow:" + __import__("json").dumps(self.allow, ensure_ascii=False)
        self._lib = list(self.prs.slides._sldIdLst)
        self.layouts = {l.name: l for l in self.prs.slide_masters[0].slide_layouts}

    def filename(self, title, version, folder=".", date=None):
        """규칙 이름 제목_작성자_v버전_YYMMDD.pptx (띄어쓰기는 밑줄)"""
        last = self.author.split("/")[-1].strip()
        if "/" not in self.author or not last:
            raise ValueError(f"작성자 '{self.author}' 는 '소속 / 이름 직급' 꼴이어야 파일 이름을 만듦")
        name = last.split()[0]
        title = re.sub(r"[\s_]+", "_", re.sub(r'[\\/:*?"<>|]+', "", title)).strip("_")
        if not title:
            raise ValueError("파일 제목이 비어 있음")
        day = (date or datetime.date.today()).strftime("%y%m%d")
        return os.path.join(folder, f"{title}_{name}_v{str(version).lstrip('v')}_{day}.pptx")

    def lib(self, n):
        """템플릿 n 번째 장 (부품 원본)"""
        return self.prs.slides[n - 1]

    # 표지, 목차, 본문, 끝 장 레이아웃만 씀
    LAYOUTS = ("3_사용자 지정 레이아웃", "5_contents-02", "본문_소제목 O", "앤드페이지_다크네이비")

    def _add(self, layout, keep=()):
        if layout not in self.LAYOUTS:
            raise ValueError(f"레이아웃 '{layout}' 금지: 표지는 cover(), 본문은 모두 '본문_소제목 O'")
        s = self.prs.slides.add_slide(self.layouts[layout])
        for ph in list(s.placeholders):
            if ph.placeholder_format.idx not in keep:
                ph._element.getparent().remove(ph._element)
        return s

    def cover(self, title, date):
        """표지. 날짜는 덱 기준일 (자동 아님)"""
        if not date:
            raise ValueError("cover: 덱 기준일을 넣는다 (예 '2026.11.04')")
        if sum(_em(c) for c in title) > 16.5:   # 실측: 한 줄은 한글 약 15자
            raise ValueError(f"표지 제목 '{title}' 이 한 줄을 넘어 날짜를 가림: 줄임")
        s = self._add("3_사용자 지정 레이아웃", keep=(10, 11))
        s.placeholders[10].text_frame.text = title
        s.placeholders[11].text_frame.text = date
        tb = self.clone(s, 1, text="AX Data Design / 데이터 표준화")
        tb.text_frame.paragraphs[0].runs[0].text = self.author
        for r in tb.text_frame.paragraphs[0].runs[1:]:
            r.text = ""
        if self.parts_only:
            raise ValueError("parts_only 덱은 표지를 만들지 않음 (본문 장만 옮겨 넣는 빌더용)")
        self._toc = self._add("5_contents-02")
        self._toc_items = []
        return s

    def _fill_toc(self):
        """목차: 절 번호, 절 이름, 장 주제"""
        items = getattr(self, "_toc_items", [])
        if not items:
            raise ValueError("목차에 넣을 절이 없음: d.section() 과 d.page() 로 장을 만든다")
        s = self._toc
        cols = 1 if len(items) <= 5 else 2
        x0s, w = ((0.75,), 6.6) if cols == 1 else ((0.75, 4.55), 3.6)
        per = math.ceil(len(items) / cols)
        top, bottom = 1.55, 6.75
        hs = [0.42 + 0.22 * len(pages) for _, pages in items]
        for c in range(cols):
            chunk = list(enumerate(items))[c * per:(c + 1) * per]
            need = sum(hs[i] for i, _ in chunk)
            gap = min(0.35, max(0.12, (bottom - top - need) / max(len(chunk) - 1, 1)))
            y = top
            if need > bottom - top:
                _warn(s, f"목차가 칸을 넘침 ({need:.2f} in, 칸 {bottom - top:.2f} in): 절을 합치거나 장 주제를 줄임")
            for i, (name, pages) in chunk:
                if _text_w(name, 14) > w - 0.7 or any(_text_w(pg, 9) > w - 0.7 for pg in pages):
                    _warn(s, f"목차 '{name}' 의 글이 칸보다 김: 절 이름과 장 주제를 짧게")
                text(s, x0s[c], y, 0.6, 0.36, f"{i + 1:02d}", size=20, bold=True, color=ORANGE)
                text(s, x0s[c] + 0.65, y, w - 0.65, 0.36, name, size=14, bold=True, color=BLACK)
                for k, pg in enumerate(pages):
                    text(s, x0s[c] + 0.67, y + 0.42 + 0.22 * k, w - 0.67, 0.2, pg, size=9, color=SLATE9)
                y += hs[i] + gap
                box(s, x0s[c], y - gap / 2, w, 0.01, fill=GRAY_L) if i != chunk[-1][0] else None

    def _end(self):
        return self._add("앤드페이지_다크네이비")

    def report_slide(self, title, subtitle, lead=None):
        """보고형 머리 띠와 요약 한 줄"""
        s = self._add("본문_소제목 O", keep=(10, 11))
        for idx, t in ((10, title), (11, subtitle)):
            ph = s.placeholders[idx]
            tf = ph.text_frame
            tf.text = t
            for r in tf.paragraphs[0].runs:
                r.font.bold = True
            # 자리표시자가 좁아 접힘: 장식 앞까지 넓힘
            l_, t_, h_ = ph.left, ph.top, ph.height   # 물려받은 값. 폭만 바꾸면 위치가 (0, 0)
            ph.left, ph.top, ph.height = l_, t_, h_
            ph.width = Inches(8.6 if idx == 10 else 11.8) - l_
        if lead:
            summary(s, lead)
        return s

    def section(self, name):
        """절 시작. 이후 page() 제목에 번호"""
        self._sec = getattr(self, "_sec", 0) + 1
        self._sec_name = name
        self._page = 0
        getattr(self, "_toc_items", []).append((name, []))

    def page(self, subtitle, lead=None, doc=False):
        """다음 장. doc=True 면 구성도 장"""
        if not getattr(self, "_sec", 0):
            raise ValueError("page: 먼저 d.section('절 이름')")
        if getattr(self, "_closed", False):
            raise ValueError("저장한 뒤 장을 더하지 않는다 (목차와 Thank you 가 이미 들어감): build.py 를 처음부터 다시 돌린다")
        self._page += 1
        if getattr(self, "_toc_items", None):
            self._toc_items[-1][1].append(subtitle.split(" (")[0])
        title = f"{self._sec}. {self._sec_name}"
        sub = f"{self._sec}-{self._page}. {subtitle}"
        return (self.doc_slide if doc else self.report_slide)(title, sub, lead=lead)

    def doc_slide(self, title, subtitle, lead=None):
        """설계문서형 장 (넓은 도식 자리)"""
        s = self.report_slide(title, subtitle, lead=lead)
        _reg(s)["doc_top"] = REPORT_TOP if lead else 1.45
        return s

    # 템플릿 14장 AAS 그림의 도형 번호
    _SHELL = {"frame": 19, "title": 6, "chips": (45, 7, 8, 9, 10, 11, 12),
              "boxes": (16, 17, 31, 35, 13, 14, 15), "more": 18, "tab": 2, "pic": 20}
    _SHELL_PITCH = 0.2765

    def aas_shell(self, s, x, y, sms, scale=0.75, more=True, focus_=None, widen=0.0):
        """템플릿 AAS 그림 복제"""
        if not 1 <= len(sms) <= 7:
            raise ValueError("aas_shell: SM 은 1 ~ 7개")
        S = self._SHELL
        cut = (7 - len(sms)) * self._SHELL_PITCH
        lib = self.lib(14)
        by_id = {sh.shape_id: sh for sh in lib.shapes}
        fr = by_id[S["frame"]]
        ax, ay = fr.left, fr.top
        made, boxes = [], {}
        keep_ids = [S["frame"], S["title"], S["tab"], S["pic"]] + list(S["chips"][:len(sms)]) + \
            list(S["boxes"][:len(sms)]) + ([S["more"]] if more else [])
        for sid in keep_ids:
            sh = self.clone(s, 14, shape_id=sid)
            if sid == S["frame"]:
                sh.height = sh.height - Inches(cut)
            if sid in (S["tab"], S["pic"], S["more"]):
                sh.top = sh.top - Inches(cut)
            if widen:  # 긴 SM 이름: 틀을 넓히고 나머지는 가운데로
                if sid in (S["frame"], S["title"]) or sid in S["boxes"]:
                    sh.width = sh.width + Inches(widen)
                elif sid in (S["tab"], S["pic"], S["more"]):
                    sh.left = sh.left + Inches(widen / 2)
            made.append(sh)
            if sid in S["boxes"]:
                name = sms[S["boxes"].index(sid)]
                sh.text_frame.paragraphs[0].runs[0].text = name
                sh.text_frame.margin_left = sh.text_frame.margin_right = Inches(0.03)
                for r in sh.text_frame.paragraphs[0].runs[1:]:
                    r.text = ""
                boxes[name] = sh
        k = scale
        for sh in made:
            sh.left = Inches(x) + int((sh.left - ax) * k)
            sh.top = Inches(y) + int((sh.top - ay) * k)
            sh.width, sh.height = int(sh.width * k), int(sh.height * k)
            for rpr in sh._element.iter(qn("a:rPr"), qn("a:endParaRPr"), qn("a:defRPr")):
                if rpr.get("sz"):
                    rpr.set("sz", str(int(round(_snap_pt(int(rpr.get("sz")) * k / 100) * 100))))
        if focus_ in boxes:
            focus(boxes[focus_])
        bb = [_bb(sh) for sh in made]
        box_ = (min(b[0] for b in bb), min(b[1] for b in bb), max(b[2] for b in bb), max(b[3] for b in bb))
        _reg(s)["obst"].append((box_, 0))
        return {"boxes": boxes, "bbox": box_}

    def clone(self, s, src_slide, shape_id=None, text=None, dx=0.0, dy=0.0):
        """템플릿 장의 도형을 복사 (shape_id, text)"""
        src_s = self.lib(src_slide)
        hits = [sh for sh in src_s.shapes if (shape_id is not None and sh.shape_id == shape_id)
                or (text is not None and sh.has_text_frame and sh.text_frame.text.strip() == text)]
        if len(hits) != 1:
            raise ValueError(f"{src_slide} 장에서 {shape_id or text!r} 도형이 {len(hits)}개")
        src = hits[0]
        el = copy.deepcopy(src._element)
        nid = max(int(e.get("id")) for e in s.shapes._spTree.iter(qn("p:cNvPr"))) + 1
        for e in el.iter(qn("p:cNvPr")):
            e.set("id", str(nid))
            nid += 1
        for blip in el.iter(qn("a:blip")):
            rid = blip.get(qn("r:embed"))
            if rid:
                part = src_s.part.related_part(rid)
                blip.set(qn("r:embed"), s.part.relate_to(part, RT.IMAGE))
        s.shapes._spTree.append(el)
        new = s.shapes[-1]
        new.left, new.top = new.left + Inches(dx), new.top + Inches(dy)
        return new

    def save(self, path, strict=True):
        """저장. strict 면 문제가 있을 때 멈춤"""
        if os.environ.get("INTERX_OUT_DIR"):   # 회귀 시험: 원래 결과물을 덮지 않음
            path = os.path.join(os.environ["INTERX_OUT_DIR"], os.path.basename(path))
        if not hasattr(self, "_toc") and not self.parts_only:
            raise ValueError("표지가 없음: d.cover(제목, 날짜) 로 시작 (표지 뒤 목차가 따라 붙음)")
        if not getattr(self, "_closed", False) and not self.parts_only:
            self._fill_toc()
            self._end()
            self._closed = True
        lib_n = 0 if getattr(self, "_saved", False) else len(self._lib)
        problems = []
        body = [sl for sl in list(self.prs.slides)[lib_n:] if sl.slide_layout.name == "본문_소제목 O"]
        kpi_slides = [i for i, sl in enumerate(body, 1) if _reg(sl).get("kpi_items", 0)]
        for i, sl in enumerate(body, 1):
            if _reg(sl).get("kpi_rows", 0) > 1 or _reg(sl).get("kpi_items", 0) > 4:
                problems.append(f"본문 {i}장: 핵심 수치 줄 {_reg(sl).get('kpi_rows', 0)}개, 수치 {_reg(sl).get('kpi_items', 0)}개 "
                                "(한 장에 한 줄, 4개까지)")
        seen = {}
        for i, sl in enumerate(body, 1):
            for t in set(_reg(sl).get("tab_labels", [])):
                seen.setdefault(t, []).append(i)
        for t, at in seen.items():
            if len(at) >= 3:   # 장마다 같은 이름표는 틀로 찍은 느낌
                problems.append(f"이름표 '{t}' 가 본문 {len(at)}장에 되풀이 ({at}): 장마다 내용을 담은 명사로")
        if len(kpi_slides) > max(1, len(body) // 5):
            problems.append(f"핵심 수치 줄을 쓴 본문 장 {len(kpi_slides)}장 (본문 {len(body)}장의 1/5 인 "
                            f"{max(1, len(body) // 5)}장까지): 숫자가 메시지인 장에만")
        for i, sl in enumerate(self.prs.slides, 1):
            _flush(sl)
            if i > lib_n:
                for m in dict.fromkeys(layout_problems(sl)):
                    print(f"주의 {i - lib_n}장: {m}")
                    problems.append(f"{i - lib_n}장: {m}")
                alt = sl._element.xpath(".//p:cNvPr/@descr")   # 그림 대체 텍스트 (파일 이름)
                for t in list(slide_texts(sl)) + alt:
                    for p in self.forbid:
                        if p.search(t):
                            print(f"주의 {i - lib_n}장: 금지 문구 '{p.search(t).group()}'")
                            problems.append(f"{i - lib_n}장: 금지 문구 '{p.search(t).group()}' (forbid)")
        if strict and problems:
            raise ValueError("덱 문제 " + str(len(problems)) + "건, 저장하지 않음:\n" + "\n".join(problems))
        if not getattr(self, "_saved", False):
            lst = self.prs.slides._sldIdLst
            for sid in self._lib:
                self.prs.part.drop_rel(sid.get(qn("r:id")))
                lst.remove(sid)
            self._saved = True
        for sl in list(self.prs.slides)[(0 if getattr(self, "_saved_once", False) else 0):]:
            _apply_ko_breaks(sl)
        self._saved_once = True
        self.prs.save(path)
        normalize_fonts(path)
        return path
