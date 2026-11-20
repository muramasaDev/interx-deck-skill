"""INTERX 양식 PPT 를 도형으로 그리는 엔진 (인치)"""
from __future__ import annotations

import copy
import datetime
import itertools
import math
import os
import re
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from check_ban import SRC_LINE, slide_texts

HERE = os.path.dirname(os.path.realpath(__file__))
ASSETS = os.path.join(os.path.dirname(HERE), "assets")
TEMPLATE = os.path.join(ASSETS, "interx_template.pptx")
ICONS = os.path.join(ASSETS, "icons")
LOGOS = os.path.join(ASSETS, "logos")
FONT = "Pretendard"
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
# 기능 묶음 색 (머리띠 색, 옅은 바탕): 같은 기능은 같은 색, 한 장에 넷 이하. collect 수집, store 저장, send 전송, manage 관리, report 현황 및 알림
TONES = {"collect": ("1F4E8C", "EAF0F8"), "store": ("546680", "EFF1F5"), "send": ("299288", "E6F3F1"),
         "manage": ("6B4E8F", "F0EBF5"), "report": ("A86B12", "FBF3E3")}
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
# 숫자 및 단위 사이 공백 (줄 갈림 방지용 NBSP 로 바꿀 자리)
_NUM_UNIT = re.compile(r"(?<=\d) (?=(?:%|℃|[µμ]m|mm|cm|kg|ms|pt|in|m|s|시간|초|분|일|개|건|회|명|대|장|원|배|행|열|줄|월|년)"
                       r"(?![A-Za-z0-9]))")
STAMP_X, STAMP_Y = 12.55, 0.2   # 도장 오른쪽 끝, 위
PURPOSES = ("사내보고", "회의", "고객", "파트너", "문서", "영업", "제안")
# 블록 글자 한도 (넘으면 저장 실패). 3열 기준 칸 폭의 값이고 칸 폭에 비례
LIMITS = {"num_cards 제목": 16, "num_cards 설명": 40, "tab_box 줄": 46, "kpis 이름": 12, "section 행": 30,
          "rule_list 글": 30}
KPI_HEAD = ["지표", "기존 수치", "목표 수치", "계산식", "근거"]
COLLECT_HEAD = ["공정", "데이터", "집계 방법"]
COLLECT_METHODS = ("자동", "반자동", "수동")
# 회사 방법론 (skills/design/references/A1_방법론.md)
DAS_HEAD = ["구분", "Data", "Data Type", "Range", "Acq. Time", "Unit", "Stage", "Connected"]
DAS_KIND = re.compile(r"^(X[1-5]|Y\d*|Yn)$")
QUALITY_W = {"완전성 (누락)": 60, "유일성 (중복)": 10, "유효성 (유효)": 10, "일관성 (표현)": 10, "정확성 (실제)": 10}
QUALITY_TARGET = 95.0   # 제조 데이터 품질 95% 이상 필수
PRIORITY_W = ((0.4, 0.3, 0.3), (0.4, 0.4, 0.2))   # 전략적 중요도, 실행의 용이성
PRIORITY_HEAD = ["과제", "기여도", "효과 크기", "시급성", "중요도", "투자", "효과 발생", "관리", "용이성", "총점", "구분"]
LEVELS = ["L1 ~ 2 모니터링", "L3 제어", "L4 최적화", "L5 자율운영"]   # L: Level


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
                for t_ in root.iter(qn("a:t")):   # 숫자 및 단위는 줄이 갈리지 않게
                    if t_.text and " " in t_.text:
                        t_.text = _NUM_UNIT.sub("\u00a0", t_.text)
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


def rule_list(s, x, y, w, title, items, pitch=0.29, color=ACCENT):
    """아이콘 + 굵은 글 목록, 왼쪽 위 남색 이름표. 바닥 y 반환"""
    if any(len(it) > 2 for it in items):
        pitch = max(pitch, 0.38)
    h = 0.27 + pitch * len(items) + 0.06
    box(s, x, y + 0.11, w, h - 0.11, fill=None, line=GRAY_L, lw=0.75)
    box(s, x, y, min(w, _text_w(title, 9) + 0.3), 0.22, title, fill=color, size=9, bold=True, color=WHITE,
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


def stamp(s, text="예시"):
    """장 오른쪽 위 기운 빨간 도장 (한 장에 하나)"""
    if not text or not text.strip():
        raise ValueError("stamp: 도장 글이 비어 있음")
    if _reg(s).get("stamp"):
        raise ValueError("stamp: 한 장에 하나만")
    w = max(0.9, _text_w(text, 9) + 0.3)
    sh = box(s, STAMP_X - w, STAMP_Y, w, 0.26, text, fill=None, line=RED, lw=1.5, size=9, bold=True, color=RED,
             name="도장")
    sh.rotation = -8
    _reg(s)["stamp"] = True
    return sh


def notes(s, x, y, w, lines, size=7.26, bold=False):
    """'※' 메모 줄"""
    return text(s, x, y, w, _notes_h(lines, w, size), ["※ " + l for l in lines], size=size, bold=bold)


def _notes_h(lines, w, size):
    n = sum(max(1, math.ceil(_text_w("※ " + l, size) / w)) for l in lines)
    return max(0.15, size * 1.3 / 72) * n + 0.05


def big_arrow(s, x, y, w=0.38, h=0.33, direction="right", colors=None, adj=(0.48473, 0.52219)):
    """패널 사이 회색 큰 화살표. colors=(꼬리 색, 머리 색) 이면 그 두 색의 그라데이션"""
    rot = {"right": 0, "down": 90, "left": 180, "up": 270}[direction]
    sh = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    _no_style_effects(sh)
    _adj(sh, *adj)
    sh.rotation = rot
    sppr = sh._element.spPr
    for el in sppr.findall(qn("a:solidFill")):
        sppr.remove(el)
    stops, ang = (f'<a:gs pos="0"><a:srgbClr val="{colors[0]}"/></a:gs><a:gs pos="100000"><a:srgbClr val="{colors[1]}"/></a:gs>',
                  0) if colors else (
        '<a:gs pos="0"><a:schemeClr val="bg1"><a:lumMod val="50000"/></a:schemeClr></a:gs>'
        '<a:gs pos="50000"><a:schemeClr val="bg1"><a:lumMod val="75000"/></a:schemeClr></a:gs>'
        '<a:gs pos="100000"><a:schemeClr val="bg1"><a:lumMod val="95000"/></a:schemeClr></a:gs>', 10800000)
    grad = etree.fromstring(f'<a:gradFill xmlns:a="{_A}" flip="none" rotWithShape="1"><a:gsLst>{stops}</a:gsLst>'
                            f'<a:lin ang="{ang}" scaled="1"/></a:gradFill>')
    geom = sppr.find(qn("a:prstGeom"))
    geom.addnext(grad)
    _line(sh, None)
    return sh


CAP_H = 0.22   # 표 제목 및 그림 제목 줄 높이
_NUM_CELL = re.compile(r"^[+\-]?\d[\d,]*(\.\d+)?$")   # 단위 없는 맨 숫자 칸 (날짜, 버전은 점이 둘이라 제외)
_YEAR = re.compile(r"^(19|20)\d\d$")


def _caption(s, kind, title, x, y, w, align, key=None):
    """[표 #. 제목] 또는 [그림 #. 제목]. 번호는 저장 때 key 순서로 (덱 전체 연번)"""
    if not title or not str(title).strip():
        _warn(s, f"{kind} 제목 없음: title='{kind} 제목' ({kind} {'왼쪽 위' if kind == '표' else '아래 가운데'}에 [{kind} 번호. 제목])")
        title = "제목 없음"
    t_ = f"[{kind} #. {title}]"
    need = _text_w(t_, 8) + 0.1
    if need > w:
        if align == "c":   # 그림 제목: 가운데를 지키며 넓힘
            x, w = x + w / 2 - need / 2, need
        else:
            _warn(s, f"{kind} 제목 '{title}' 이 칸({w:.2f} in)보다 김: 제목을 줄이거나 칸을 넓힘")
    t = text(s, x, y, w, 0.2, t_, size=8, bold=True, color=SLATE9, align=align)
    t.name = f"IXCAP_{kind}"
    _reg(s).setdefault("cap_keys", {})[t.shape_id] = key or (round(y, 1), x, y)
    return _obst(s, t)


def table_caption(s, x, y, w, title, unit=None, rows=None, key=None):
    """표 위 한 줄: 왼쪽 [표 #. 제목], 오른쪽 (단위: ...). y 는 이 줄의 위쪽"""
    uw = (_text_w(f"(단위: {unit})", 6.35) + 0.1) if unit else 0
    _caption(s, "표", title, x, y, w - uw, "l", key=key)
    if unit:
        u = text(s, x + w - uw, y, uw, 0.2, f"(단위: {unit})", size=6.35, color=SLATE9, align="r")
        u.name = "IXUNIT"
        _obst(s, u)
    elif rows and not any("단위" in str(h) for h in rows[0]) and sum(bool(_NUM_CELL.match(str(c).strip())) and not _YEAR.match(str(c).strip())
                                               for r in rows[1:] for c in r[1:]) >= 2:   # 첫 열 제외, 단위 열이 있으면 통과
        _warn(s, f"숫자 표 '{title}' 에 단위 없음: unit='건, %' (표 오른쪽 위에 (단위: ...))")


def figure_caption(s, title, x, y, w, key=None):
    """그림 아래 가운데 [그림 #. 제목]. y 는 이 줄의 위쪽"""
    return _caption(s, "그림", title, x, y, w, "c", key=key)


def table(s, x, y, col_w, rows, size=7.26, row_h=0.24, header_fill=TBL_HEAD, header_color=WHITE,
          highlight=(), align=None, cell_hl=(), row_hs=None):
    """rows[0] 머리행. highlight: 남색 행, cell_hl: 남색 칸, row_hs: 데이터 행마다 높이 (없으면 모두 row_h)"""
    nr, nc = len(rows), len(col_w)
    head_h = min(row_h, 0.27)
    hs = list(row_hs) if row_hs else [row_h] * (nr - 1)
    if len(hs) != nr - 1:
        raise ValueError(f"table row_hs {len(hs)}개: 데이터 행 {nr - 1}개와 같아야 함")
    gf = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(sum(col_w)), Inches(head_h + sum(hs)))
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
        tbl.rows[i].height = Inches(head_h if i == 0 else hs[i - 1])
        for j in range(nc):
            c = tbl.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.06)
            c.margin_top = c.margin_bottom = Inches(0.01)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            head = i == 0
            c.fill.solid()
            c.fill.fore_color.rgb = _rgb(header_fill if head else WHITE)
            hl = (i - 1) in highlight or (i - 1, j) in cell_hl
            col = header_color if head else (ACCENT if hl else BLACK)
            p = c.text_frame.paragraphs[0]
            a = (align[j] if align else ("c" if head else "l"))
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[a]
            r = p.add_run()
            r.text = str(rows[i][j])
            _font(r, size, bold=head or hl, color=col)
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
        box(s, x, y, w, 0.24, lines, fill=SLATE3, size=7.26 if not sub else 6.35, bold=True, inset=(0.02, 0, 0.02, 0))
    else:
        body = box(s, x, y + 0.27, w, h - 0.27, fill=WHITE, line=BLACK, lw=0.5, shadow=True)
        box(s, x, y, w, 0.24, title, fill=SLATE6, size=7.26, bold=True, color=WHITE, inset=(0.02, 0, 0.02, 0))
    return body


def server(s, x, y, w, h, t, fill=GRAY, size=8.16):
    """서버 상자 (회색)"""
    return _obst(s, box(s, x, y, w, h, t, fill=fill, size=size, bold=True, color=WHITE))


def db(s, x, y, w, h, t, size=7.26, line=None):
    """모듈, DB 상자 (남회색)"""
    return _obst(s, box(s, x, y, w, h, t, fill=SLATE6, line=line, lw=0.75, size=size, bold=True, color=WHITE))


def area(s, x, y, w, h, line=SLATE4, fill=SLATE0):
    """서버 아래 모듈 묶음 영역"""
    a = box(s, x, y, w, h, fill=fill, line=line, lw=0.25, geom="round", radius=0)
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
    """종류 범례. items=[(색, 이름)], kind="square" 는 색 네모(기능 묶음 색), kind="line" 에서 [(색, 이름, 점선 종류)]"""
    for i, it in enumerate(items):
        c, t, dash = (tuple(it) + (None,))[:3]
        yy = y + i * pitch
        if kind == "oval":
            box(s, x, yy + 0.02, 0.1, 0.1, fill=None, line=c, lw=1.25, geom="oval")
        elif kind == "square":   # 기능 묶음 색 범례: 색 네모 + 이름
            box(s, x, yy + 0.02, 0.12, 0.12, fill=c)
        elif dash:
            ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(yy + 0.078), Inches(0.16), 0)   # 높이 0 테두리 = 선
            ln.fill.background()
            _line(ln, c, 1.25, dash)
        else:
            box(s, x, yy + 0.065, 0.16, 0.025, fill=c)
        text(s, x + 0.2, yy, _text_w(t, 6.35) + 0.1, 0.2, t, size=6.35, color=SLATE9, align="l")
    w = 0.2 + max(_text_w(it[1], 6.35) + 0.1 for it in items)
    _reg(s)["obst"].append(((x, y, x + w, y + pitch * (len(items) - 1) + 0.2), 0))


def tag(s, x, y, t="EAC", fill=GREEN):
    """연결선 위 작은 초록 표시 (EAC 등)"""
    t_ = box(s, x, y, 0.23, 0.16, t, fill=fill, size=5.44, bold=True, color=WHITE, inset=(0, 0, 0, 0))
    _reg(s)["obst"].append((_bb(t_), -1))  # 선 위 표시라 선 지남 검사에서 뺌
    return t_


def step(s, x, y, n, label=None, color=ACCENT, label_w=None, label_color=BLACK, on_line=None, owner=0, cond=None, cw=0):
    """번호 원 (0.15 in) 과 오른쪽 설명. on_line: 'v' 세로선 위(글은 원 오른쪽), 'hb' 가로선 위에 원만 얹고 글은 원 위쪽에 붙임.
    선 위에서는 원에 흰 테두리 (글 바탕은 칠하지 않음: 색 바탕 위에서 흰 상자가 보임)"""
    if label_w is None:
        label_w = _text_w(label or "", 7.26) + 0.12
    c = num_badge(s, x, y, n, color, d=NUM_D_LINE, size=NUM_PT_LINE, line=WHITE if on_line else None)
    if on_line == "hb":
        if label:
            text(s, x, y - 0.17, label_w, 0.16, label, size=7.26, bold=True, color=label_color, anchor="b")
        bb = (x, y - 0.17, x + max(0.15, label_w), y + 0.15)
        if cond:
            text(s, x + 0.19, y + 0.17, cw, 0.15, cond, size=6.35, bold=True, color=RED, align="l")   # 이 함수의 label 은 글 인자라 label() 을 못 씀
    else:
        if label:
            text(s, x + 0.19, y - 0.03, label_w, 0.22, label, size=7.26, bold=True, color=label_color, anchor="m")
            _reg(s).setdefault("rows", []).append((x, x + 0.19 + label_w, y + 0.075))
        bb = (x, y - 0.03, x + 0.19 + (label_w if label else 0), y + 0.19)
    _reg(s)["obst"].append((bb, owner))   # owner: 번호가 얹힌 선의 시작 상자 (자기 선은 지나도 됨)
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


def _duo_keys():
    """두 색 구성도 아이콘 세트 이름 (build_icons.SET)"""
    from build_icons import SET
    return set(SET)


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
    _reg(s).setdefault("icon_keys", set()).add(key)
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


def connect(s, a, side_a, b, side_b, color=GRAY, width=1.5, head="end", dash=None, mid=0.5, text=None,
            bypass=None):
    """붙은 연결선. text 는 선 이름, bypass 는 바깥 우회 거리"""
    _before_draw(s)
    c = _connect(s, a, side_a, b, side_b, color, width, head, dash, mid, bypass=bypass)
    R = _reg(s)
    pts = (_bypass_pts(a, side_a, b, side_b, bypass) if bypass
           else _route(a, side_a, b, side_b, mid, check=False))
    R["segs"].extend(zip(pts, pts[1:]))
    R["links"].append((pts, a.shape_id, b.shape_id, f"{a.name} -> {b.name}"))
    if head != "end":   # 양방향, 머리 없음, 시작 머리는 방향을 단정하지 않고 양쪽 끝이 받고 내보내는 것으로 셈
        R.setdefault("bidir", set()).add((a.shape_id, b.shape_id))
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


BYPASS_TINY = 0.0014   # 바깥 우회선의 바탕 폭 (in). 꺾임 위치 = 폭 x adj


def _bypass_pts(a, sa, b, sb, off):
    """같은 쪽 변끼리 바깥으로 도는 선의 꺾임점 (right, bottom)"""
    (x1, y1), (x2, y2) = _pt(a, sa), _pt(b, sb)
    if sa != sb or sa not in ("right", "bottom"):
        raise ValueError(f"우회선은 right 와 right 또는 bottom 과 bottom 만: {sa}, {sb}")
    return ([(x1, y1), (x1 + off, y1), (x1 + off, y2), (x2, y2)] if sa == "right"
            else [(x1, y1), (x1, y1 + off), (x2, y1 + off), (x2, y2)])


def _connect(s, a, side_a, b, side_b, color=GRAY, width=1.5, head="end", dash=None, mid=0.5, htype="triangle",
             bypass=None):
    """a.side_a 에서 b.side_b 로 붙는 선. bypass: 바깥 우회 거리"""
    (x1, y1), (x2, y2) = _pt(a, side_a), _pt(b, side_b)
    ha, hb = side_a in ("left", "right"), side_b in ("left", "right")
    swap = False
    if ha != hb and not ha:  # 세로에서 가로로 가는 선은 역방향으로 그림
        a, side_a, b, side_b = b, side_b, a, side_a
        (x1, y1), (x2, y2) = (x2, y2), (x1, y1)
        ha, hb = True, False
        swap = True
    straight = abs(x1 - x2) < 0.005 or abs(y1 - y2) < 0.005
    if bypass:   # 같은 줄의 두 변: 폭 거의 0, 꺾임을 크게
        prst, rot = "bentConnector3", 0 if ha else 90
        mid = bypass / BYPASS_TINY
    elif straight and ha == hb:
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
    if bypass:
        dx, dy = (BYPASS_TINY, dy) if rot == 0 else (dx, BYPASS_TINY)
        x2, y2 = x1 + dx, y1 + dy
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
        etree.SubElement(av, qn("a:gd"), name="adj1", fmla=f"val {round(mid * 100000)}")
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
FIT_GAP = 0.18                                # 패널 칸 블록 사이 기본 틈
LIMIT_REF_W = sum(REPORT_WIDTHS) / 3 - 2 * PAD   # 글자 한도 기준 칸 폭 (3열 평균)
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
        raw = im.convert("RGB").resize((200, 200), _Im.NEAREST).tobytes()   # getdata 는 Pillow 14 에서 없어짐
    px = [raw[i:i + 3] for i in range(0, len(raw), 3)]
    hits = sum(1 for p in px if p in (bytes((255, 127, 14)), bytes((31, 119, 180))))
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
          vmax=None, labels=True, y_title=None):
    """편집 가능한 차트. kind: col, bar, stack, line
    - y_title 이나 number_format 의 단위 글자 필요, 계열은 4개까지"""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_MARKER_STYLE
    types = {"col": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
             "stack": XL_CHART_TYPE.COLUMN_STACKED, "line": XL_CHART_TYPE.LINE_MARKERS}
    if kind not in types:
        raise ValueError(f"chart kind '{kind}' 없음: {list(types)}")
    if not categories or not series:
        raise ValueError("chart: 항목이나 계열이 없음")
    if not y_title and not re.search(r'"[^"]+"|%|[$₩€¥]', number_format):
        _warn(s, "차트 단위가 없음: y_title='건수 (건)' 또는 number_format='0\"건\"'")
    if len(series) > 4:
        print(f"주의: 차트 계열 {len(series)}개 (4개 넘으면 색 구분이 어려움)")
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
    if y_title:
        _axis_title(va, y_title)
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
    _reg(s).setdefault("gantt_names", []).extend(t[0] for t in tasks if t[3] != "mile")   # d.deliverables 대조용, 마일스톤 제외
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
    bar_txt = {"done": SLATE9, "now": WHITE, "plan": WHITE}   # 연한 막대 위는 진한 글자
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
                fill=fills[kind], geom="round", radius=0.5, size=6.35, bold=True, color=bar_txt[kind],
                inset=(0.04, 0, 0.04, 0))
    if today is not None:
        tx = x + label_w + today * pw
        box(s, tx, y + hh, 0.001, body_h, fill=None, line=RED, lw=1, dash="dash")
        text(s, tx - 0.3, y + hh + body_h + 0.02, 0.6, 0.18, today_label, size=6.35, bold=True, color=RED, align="c")
    return y + hh + body_h


class _Cols(list):
    """report_panels 결과: 칸 목록 + gaps"""
    fit = False

    def shrink(self):
        """모든 칸 done() 뒤 패널을 내용 높이로 줄임"""
        end = max(c._end for c in self) + (REPORT_BOTTOM - BODY_BOTTOM)
        if end < self.bottom - 0.01:
            for b_ in self.bodies:
                b_.height = Inches(end - self.top - 0.29)
            for sh in self.mids:   # 패널 틈 화살표와 배지는 새 가운데로
                sh.top += Inches((end - self.bottom) / 2)
            for sh in self.foot:
                sh.top += Inches(end - self.bottom)
            self.bottom = end
        if self.bottom - self.top < 0.45 * (REPORT_BOTTOM - REPORT_TOP):
            _notice(self.s, f"패널 높이 {self.bottom - self.top:.2f} in (본문 높이의 45% 미만): 내용이 너무 적은 장")


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


def report_panels(s, titles, widths=None, arrows=False, hub_gap=None, steps=None, top=REPORT_TOP, footnote=None,
                  fit=True):
    """보고형 패널과 칸 (Col) 목록. fit 이면 패널이 내용 높이로 줄어듦"""
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
    cols, x = _Cols(), REPORT_LEFT
    cols.s, cols.fit, cols.top, cols.bottom = s, fit, top, bottom
    cols.bodies, cols.mids, cols.foot = [], [], []
    if footnote:  # 패널 아래 각주 한 줄
        cols.foot.append(text(s, REPORT_LEFT, bottom + 0.04, REPORT_RIGHT - REPORT_LEFT, 0.17, footnote, size=6.35,
                              color=SLATE9))
    mid = (top + 0.29 + bottom) / 2
    cols.gaps = []
    for i, (t, w) in enumerate(zip(titles, widths)):
        cols.bodies.append(panel(s, x, top, w, bottom - top, t))
        _reg(s).setdefault("panel_titles", set()).add(t)
        cols.append(Col(s, x + PAD, w - 2 * PAD, top=top + (BODY_TOP - REPORT_TOP), bottom=bottom - (REPORT_BOTTOM - BODY_BOTTOM)))
        cols[-1].grp = cols
        if i < n - 1:
            cols.gaps.append((x + w, x + w + gaps[i]))
            n0 = len(s.shapes)
            if steps and i < len(steps) and steps[i]:
                step_badge(s, x + w, x + w + gaps[i], mid, *steps[i])
            elif arrows and i != hub_gap:
                big_arrow(s, x + w + gaps[i] / 2 - 0.19, mid - 0.165)
            cols.mids += list(s.shapes)[n0:]
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
_ABBR_END = re.compile(r"(^|[\s\[(])([A-Za-z]{1,4}|\d+)\.$")   # 'Acq.', '[표 1.' 같은 약어와 번호 뒤 점


def ko_breaks(text, width, size):
    """문장 끝, 긴 문장은 쉼표 뒤에서 줄 바꿈"""
    import re
    out = []
    parts = re.split(r"(?<=[.!?])(?<!\.\.)(?<![.][.][.])\s+(?=\S)", text)
    sents = []
    for p_ in parts:   # 약어로 끝난 조각은 다음과 붙임
        if sents and (sents[-1].endswith(_NOT_END) or _ABBR_END.search(sents[-1])):
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


def _limit(s, key, texts, w=LIMIT_REF_W):
    """블록 글자 한도(LIMITS x 칸 폭 비례) 넘는 글은 경고. strict 저장이 멈춤"""
    lim = int(LIMITS[key] * w / LIMIT_REF_W + 1e-6)
    for t_ in texts:
        if len(t_) > lim:
            _warn(s, f"{key} {len(t_)}자 (한도 {lim}자, 칸 폭 {w:.2f} in): '{t_[:20]}'")


class Col:
    """패널 안 세로 칸, done() 에서 높이 분배"""
    MIN_GAP, MAX_GAP, MAX_ROW = 0.12, 0.50, 0.32

    def __init__(self, s, x, w, top=BODY_TOP, bottom=BODY_BOTTOM):
        self.s, self.x, self.w, self.top, self.bottom = s, x, w, top, bottom
        self.items = []  # (높이, 그리기, 행 수, 행 추가 한도)
        _reg(s).setdefault("cols", []).append(self)

    def _ck(self, y):
        """제목 번호 순서: 칸 묶음 위 → 칸 왼쪽 → 블록 위"""
        return (round(self.top, 1), round(self.x, 1), y)

    def _add(self, h, draw, rows=0, room=0.0):
        self.items.append((h, draw, rows, room))
        self.kinds = getattr(self, "kinds", []) + [sys._getframe(1).f_code.co_name]   # 같은 틀 되풀이 확인용
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

    def table(self, rows, ratio=None, row_h=0.26, title=None, unit=None, **kw):
        """표 블록. title: 왼쪽 위 [표 #. 제목], unit: 오른쪽 위 (단위: ...)"""
        if not rows or not rows[0]:
            raise ValueError("표 행이 없음")
        ratio = ratio or [1] * len(rows[0])
        cw = [self.w * k / sum(ratio) for k in ratio]

        def draw(y, r):
            table_caption(self.s, self.x, y, self.w, title, unit, rows, key=self._ck(y))
            return table(self.s, self.x, y + CAP_H, cw, rows, row_h=row_h + r, **kw)
        return self._add(CAP_H + min(row_h, 0.27) + row_h * (len(rows) - 1), draw,
                         rows=len(rows) - 1, room=max(self.MAX_ROW - row_h, 0))

    def _fixed_table(self, head, rows, ratio, row_h, **kw):
        """머리행 고정 표. 빈 칸이면 ValueError"""
        if not rows:
            raise ValueError(f"{head[0]} 표 행이 없음")
        for r in rows:
            if len(r) != len(head) or any(not str(c).strip() for c in r):
                raise ValueError(f"표 행은 {head} 칸을 모두 채움: {r}")
        return self.table([head, *[list(r) for r in rows]], ratio=ratio, row_h=row_h, **kw)

    def kpi_table(self, rows, row_h=0.3, last=KPI_HEAD[-1], title=None, unit=None):
        """KPI 표. rows=[(지표, 기존, 목표, 계산식, 근거)]. last: 마지막 열 이름"""
        _reg(self.s)["kpi_table"] = True
        return self._fixed_table([*KPI_HEAD[:-1], last], rows, [1.25, 1, 1.3, 2.25, 2.2], row_h, title=title,
                                 unit=unit)

    def compare_table(self, rows, left="현행", right="개선", highlight=(), row_h=0.3, title=None, unit=None):
        """현행 및 개선 한 표. highlight: 바뀐 행 (0부터)"""
        n = len(rows)
        if any(not isinstance(i, int) or not 0 <= i < n for i in highlight):
            raise ValueError(f"compare_table highlight {tuple(highlight)} 는 행 번호 0 ~ {n - 1}")
        return self._fixed_table(["항목", left, right], rows, [1, 1.6, 1.6], row_h, cell_hl=[(i, 2) for i in highlight],
                                 title=title, unit=unit)

    def collect_table(self, rows, row_h=0.3, first="공정", title=None, unit=None):
        """수집 항목 표. first: 첫 열 이름 (기본 '공정')"""
        for r in rows:
            if len(r) == 3 and r[2] not in COLLECT_METHODS:
                raise ValueError(f"집계 방법 '{r[2]}' 은 {COLLECT_METHODS} 중 하나")
        return self._fixed_table([first, *COLLECT_HEAD[1:]], rows, [1, 2.4, 1], row_h, title=title, unit=unit)

    def das_table(self, rows, title=None, row_h=0.28):
        """데이터 수집 명세서 (DAS) 표. rows=[(구분 X1 ~ X5 또는 Yn, Data, Data Type, Range, Acq. Time, Unit, Stage, Connected)]"""
        for r in rows:
            if len(r) and not DAS_KIND.match(str(r[0]).strip()):
                raise ValueError(f"DAS 구분 '{r[0]}' 은 X1 ~ X5 (작업, 환경, 셋팅, 실행1, 실행2) 또는 Yn (결과)")
        return self._fixed_table(DAS_HEAD, rows, [0.55, 1.55, 0.85, 0.9, 0.95, 0.6, 0.75, 1.2], row_h,
                                 title=title or "데이터 수집 명세서 (DAS)")

    def quality_table(self, values, title=None, row_h=0.28):
        """수집 데이터 품질 지수 표. values={완전성: %, 유일성: %, 유효성: %, 일관성: %, 정확성: %} (키는 앞 글자만 맞으면 됨)"""
        got = {}
        for k in QUALITY_W:
            hit = [v for kk, v in values.items() if k.startswith(str(kk).split()[0])]
            if len(hit) != 1:
                raise ValueError(f"quality_table: '{k.split()[0]}' 값이 없음 (완전성, 유일성, 유효성, 일관성, 정확성 다섯)")
            if not 0 <= float(hit[0]) <= 100:
                raise ValueError(f"quality_table: '{k.split()[0]}' {hit[0]} 은 0 ~ 100 (%)")
            got[k] = float(hit[0])
        idx = sum(v * QUALITY_W[k] / 100 for k, v in got.items())
        ok = "충족" if idx >= QUALITY_TARGET else "미달"
        rows = [["구분", "품질지수", "가중치", "가중치지수", "오류율"]]
        rows += [[k, f"{v:.2f}", f"{QUALITY_W[k]}", f"{v * QUALITY_W[k] / 100:.2f}", f"{100 - v:.2f}"]
                 for k, v in got.items()]
        rows.append([f"품질 지수 ({ok}, 기준 {QUALITY_TARGET:g} 이상)", "", "100", f"{idx:.2f}", f"{100 - idx:.2f}"])
        return self.table(rows, ratio=[2.2, 1, 1, 1, 1], row_h=row_h, title=title or "수집 데이터 품질 지수", unit="%",
                          highlight=(len(rows) - 2,))

    @staticmethod
    def _priority(rows, cut):
        out = []
        for r in rows:
            if len(r) != 7 or any(not isinstance(v, (int, float)) or not 1 <= v <= 5 for v in r[1:]):
                raise ValueError(f"우선순위 행은 (과제, 기여도, 효과, 시급성, 투자, 효과발생, 관리), 점수 1 ~ 5: {r}")
            a = sum(w * v for w, v in zip(PRIORITY_W[0], r[1:4]))
            b = sum(w * v for w, v in zip(PRIORITY_W[1], r[4:7]))
            q = ("QW" if b >= cut else "SI") if a >= cut else ("NH" if b >= cut else "LN")
            out.append((r[0], r[1:4], round(a, 1), r[4:7], round(b, 1), q))
        return out

    def priority_table(self, rows, title=None, cut=3.5, row_h=0.28):
        """과제 우선순위 평가표 (중요도 50%, 용이성 50%). cut: 4분면 기준 점수"""
        rs = [[n, *map(str, s1), f"{a:.1f}", *map(str, s2), f"{b:.1f}", f"{a + b:.1f}", q]
              for n, s1, a, s2, b, q in self._priority(rows, cut)]
        return self.table([PRIORITY_HEAD, *rs], ratio=[2.2, 0.8, 0.8, 0.7, 0.7, 0.8, 0.8, 0.8, 0.7, 0.7, 0.6],
                          row_h=row_h, title=title or "과제 우선순위 평가", unit="점")

    def priority_quad(self, rows, title=None, cut=3.5, h=2.4):
        """과제 우선순위 4분면 (세로 전략적 중요도, 가로 실행의 용이성). 칸마다 과제 이름"""
        ps = self._priority(rows, cut)
        names = {q: [n for n, *_, qq in ps if qq == q] for q in ("SI", "QW", "LN", "NH")}
        heads = {"SI": "Strategic Initiatives (중장기)", "QW": "Quick Win (단기 실행)", "LN": "Leave for Now (배제)",
                 "NH": "Nice to Have (선택)"}

        def draw(y, r):
            label(self.s, self.x, y, self.w, f"세로 전략적 중요도, 가로 실행의 용이성 (기준 {cut:g}점)", align="l")
            gy, gw, gh = y + 0.2, (self.w - 0.06) / 2, (h - 0.2 - 0.06) / 2
            for k, q in enumerate(("SI", "QW", "LN", "NH")):
                bx, by = self.x + (k % 2) * (gw + 0.06), gy + (k // 2) * (gh + 0.06)
                _obst(self.s, box(self.s, bx, by, gw, gh, fill=SLATE0, line=GRAY_L, lw=0.75))
                box(self.s, bx, by, gw, 0.24, heads[q], fill=ACCENT if q == "QW" else SLATE6, size=7.26, bold=True,
                    color=WHITE)
                text(self.s, bx + 0.1, by + 0.3, gw - 0.2, gh - 0.36, "\n".join(names[q]) or "없음", size=7.26,
                     color=SLATE9 if names[q] else GRAY, anchor="t")
            figure_caption(self.s, title or "과제 우선순위 4분면", self.x, y + h + 0.02, self.w, key=self._ck(y))
        return self._add(h + CAP_H, draw)

    def levels(self, current, h=0.32):
        """수준 단계 띠 (Level 1 ~ 2 모니터링, 3 제어, 4 최적화, 5 자율운영). current: 0 ~ 3"""
        if not 0 <= current < len(LEVELS):
            raise ValueError(f"levels current 는 0 ~ {len(LEVELS) - 1} ({', '.join(LEVELS)})")
        return self.chevrons(LEVELS, current=current, h=h)

    def chart(self, kind, categories, series, h=1.9, title=None, **kw):
        """차트 블록. title: 아래 가운데 [그림 #. 제목]"""
        def draw(y, r):
            g = chart(self.s, self.x, y, self.w, h + r, kind, categories, series, **kw)
            figure_caption(self.s, title, self.x, y + h + r + 0.02, self.w, key=self._ck(y))
            return g
        return self._add(h + CAP_H, draw, rows=1, room=1.2)

    def scatter(self, series, x_title, y_title, h=2.2, title=None, **kw):
        """산점도 블록. 축 이름에 단위 (예: '목표 두께 (µm)'), title: 그림 제목"""
        def draw(y, r):
            g = scatter_chart(self.s, self.x, y, self.w, h + r, series, x_title, y_title, **kw)
            figure_caption(self.s, title, self.x, y + h + r + 0.02, self.w, key=self._ck(y))
            return g
        return self._add(h + CAP_H, draw, rows=1, room=1.2)

    def hist(self, values, edges, name="건수", unit="", count_unit="건", h=1.9, **kw):
        """히스토그램 블록 (편집 가능한 세로 막대)"""
        cats, counts = hist_bins(values, edges, unit)
        return self.chart("col", cats, {name: counts}, h=h, number_format=f'#,##0"{count_unit}"', **kw)

    def icon_cards(self, items, card=0.46):
        """아이콘 카드 줄"""
        return self._add(card + 0.12, lambda y, r: icon_cards(self.s, self.x, y, self.w, items, card=card))

    def kpis(self, items, h=0.85, gap=0.15):
        """핵심 수치 한 줄. items=[(값, 단위, 이름)]"""
        _limit(self.s, "kpis 이름", [it[2] for it in items], self.w)
        R = _reg(self.s)
        R["kpi_rows"] = R.get("kpi_rows", 0) + 1
        n = len(items)
        cw = (self.w - gap * (n - 1)) / n

        def draw(y, r):
            return [kpi(self.s, self.x + i * (cw + gap), y, cw, h, *it) for i, it in enumerate(items)]
        return self._add(h, draw)

    def section(self, title, rows, row_h=0.3, **kw):
        """회색 머리 띠 목록"""
        _limit(self.s, "section 행", rows, self.w)
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

    def gantt(self, periods, tasks, row_h=0.3, title=None, **kw):
        """일정 블록. title: 아래 가운데 [그림 #. 제목]"""
        gh = 0.27 + row_h * len(tasks) + (0.22 if kw.get("today") is not None else 0)

        def draw(y, r):
            end = gantt(self.s, self.x, y, self.w, periods, tasks, row_h=row_h, **kw)
            figure_caption(self.s, title, self.x, y + gh + 0.02, self.w, key=self._ck(y))
            return end
        return self._add(gh + CAP_H, draw)

    def image(self, path, h, marks=(), title=None):
        """그림 블록. h 는 최대 높이, title: 아래 가운데 [그림 #. 제목]"""
        from PIL import Image as _Im
        with _Im.open(path) as im:
            iw, ih = im.size
        real = min(h, self.w * ih / iw)
        if _default_mpl(path):
            _warn(self.s, f"그림 {os.path.basename(path)} 이 matplotlib 기본 색 (주황, 파랑): 그리기 전 mpl_style()")
        pad = 0.12 if marks else 0.0  # 번호 원이 그림 위로 0.17 in 나옴
        def draw(y, r):
            pic = screenshot(self.s, path, self.x, y + pad, self.w, real, marks)
            figure_caption(self.s, title, self.x, y + pad + real + 0.02, self.w, key=self._ck(y))
            return pic
        return self._add(real + pad + CAP_H, draw)

    def rule_list(self, title, items, pitch=0.29):
        """items=[(아이콘, 글[, 설명])]"""
        _limit(self.s, "rule_list 글", [it[1] for it in items], self.w)
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
        _limit(self.s, "tab_box 줄", lines, self.w)
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

    def num_cards(self, items, color=ACCENT, numbered=True):
        """카드 목록. numbered=False 면 번호 원 없음"""
        _limit(self.s, "num_cards 제목", [ttl for ttl, _ in items], self.w)
        _limit(self.s, "num_cards 설명", [x_ for _, d in items for x_ in (d if isinstance(d, list) else [d or ""])], self.w)
        lead = 0.12 + (NUM_D + 0.1 if numbered else 0)
        w = self.w - lead - 0.1
        for ttl, d in items:
            _unit(self.s, "num_cards", _td_h(ttl, d, w) + 0.2)
        gap = 0.12

        def draw(y, r):
            U = _unit_of(self.s, "num_cards")
            yy = y
            for n, (ttl, d) in enumerate(items, 1):
                box(self.s, self.x, yy, self.w, U, fill=WHITE, line=GRAY_L, lw=0.75, geom="round", radius=0.08)
                top = title_desc(self.s, self.x + lead, yy, w, U, ttl, d)
                if numbered:
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
        self._end = self.top
        g = getattr(self, "grp", None)
        if not self.items:
            _warn(self.s, f"패널 칸이 비었음 (x {self.x:.2f}). 블록을 넣거나 패널을 뺌")
            self._shrink_if_last(g)
            return self.top
        self.items = [((h() if callable(h) else h), d_, n_, r_) for h, d_, n_, r_ in self.items]
        need = sum(i[0] for i in self.items)
        gaps = len(self.items) - 1
        bottom = self.bottom
        if g is not None and g.fit:   # 같은 줄에서 가장 긴 칸의 높이에 맞춤
            bottom = min(bottom, self.top + max(c._natural() for c in g if c.items))
        free = bottom - self.top - need - self.MIN_GAP * gaps
        if free < -1e-6:
            _fail(f"패널 칸(x {self.x:.2f})이 {-free:.2f} in 넘침: 블록을 줄이거나 표 행 높이를 낮춤")
            free = 0.0
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
        self._end = y
        if g is None or not g.fit:
            if self.bottom - y > 0.3:
                _notice(self.s, f"패널 칸 아래 {self.bottom - y + 0.2:.2f} in 빔 (x {self.x:.2f}). fit=True 로 패널을 줄임")
        else:
            self._shrink_if_last(g)
        return y

    def _natural(self):
        """블록 높이 + 고른 틈 (done 전후 같음)"""
        hs = [(h() if callable(h) else h) for h, *_ in self.items]
        return sum(hs) + FIT_GAP * max(len(hs) - 1, 0)

    @staticmethod
    def _shrink_if_last(g):
        if g is not None and g.fit and all(getattr(c, "_done", False) for c in g):
            g.shrink()


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


SOFT = os.environ.get("INTERX_REGRESS") == "1"   # regress.py 만 켬: 엄격 실패도 저장해 그림 비교


def _fail(msg):
    if not SOFT:
        raise ValueError(msg)
    print("엄격 실패: " + msg.replace(chr(10), " / "))


def _warn(s, msg):
    """경고를 모은다. strict 저장이면 멈춤"""
    _reg(s).setdefault("warn", []).append(msg)
    print("주의: " + msg)


def _notice(s, msg):
    """주의만 출력. 저장은 멈추지 않음"""
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
    if head != "end":
        R.setdefault("bidir", set()).add((ia, ib))
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


def _on_line_spots(pts, w, h):
    """번호 원을 선 위에 얹는 자리 (x, y, 방향, 차지하는 네모). 세로 구간은 'v' (원이 선 위, 글은 원 오른쪽), 가로 구간은 'hb' (원만 선 위,
    글은 원 위쪽). 긴 구간 가운데부터. 원 가운데가 선 위에 오도록 y 는 선보다 0.105 위. h 가 0.22 보다 크면 아래 조건 글 자리"""
    segs = sorted(zip(pts, pts[1:]), key=lambda pq: -(abs(pq[1][0] - pq[0][0]) + abs(pq[1][1] - pq[0][1])))
    full, short = [], []
    for p, q in segs:
        horiz = abs(p[1] - q[1]) < 1e-9
        ln = abs(q[0] - p[0]) if horiz else abs(q[1] - p[1])
        for t in (0.5, 0.35, 0.65, 0.22, 0.78):
            mx, my = p[0] + (q[0] - p[0]) * t, p[1] + (q[1] - p[1]) * t
            if not horiz and ln >= h + 0.1:
                full.append((mx - 0.075, my - 0.105, "v", (mx - 0.075, my - 0.105, mx - 0.075 + w, my - 0.105 + h)))
            if horiz and ln >= max(w - 0.19, 0.19) + 0.2:
                bx = mx - (w - 0.19) / 2
                short.append((bx, my - 0.075, "hb", (bx, my - 0.245, bx + max(0.15, w - 0.19), my + 0.075 + max(0, h - 0.22))))
    yield from full
    yield from short


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


def _bottom_of(s):
    """제목 및 요약 줄 아래 도형들의 (왼쪽, 오른쪽, 아래)"""
    E = 914400
    xs = [(sh.left / E, (sh.left + sh.width) / E, (sh.top + sh.height) / E) for sh in s.shapes
          if not sh.is_placeholder and sh.width is not None and sh.top is not None and sh.top / E > 1.6
          and sh.name not in ("도장",)]
    return (min(a for a, _, _ in xs), max(b for _, b, _ in xs), max(c for _, _, c in xs)) if xs else None


def _place_figure(s):
    """page(figure=) 장: 모든 도형 아래 가운데에 그림 제목"""
    R = _reg(s)
    if "figure" not in R or R.get("figure_done"):
        return []
    R["figure_done"] = True
    bb = _bottom_of(s)
    if bb is None:
        return ["figure= 를 줬지만 도식이 없음"]
    x0, x1, y = bb
    figure_caption(s, R["figure"], x0, y + 0.06, x1 - x0, key=(99, 0, 0))
    if y + 0.06 + 0.2 > REPORT_BOTTOM + 0.1:
        return [f"그림 제목 자리 없음 (도식 아래 {y:.2f} in): 도식 아래를 0.3 in 비움 (DocGrid(bottom=) 를 줄임)"]
    return []


def caption_problems(s):
    """표마다 [표 ...], 차트와 큰 그림과 구성도마다 [그림 ...]"""
    E = 914400
    caps = {k: sum(sh.name == f"IXCAP_{k}" for sh in s.shapes) for k in ("표", "그림")}
    tables = sum(sh.has_table for sh in s.shapes)
    figs = sum(sh.has_chart for sh in s.shapes) + sum(
        sh._element.tag.endswith("}pic") and sh.width / E >= 2.0 for sh in s.shapes if not sh.is_placeholder)
    if "doc_top" in _reg(s) and not figs:
        figs = 1   # 구성도 장 (d.page(doc=True))
    out = []
    if tables > caps["표"]:
        out.append(f"표 {tables}개에 표 제목 {caps['표']}개: 표 왼쪽 위 [표 번호. 제목] (Col.table(title=, unit=))")
    if figs > caps["그림"]:
        out.append(f"그림 {figs}개(차트, 화면, 구성도)에 그림 제목 {caps['그림']}개: 아래 가운데 [그림 번호. 제목] "
                   "(블록 title=, 구성도 장은 d.page(..., figure=))")
    return out


def _number_captions(slides):
    """덱 전체 연번: 장 순서, 장 안은 칸 묶음 위 → 칸 왼쪽 → 블록 위, 구성도 장 제목은 끝"""
    n = {"표": 0, "그림": 0}
    for sl in slides:
        keys = _reg(sl).get("cap_keys", {})
        caps = sorted((sh for sh in sl.shapes if sh.name in ("IXCAP_표", "IXCAP_그림")),
                      key=lambda sh: keys.get(sh.shape_id, (round(sh.top / 914400, 1), sh.left / 914400, 0)))
        for sh in caps:
            k = sh.name[6:]
            for r in (r for p in sh.text_frame.paragraphs for r in p.runs):
                if f"[{k} " in r.text:
                    n[k] += 1
                    r.text = re.sub(rf"^\[{k} (#|\d+)\. ", f"[{k} {n[k]}. ", r.text)


_VAGUE = re.compile(r"^(?:(?:수집|전송|처리|관리|분석|저장|보고|운영|제어|연계) )?(계획|정보|데이터|자료|값|내용|설정|기능|서비스|모듈|결과)"
                    r"(?: (?:제공|전달|처리|관리|연계|확인|조회|생성|보관|전송|수신|송신|저장))?$")   # 앞말이 행위이거나 없고 대상(무슨 계획인지)이 빠진 이름


def diagram_notes(s):
    """구성도 문법 주의: 선 색은 의미 셋까지 (데이터 흐름, 제어 흐름, 굵은 공통 버스, 빨강 끊김은 별도), 아이콘은 한 세트"""
    out = []
    cols = set(s._element.xpath(".//p:cxnSp/p:spPr/a:ln/a:solidFill/a:srgbClr/@val")) - {RED}   # 빨강은 끊김 표시
    if len(cols) > 3:
        out.append(f"연결선 색 {len(cols)}종 ({', '.join(sorted(cols))}): 의미 셋(데이터, 제어, 공통 버스)까지, 나머지 구분은 선 모양으로, 색마다 범례")
    keys = _reg(s).get("icon_keys", set())
    icons = {k for k in keys if any(os.path.splitext(f)[0] == k for f in _icon_files())}
    duo = _duo_keys()
    if icons & duo and icons - duo:
        out.append(f"두 색 구성도 아이콘과 옛 색 아이콘이 섞임 ({', '.join(sorted(icons - duo))}): 한 장은 한 세트로")
    ins, outs = {}, {}   # 상자마다 들어오는 선과 나가는 선. fan, bundle, 양방향 선은 양쪽 끝이 받고 내보내는 것으로 셈
    for _, a_, b_, nm in _reg(s).get("links", ()):
        outs[a_], ins[b_] = outs.get(a_, 0) + 1, ins.get(b_, 0) + 1
        if nm.startswith(("fan ", "bundle")):
            outs[b_], ins[a_] = outs.get(b_, 0) + 1, ins.get(a_, 0) + 1
    for a_, b_ in _reg(s).get("bidir", ()):
        outs[b_], ins[a_] = outs.get(b_, 0) + 1, ins.get(a_, 0) + 1
    terms = _reg(s).get("terms", set())   # 이름이 모호한 상자: 처음 보는 사람이 '무슨 계획?' 하고 묻게 되는 이름 (대상이 빠진 것)
    for name in dict.fromkeys(n_ for _, n_ in list(_reg(s).get("grid_boxes", ())) + list(_reg(s).get("box_names", ()))):
        if _VAGUE.match(name) and not any(k in name for k in terms):
            out.append(f"상자 이름 '{name}' 이 모호함: 무슨 {_VAGUE.match(name).group(1)} 인지 처음 보는 사람이 모름. '대상 + 동작' 으로 구체화 "
                       "(예: 수집 항목 및 주기 제공) 하거나 abbr_legend 로 용어를 풀이")
    for sid, (role, name) in _reg(s).get("flow_roles", {}).items():
        i_, o_ = ins.get(sid, 0), outs.get(sid, 0)
        if role == "process" and not (i_ and o_):
            out.append(f"처리 상자 '{name}' 에 {'들어오는' if not i_ else '나가는'} 선이 없음: 값이 어디서 들어와 어디로 가는지 (원천과 도착) 선으로 보여야 함")
        elif role == "source" and not o_:
            out.append(f"원천 상자 '{name}' 에 나가는 선이 없음")
        elif role == "sink" and not i_:
            out.append(f"도착 상자 '{name}' 에 들어오는 선이 없음")
    return out


def _icon_files():
    return os.listdir(ICONS)


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
        spot, on_line = None, None
        if kind == "step":   # 번호 원은 선 위에 얹어 어느 선의 번호인지 바로 보이게 (자기 선은 비켜 가지 않고 나머지만 피함)
            others = [sg for sg in R["segs"] if sg not in set(zip(where[0], where[0][1:]))]
            r2 = {**R, "segs": others}
            cands = list(_on_line_spots(where[0], w, h))
            for walls in (True, False, None):   # None: 묶음 테두리는 따지지 않음 (선 위이므로 어느 선의 번호인지가 먼저)
                hit = next((c for c in cands if _free({**r2, "walls": []} if walls is None else r2, c[3], bool(walls))), None)
                if hit:
                    spot, on_line = (hit[0], hit[1]), hit[2]
                    break
            if spot is None:
                _notice(s, f"번호 {n} '{t}' 를 선 위에 놓을 자리가 없어 선 옆에 둠 (선이 짧거나 주변이 막힘): 어느 선의 번호인지 흐려지므로 번호 대신 선 이름(text=)을 쓰거나 상자 사이를 넓힘")
        if spot is None and kind != "boxstep":
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
            step(s, x, y + (0 if on_line == "hb" else 0.03), n, t, color=color, label_w=lw, on_line=on_line,
                 owner=where[1].shape_id if on_line else 0, cond=cond if on_line == "hb" else None, cw=cw)

            if cond and on_line != "hb":
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

    def group(self, col, title, kind="platform", sub=None, rows=None, header=None, cols=1, area=True, badge=None, tone=None):
        """header: 묶음 안 제품이나 프로그램 이름 띠. area=True 면 그 아래 영역이 하위 묶음을 감싸 '이 안에 들어 있음'이 보임. tone: 기능 색(TONES)"""
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
        g_ = _Group(self, x, w, top, bottom, title, kind, sub, header, cols, use_area=area, tone=tone)
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
    def __init__(self, g, x, w, top, bottom, title, kind, sub, header, cols, use_area, tone=None):
        self.g, self.s = g, g.s
        if tone is not None and tone not in TONES:
            raise ValueError(f"tone '{tone}' 은 {', '.join(TONES)} 중 하나")
        if tone and not header and not use_area:
            _notice(self.s, f"묶음 '{title}' 의 tone '{tone}' 을 칠할 자리가 없음 (header 와 area 가 모두 없음)")
        if header and not use_area:
            _notice(self.s, f"묶음 제목 '{header}' 아래에 하위 묶음을 감싸는 영역이 없음 (area=False): 이 이름 안에 하위 묶음이 들어 있는지 보이지 않음. area=True 로")
        self.body = frame(self.s, x, top, w, bottom - top, title, kind=kind, sub=sub)
        ix, iw = x + g.PAD, w - 2 * g.PAD
        y = top + 0.54
        self.header = None
        if header:
            self.header = server(self.s, ix, y, iw, 0.27, header, fill=TONES[tone][0] if tone else GRAY)
            y += 0.45
        self.area = area(self.s, ix, y, iw, bottom - 0.25 - y, fill=TONES[tone][1] if tone else SLATE0) if use_area else None
        self.ix, self.iw = (ix + g.INNER, iw - 2 * g.INNER) if use_area else (ix, iw)
        self.cw = self.iw / cols
        self.ncols = cols

    def box(self, r, c, t, kind="db", span=1, w=None, icon_=None, flow=None):
        """줄 r, 칸 c 에 상자. kind: db, server. flow: source(값을 내보내기만), sink(받기만), process(받아서 내보냄)이면 선이 빠졌는지 확인"""
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
        _set_flow(self.s, b, t, flow)
        pic = None
        if icon_:
            ih = self.g.bh + 0.12
            pic = icon(self.s, icon_, cx + bw / 2 + 0.08, y - 0.06, h=ih)
        if not kw_w:
            _col_box(self.s, (id(self), c, span, bool(icon_)), b, self.cw * span - (0.5 if icon_ else 0.42), pic)
        return b

    def sub(self, c, r0, r1, header=None, span=1, caption=None, tone=None):
        """묶음 안 작은 묶음 (칸 c, 줄 r0 ~ r1). tone: 기능 색(TONES), 같은 기능은 같은 색"""
        if tone is not None and tone not in TONES:
            raise ValueError(f"tone '{tone}' 은 {', '.join(TONES)} 중 하나")
        if not (0 <= r0 <= r1 < self.g.n) or not (0 <= c and c + span <= self.ncols):
            raise ValueError(f"sub 자리 (칸 {c}, span {span}, 줄 {r0} ~ {r1}) 가 묶음을 벗어남")
        g = self.g
        x0, w = self.ix + self.cw * c + 0.06, self.cw * span - 0.12
        out = _Sub(self, x0, w)
        top_r = r0
        if header:
            out.header = server(self.s, x0, g.y(r0) - g.bh / 2 - 0.015, w, g.bh + 0.03, header, fill=TONES[tone][0] if tone else GRAY)
            top_r = r0 + 1
            if top_r > r1:
                return out
        ay0 = g.y(top_r) - g.bh / 2 - 0.12
        ay1 = g.y(r1) + g.bh / 2 + 0.12
        out.area = area(self.s, x0, ay0, w, ay1 - ay0, fill=TONES[tone][1] if tone else SLATE0)
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

    def box(self, r, t, kind="db", w=None, flow=None):
        g = self.grp.g
        bw = w or min(self.w - 0.24, max(1.0, _text_w(t, 7.26) * 1.1 + 0.2))
        fn = db if kind == "db" else server
        b = fn(self.s, self.x + (self.w - bw) / 2, g.y(r) - g.bh / 2, bw, g.bh, t)
        _reg(self.s).setdefault("grid_boxes", []).append((b.shape_id, t))
        _set_flow(self.s, b, t, flow)
        if not w:
            _col_box(self.s, (id(self),), b, self.w - 0.24)
        return b


def _set_flow(s, shape, t, flow):
    """상자의 흐름 역할을 적어 둠 (구성도 주의가 선이 빠졌는지 봄)"""
    if flow is None:
        return
    if flow not in ("source", "sink", "process"):
        raise ValueError(f"flow '{flow}' 은 source(내보내기만), sink(받기만), process(받아서 내보냄) 중 하나")
    _reg(s).setdefault("flow_roles", {})[shape.shape_id] = (flow, t)


def note_by(s, shape, head, lines=()):
    """상자 옆 빈 자리에 설명 (head, lines)"""
    _reg(s)["pending"].append(("note", shape, head, list(lines), SLATE9, None))


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


# --- Mermaid 를 편집 가능한 도형으로
_MM_NODE = re.compile(r"\s*([A-Za-z_]\w*)(?:\[([^\[\]]+)\])?\s*")
_MM_EDGE = re.compile(r"\s*(?:--\s+([^>]+?)\s+-->|-->)\s*")
_MM_PART = re.compile(r"^participant\s+(\w+)(?:\s+as\s+(.+))?$")
_MM_MSG = re.compile(r"^(\w+)\s*(-->>|->>)\s*(\w+)\s*:\s*(.+)$")
MM_NODE_H = 0.34    # 노드와 참가자 상자 높이 (한 줄)
MM_LINE_H = 0.16    # 줄이 하나 늘 때 더하는 높이
MM_BYPASS0, MM_BYPASS_STEP = 0.3, 0.22   # 체인 옆 우회선: 첫 거리, 겹치는 선마다 더하는 거리 (in)


def _mm_text(t, br="\n"):
    """노드 글: 따옴표 제거, <br> 은 줄 바꿈 (br)"""
    return re.sub(r"\s*<br\s*/?>\s*", br, t.replace('"', "")).strip()


def _mm_size(texts):
    """글 모음에 맞는 상자 (폭, 높이)"""
    ls = [ln for t in texts for ln in t.split("\n")]
    n = max(len(t.split("\n")) for t in texts)
    return max(0.9, max(_text_w(ln, 7.26) for ln in ls) + 0.3), max(MM_NODE_H, MM_LINE_H * n + 0.08)


def _mm_levels(spans):
    """겹치는 구간끼리 다른 번호 (우회선 거리 층)"""
    lv = []
    for i, (lo, hi) in enumerate(spans):
        used = {lv[j] for j, (l2, h2) in enumerate(spans[:i]) if lo <= h2 and l2 <= hi}
        lv.append(next(k for k in itertools.count() if k not in used))
    return lv


def _mm_parse_flow(lines):
    """flowchart 줄을 (가로 여부, 노드, 선) 으로"""
    head = lines[0].split()
    if len(head) != 2 or head[0] not in ("graph", "flowchart") or head[1] not in ("TD", "TB", "LR"):
        raise ValueError(f"mermaid: '{lines[0]}' 지원 밖 (graph TD, graph TB, graph LR, sequenceDiagram 만)")
    nodes, edges = {}, []
    for ln in lines[1:]:
        pos, prev, lab = 0, None, None
        while True:
            m = _MM_NODE.match(ln, pos)
            if not m:
                raise ValueError(f"mermaid: '{ln}' 지원 밖 문법 (A[이름], A --> B, A -- 글 --> B 만)")
            nid, t = m.group(1), m.group(2)
            if t:
                nodes[nid] = _mm_text(t) or nid
            else:
                nodes.setdefault(nid, nid)
            if prev is not None:
                if prev == nid:
                    raise ValueError(f"mermaid: '{ln}' 자기 자신으로 가는 선은 지원 밖")
                edges.append((prev, nid, lab))
            prev, pos = nid, m.end()
            if pos >= len(ln):
                break
            e = _MM_EDGE.match(ln, pos)
            if not e or e.end() >= len(ln):
                raise ValueError(f"mermaid: '{ln}' 지원 밖 문법 (A[이름], A --> B, A -- 글 --> B 만)")
            lab, pos = _mm_text(e.group(1) or "", " ") or None, e.end()
    return head[1] == "LR", nodes, edges


def _mm_layers(nodes, edges):
    """가장 긴 경로로 층 나눔 (되돌아가는 선 제외)"""
    out = {i: [] for i in nodes}
    for a, b, _ in edges:
        out[a].append(b)
    state, back = {}, set()

    def dfs(u):
        state[u] = 1
        for v in out[u]:
            if state.get(v) == 1:
                back.add((u, v))
            elif v not in state:
                dfs(v)
        state[u] = 2
    for i in nodes:
        if i not in state:
            dfs(i)
    fwd = [(a, b) for a, b, _ in edges if (a, b) not in back]
    layer = dict.fromkeys(nodes, 0)
    for _ in nodes:
        for a, b in fwd:
            layer[b] = max(layer[b], layer[a] + 1)
    rows = [[i for i in nodes if layer[i] == k] for k in range(max(layer.values()) + 1)]
    pred = {i: [a for a, b in fwd if b == i] for i in nodes}
    frac = {}
    for r in rows:
        r.sort(key=lambda i: sum(frac[a] for a in pred[i]) / len(pred[i]) if pred[i] else 0.5)
        frac.update({i: (k + 0.5) / len(r) for k, i in enumerate(r)})
    return rows


def _mm_flow(s, lines, x, y, w, h):
    """체인의 건너뛰기, 되돌아가기 선은 체인 옆 우회"""
    horiz, nodes, edges = _mm_parse_flow(lines)
    rows = _mm_layers(nodes, edges)
    L, M = len(rows), max(map(len, rows))
    lay = {i: k for k, r in enumerate(rows) for i in r}
    nw, nh = _mm_size(nodes.values())
    skips = [(a, b) for a, b, _ in edges if lay[b] - lay[a] != 1]
    levels = _mm_levels([(min(lay[a], lay[b]), max(lay[a], lay[b])) for a, b in skips]) if M == 1 else []
    far = MM_BYPASS0 + MM_BYPASS_STEP * max(levels) if levels else 0.0   # 가장 바깥 우회선 거리
    if horiz:
        cell = (h - nh) / (M - 1) if M > 1 else h   # 바깥 노드 가장자리가 영역 끝에 닿음
        step_ = (w - nw) / (L - 1) if L > 1 else 0
        if (M > 1 and cell < nh + 0.14) or (L > 1 and step_ - nw < 0.6) or nh + far > h:
            raise ValueError(f"mermaid: flowchart 가 {w:.1f} x {h:.1f} in 에 안 들어감 (층 {L}, 한 층 최대 {M}개)")
    else:
        cell = (w - nw) / (M - 1) if M > 1 else w
        step_ = (h - nh) / (L - 1) if L > 1 else 0
        if (M > 1 and cell < nw + 0.2) or (L > 1 and step_ - nh < 0.35) or nw + far > w:
            raise ValueError(f"mermaid: flowchart 가 {w:.1f} x {h:.1f} in 에 안 들어감 (층 {L}, 한 층 최대 {M}개)")
    cx0, cy0 = x + (w - far) / 2, y + (h - far) / 2   # 우회선 폭만큼 체인을 밀어 가운데에
    shapes = {}
    for k, r in enumerate(rows):
        for j, nid in enumerate(r):
            off = (j - (len(r) - 1) / 2) * cell
            if horiz:
                lx, ty = (x + k * step_ if L > 1 else x + (w - nw) / 2), cy0 + off - nh / 2
            else:
                lx, ty = cx0 + off - nw / 2, (y + k * step_ if L > 1 else y + (h - nh) / 2)
            shapes[nid] = _obst(s, box(s, lx, ty, nw, nh, nodes[nid], fill=WHITE, line=ACCENT, lw=1.0,
                                       name=f"노드 {nid}"))
    R, lv = _reg(s), iter(levels)
    for a, b, lab in edges:
        A, B = shapes[a], shapes[b]
        if lay[b] - lay[a] == 1:
            sa, sb = ("right", "left") if horiz else ("bottom", "top")
            connect(s, A, sa, B, sb)
            pts = _route(A, sa, B, sb, 0.5, check=False)
        elif M == 1:   # 건너뛰거나 되돌아가는 선: 체인 옆으로 돌아감
            sa = sb = "bottom" if horiz else "right"
            off = MM_BYPASS0 + MM_BYPASS_STEP * next(lv)
            connect(s, A, sa, B, sb, bypass=off)
            pts = _bypass_pts(A, sa, B, sb, off)
        else:
            n0 = len(R.get("warn", []))
            link(s, A, B, text=lab)
            if len(R.get("warn", [])) > n0:
                raise ValueError(f"mermaid: 선 {a} -> {b} 이 상자를 피해 그려지지 않음 (한 줄 체인이 아닌 그림의 "
                                 "건너뛰거나 되돌아가는 선). 중간 노드를 거치게 고치거나 DocGrid 로")
            continue
        if lab:   # 선 이름은 저장 직전에 빈 자리로
            R["pending"].append(("text", (pts, A, B), lab, None, SLATE9, None))
    return shapes


def _mm_seq(s, lines, x, y, w, h):
    parts, msgs = {}, []
    for ln in lines[1:]:
        pm, mm = _MM_PART.match(ln), _MM_MSG.match(ln)
        if pm:
            parts.setdefault(pm.group(1), _mm_text(pm.group(2) or pm.group(1)))
        elif mm:
            a, arrow, b, t = mm.groups()
            if a == b:
                raise ValueError(f"mermaid: '{ln}' 자기 자신에게 보내는 메시지는 지원 밖")
            parts.setdefault(a, a)
            parts.setdefault(b, b)
            msgs.append((a, b, _mm_text(t, " "), arrow == "-->>"))
        else:
            raise ValueError(f"mermaid: '{ln}' 지원 밖 문법 (participant, A->>B: 글, A-->>B: 글 만)")
    if not msgs:
        raise ValueError("mermaid: sequenceDiagram 에 메시지가 없음")
    ids = list(parts)
    cell = w / len(ids)
    pw, nh = _mm_size(parts.values())
    pw = max(1.0, pw)
    pitch = (h - nh - 0.1) / len(msgs)
    if pw > cell - 0.2 or pitch < 0.42:
        raise ValueError(f"mermaid: sequenceDiagram 이 {w:.1f} x {h:.1f} in 에 안 들어감 "
                         f"(참가자 {len(ids)}, 메시지 {len(msgs)})")
    cx = {p: x + cell * (i + 0.5) for i, p in enumerate(ids)}
    for p in ids:   # 세로 점선은 연결선이 아닌 선 도형
        ln_ = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx[p]), Inches(y + nh), 0, Inches(h - nh))
        ln_._element.spPr.find(qn("a:prstGeom")).set("prst", "line")
        _no_style_effects(ln_)
        _fill(ln_, None)
        _line(ln_, GRAY, 0.75, "dash")
        ln_.name = f"생명선 {p}"
    shapes = {p: _obst(s, box(s, cx[p] - pw / 2, y, pw, nh, parts[p], fill=ACCENT, size=7.26, bold=True,
                              color=WHITE, name=f"참가자 {p}")) for p in ids}
    for k, (a, b, t, reply) in enumerate(msgs):
        cy = y + nh + 0.1 + pitch * (k + 0.5)
        if _text_w(t, 6.35) + 0.08 > abs(cx[a] - cx[b]) - 0.2:
            raise ValueError(f"mermaid: 메시지 '{t}' 가 길어 참가자 간격 {abs(cx[a] - cx[b]):.2f} in 에 안 들어감")
        bar = {p: _obst(s, box(s, cx[p] - 0.05, cy - 0.11, 0.1, 0.22, fill=ACCENT, name=f"활성 막대 {k + 1} {p}"))
               for p in (a, b)}
        right = cx[b] > cx[a]
        connect(s, bar[a], "right" if right else "left", bar[b], "left" if right else "right",
                color=GRAY if reply else ACCENT, dash="dash" if reply else None, text=t)
    return shapes


def mermaid(s, src, x, y, w, h, title=None):
    """Mermaid 를 도형과 붙은 선으로 (flowchart, sequenceDiagram). title: 아래 가운데 [그림 #. 제목]
    - 따옴표는 빼고 <br> 은 줄 바꿈. 지원 밖은 ValueError"""
    lines = [ln.strip() for ln in src.strip().splitlines() if ln.strip() and not ln.strip().startswith("%%")]
    if not lines:
        raise ValueError("mermaid: 글이 비어 있음")
    out = (_mm_seq if lines[0] == "sequenceDiagram" else _mm_flow)(s, lines, x, y, w, h - CAP_H)
    figure_caption(s, title, x, y + h - CAP_H + 0.02, w)
    return out


# --- 덱
def _apply_ko_breaks(sl):
    """글 조각 하나인 문단에 ko_breaks 적용"""
    def frames():
        for sh in sl.shapes:
            if sh.is_placeholder or sh.name.startswith(("IXCAP_", "IXUNIT")):   # 표 및 그림 제목은 한 줄
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


from ixpaths import PROFILE_DIR  # noqa: E402  개인 설정 (플러그인 밖)


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

    def __init__(self, author, template=TEMPLATE, parts_only=False, forbid=(), allow=(), purpose=None, minutes=None, notes_required=False):
        """parts_only: 본문만. forbid: 덱에 나가면 안 되는 정규식
        - allow: 검사에서 뺄 정규식 (화면 문구 그대로 옮긴 글 등)
        - purpose: 쓰임새 (PURPOSES 중 하나). minutes: 발표 분 (노트 글자 수로 예상 시간을 잼)
        - notes_required: True 면 부록 아닌 본문 장마다 발표자 노트(30자 이상)가 있어야 저장 검사를 통과 (발표하는 덱)"""
        if purpose is not None and purpose not in PURPOSES:
            raise ValueError(f"purpose '{purpose}' 는 {PURPOSES} 중 하나")
        if minutes is not None and not (isinstance(minutes, (int, float)) and minutes > 0):
            raise ValueError(f"minutes '{minutes}' 는 0 보다 큰 수")
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
        if purpose:   # 검사 도구가 읽음
            self.prs.core_properties.category = "ix-purpose:" + purpose
        flags = ([f"ix-minutes:{minutes:g}"] if minutes else []) + (["ix-notes:required"] if notes_required else [])
        if flags:   # 검사 도구가 읽음
            self.prs.core_properties.content_status = " ".join(flags)
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

    @staticmethod
    def _repeat_notes(body):
        """두 장 연속 같은 패널 폭과 같은 블록 순서면 주의 (출력만)"""
        def sig(sl):
            cs = _reg(sl).get("cols", [])
            return tuple((round(c.w, 1), tuple(getattr(c, "kinds", ()))) for c in cs) if cs else None
        sigs = [sig(sl) for sl in body]
        for i in range(1, len(sigs)):
            if sigs[i] and sigs[i] == sigs[i - 1]:
                print(f"주의 본문 {i}, {i + 1}장: 패널 폭과 블록 순서가 같음 (틀로 찍은 느낌). 한 장은 블록 조합이나 패널 수를 바꿈")

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
        """목차: 절 번호, 절 이름, 장 주제. 한 열(절 5개 이하)은 절마다 자기 줄 수만큼, 두 열은 모든 절을 같은 높이의 격자에 놓아
        열이 달라도 번호, 이름, 구분선이 같은 높이에 섬. 한 절은 최대 4줄(넘으면 '외 N개'), 칸이 모자라면 줄 간격과 줄 수를 줄임"""
        items = getattr(self, "_toc_items", [])
        if not items:
            raise ValueError("목차에 넣을 절이 없음: d.section() 과 d.page() 로 장을 만든다")
        s = self._toc
        top, bottom, gap0 = 1.45, 7.1, 0.14
        avail = bottom - top

        def shown(cap):
            def cut(pg):
                if len(pg) <= cap:
                    return pg
                return [f"{len(pg)}개 장"] if cap == 1 else pg[:cap - 1] + [f"외 {len(pg) - (cap - 1)}개"]
            return [(n, cut(pg)) for n, pg in items]

        def heights(sh, per, uniform):   # 줄 간격을 줄여 칸에 맞춤. (절별 높이 목록, 줄 간격, 필요 높이)
            maxp, pitch = max(len(pg) for _, pg in sh), 0.21
            while True:
                hs = [0.46 + pitch * (maxp if uniform else len(pg)) for _, pg in sh]
                need = (per * hs[0] if uniform else sum(hs)) + (per - 1) * gap0
                if need <= avail + 1e-6 or pitch <= 0.17:
                    return hs, pitch, need
                pitch -= 0.01

        pick = None
        for cap in (4, 3, 2, 1):   # 줄 수를 줄이기 전에 칸 수를 늘림
            sh = shown(cap)
            for cols in ((1, 2) if len(items) <= 5 else (2,)):
                per = math.ceil(len(sh) / cols)
                r = heights(sh, per, uniform=cols == 2)
                if r[2] <= avail + 1e-6:
                    pick = (sh, cols, per, r)
                    break
            if pick:
                break
        if not pick:
            sh = shown(1)
            per = math.ceil(len(sh) / 2)
            pick = (sh, 2, per, heights(sh, per, True))
        sh, cols, per, (hs, pitch, need) = pick
        if need > avail + 1e-6:
            _warn(s, f"목차가 칸을 넘침 ({need:.2f} in, 칸 {avail:.2f} in): 절을 합치거나 절 수를 줄임")
        x0s, w = ((0.75,), 6.6) if cols == 1 else ((0.75, 4.55), 3.6)
        gap = max(gap0, min(0.3, (avail - need + (per - 1) * gap0) / max(per - 1, 1) if per > 1 else gap0))
        for c in range(cols):
            chunk = list(enumerate(sh))[c * per:(c + 1) * per]
            y = top
            for r, (i, (name, pages)) in enumerate(chunk):
                if _text_w(name, 14) > w - 0.7 or any(_text_w(pg, 9) > w - 0.7 for pg in pages):
                    _warn(s, f"목차 '{name}' 의 글이 칸보다 김: 절 이름과 장 주제를 짧게")
                text(s, x0s[c], y, 0.6, 0.36, f"{i + 1:02d}", size=20, bold=True, color=ORANGE)
                text(s, x0s[c] + 0.65, y, w - 0.65, 0.36, name, size=14, bold=True, color=BLACK)
                for k, pg in enumerate(pages):
                    text(s, x0s[c] + 0.67, y + 0.44 + pitch * k, w - 0.67, 0.2, pg, size=9, color=SLATE9)
                y += hs[i] + gap
                if r < len(chunk) - 1:
                    box(s, x0s[c], y - gap / 2, w, 0.01, fill=GRAY_L)   # 구분선: 두 열에서는 같은 높이

    def _end(self):
        return self._add("앤드페이지_다크네이비")

    def report_slide(self, title, subtitle, lead=None, notes=None):
        """보고형 머리 띠와 요약 한 줄. notes: 발표자 노트"""
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
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        return s

    def section(self, name):
        """절 시작. 이후 page() 제목에 번호"""
        self._sec = getattr(self, "_sec", 0) + 1
        self._sec_name = name
        self._page = 0
        getattr(self, "_toc_items", []).append((name, []))

    def criteria(self, items):
        """평가항목 목록. 저장 때 대응표 장이 Thank you 앞에 붙음"""
        items = [str(i).strip() for i in items]
        if not items or "" in items or len(set(items)) != len(items):
            raise ValueError("criteria: 항목은 하나 이상, 빈 글과 중복 없이")
        if len(items) > 14:
            raise ValueError("criteria: 항목은 14개까지 (대응표가 한 장)")
        self._criteria, self._covers = items, {}

    def deliverables(self, stages):
        """{단계: [산출물]}. 저장 때 gantt 작업 이름과 대조. 표 행 반환"""
        if not stages or any(not v for v in stages.values()):
            raise ValueError("deliverables: 단계마다 산출물이 하나 이상")
        self._deliv = dict(stages)
        return [["단계", "산출물"]] + [[k, ", ".join(v)] for k, v in stages.items()]

    def _crit_slide(self):
        """Thank you 앞 평가항목 대응표 장. 답 없는 항목은 저장 실패"""
        if not getattr(self, "_criteria", None):
            return
        cv = self._covers
        self._crit_problems = [f"평가항목 '{c}' 에 답한 장이 없음 (d.page(covers=[...]))" for c in self._criteria
                               if c not in cv]
        self.section("평가항목 대응")
        s = self.page("평가항목 대응표", lead=f"평가항목 {len(self._criteria)}개와 답한 쪽 번호")
        w, n = REPORT_RIGHT - REPORT_LEFT, len(self._criteria)
        panel(s, REPORT_LEFT, REPORT_TOP, w, REPORT_BOTTOM - REPORT_TOP, "평가항목별 답한 장")
        rh = min(0.55, max(0.24, (BODY_BOTTOM - BODY_TOP - 0.27 - CAP_H - 0.3) / n))   # 0.3: 출처 메모
        rows = [["평가항목", "답한 장"]] + [[c, ", ".join(cv.get(c, ["없음"]))] for c in self._criteria]
        table_caption(s, REPORT_LEFT + PAD, BODY_TOP, w - 2 * PAD, "평가항목별 답한 장")
        table(s, REPORT_LEFT + PAD, BODY_TOP + CAP_H, [(w - 2 * PAD) * 0.65, (w - 2 * PAD) * 0.35], rows, size=10.5,
              row_h=rh)
        notes(s, REPORT_LEFT + PAD, BODY_TOP + CAP_H + 0.27 + rh * n + 0.1, w - 2 * PAD, ["출처: 본문 장 꼬리말 쪽 번호"])

    def page(self, subtitle, lead=None, doc=False, notes=None, covers=None, figure=None):
        """다음 장. doc: 구성도 장, notes: 노트, covers: 평가항목, figure: 구성도 장 그림 제목 (도식 아래 가운데)"""
        if not getattr(self, "_sec", 0):
            raise ValueError("page: 먼저 d.section('절 이름')")
        if getattr(self, "_closed", False):
            raise ValueError("저장한 뒤 장을 더하지 않는다 (목차와 Thank you 가 이미 들어감): build.py 를 처음부터 다시 돌린다")
        for c in covers or ():
            if c not in getattr(self, "_criteria", ()):
                raise ValueError(f"covers '{c}' 이 d.criteria([...]) 에 없음")
        self._page += 1
        if getattr(self, "_toc_items", None):
            self._toc_items[-1][1].append(subtitle.split(" (")[0])
        title = f"{self._sec}. {self._sec_name}"
        sub = f"{self._sec}-{self._page}. {subtitle}"
        s = (self.doc_slide if doc else self.report_slide)(title, sub, lead=lead, notes=notes)
        if figure is not None:
            _reg(s)["figure"] = figure
        for c in covers or ():   # 꼬리말 쪽 번호
            self._covers.setdefault(c, []).append(f"{len(self.prs.slides) - len(self._lib)}쪽")
        return s

    def doc_slide(self, title, subtitle, lead=None, notes=None):
        """설계문서형 장 (넓은 도식 자리)"""
        s = self.report_slide(title, subtitle, lead=lead, notes=notes)
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
            self._crit_slide()
            self._fill_toc()
            self._end()
            self._closed = True
        lib_n = 0 if getattr(self, "_saved", False) else len(self._lib)
        problems = []
        body = [sl for sl in list(self.prs.slides)[lib_n:] if sl.slide_layout.name == "본문_소제목 O"]
        kpi_slides = [i for i, sl in enumerate(body, 1) if _reg(sl).get("kpi_items", 0)]
        problems.extend(getattr(self, "_crit_problems", []))
        names = {n for sl in body for n in _reg(sl).get("gantt_names", [])}
        if getattr(self, "_deliv", None):   # 일정 작업과 산출물 단계 이름이 서로 있어야 함
            problems += [f"산출물 단계 '{k}' 이 일정(gantt) 작업 이름에 없음" for k in self._deliv if k not in names]
            problems += [f"일정 작업 '{k}' 의 산출물이 없음 (d.deliverables)" for k in sorted(names - set(self._deliv))]
        for i, sl in enumerate(body, 1):
            if _reg(sl).get("kpi_table") and not any(SRC_LINE.search(ln.strip()) for t_ in slide_texts(sl)
                                                    for ln in t_.split(chr(10))):
                problems.append(f"본문 {i}장: KPI 표 아래 '※ 출처: 자료 이름 (날짜)' 메모가 없음")
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
            if i > lib_n and sl.slide_layout.name == "본문_소제목 O":
                problems += [f"{i - lib_n}장: {m}" for m in _place_figure(sl)]
                problems += [f"{i - lib_n}장: {m}" for m in caption_problems(sl)]
            if i > lib_n:
                sub = re.sub(r"^\d+-\d+\.\s*|\s*\(.*\)$", "", next(
                    (q.text_frame.text for q in sl.placeholders if q.placeholder_format.idx == 11), ""))
                for m in diagram_notes(sl):   # 출력만 (저장은 멈추지 않음)
                    print(f"주의 {i - lib_n}장: {m}")
                for t in sorted(_reg(sl).get("panel_titles", ())):   # 출력만 (저장은 멈추지 않음)
                    if sub and re.sub(r"\s*\(.*\)$", "", t) == sub:
                        print(f"주의 {i - lib_n}장: 패널 머리 '{t}' 가 소제목과 같은 말: 패널 머리를 내용 이름으로")
                for m in dict.fromkeys(layout_problems(sl)):
                    print(f"주의 {i - lib_n}장: {m}")
                    problems.append(f"{i - lib_n}장: {m}")
                alt = sl._element.xpath(".//p:cNvPr/@descr")   # 그림 대체 텍스트 (파일 이름)
                for t in list(slide_texts(sl)) + alt:
                    for p in self.forbid:
                        if p.search(t):
                            print(f"주의 {i - lib_n}장: 금지 문구 '{p.search(t).group()}'")
                            problems.append(f"{i - lib_n}장: 금지 문구 '{p.search(t).group()}' (forbid)")
        _number_captions(list(self.prs.slides)[lib_n:])
        self._repeat_notes(body)
        if strict and problems:
            _fail("덱 문제 " + str(len(problems)) + "건, 저장하지 않음:\n" + "\n".join(problems))
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


import ixdoc  # noqa: E402,F401  설계 문서 부품을 Col 과 Deck 에 붙임 (ixdoc.py)
from ixdiag import *  # noqa: E402,F401,F403  구성도 역할 부품 (ixdiag.py)
