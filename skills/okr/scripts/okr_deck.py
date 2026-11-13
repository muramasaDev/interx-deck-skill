#!/usr/bin/env python3
"""월간 OKR 덱. 사용: okr_deck.py build|check"""
import argparse
import copy
import json
import math
import os
import re
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu, Inches

sys.stdout.reconfigure(errors="replace")  # cp949 콘솔에서 멈추지 않게

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
MARK = "【확인 필요】"
PLACEHOLDERS = ("【입력】", "(입력)", MARK)  # 자리표시자. 남으면 check 실패
TICKET = re.compile(r"\b[A-Z][A-Z0-9]+-\d+\b|KR\d{7}|\(-\d+\)|\bP\d-\d+")   # 티켓 번호
UNDONE = re.compile(r"미실시|미착수|미진입|미완|보류|연기|이월")            # 실적 칸에 쓰지 않는 말
ACH_MAX = 16                                                              # 달성현황 글자 한계
MAX_ROWS = 5  # 넘으면 아래 Plan, Review 와 겹침
MERGE_COLS = (0, 1, 9)  # 그룹 단위로 세로 병합되는 열
FONTS = ('<a:latin typeface="Noto Sans KR"/><a:ea typeface="Noto Sans KR"/>'
         '<a:cs typeface="Noto Sans KR"/><a:sym typeface="Noto Sans KR"/>')


def A(tag):
    return "{%s}%s" % (NS_A, tag)


def month_of(period):
    m = re.search(r"[/_]\s*(\d+월)", period or "")
    if not m:
        sys.exit(f"period 형식 오류: '{period}' (예: 2026년 Q4 / 11월)")
    return m.group(1)


# --- 슬라이드 분류

def board_shape(slide):
    for sh in slide.shapes:
        if sh.has_table:
            head = " ".join(sh.table.cell(0, c).text for c in range(min(3, len(sh.table.columns))))
            if "Company" in head and "Key Result" in head:
                return sh
    return None


def board_table(slide):
    sh = board_shape(slide)
    return sh.table if sh is not None else None


def _own_table(slide, prefix):
    for sh in slide.shapes:
        if sh.has_table and sh.table.cell(0, 0).text.strip().startswith(prefix):
            return sh.table
    return None


def continuation_of(slide):
    """Plan/Review 만 옮겨 둔 다음 장"""
    try:
        slides = list(slide.part.package.presentation_part.presentation.slides)
        i = [s.slide_id for s in slides].index(slide.slide_id)
    except Exception:
        return None
    if i + 1 < len(slides):
        n = slides[i + 1]
        if board_shape(n) is None and _own_table(n, "Monthly Review") is not None:
            return n
    return None


def table_by_head(slide, prefix):
    t = _own_table(slide, prefix)
    if t is None and prefix.startswith("Monthly") and board_shape(slide) is not None:
        c = continuation_of(slide)
        if c is not None:
            t = _own_table(c, prefix)
    return t


def kind_of(slide):
    for sh in slide.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text
            if "일반 달성 목표" in t:
                return "general"
            if "도전 달성 목표" in t:
                return "challenge"
    return None


def period_of(slide):
    t = table_by_head(slide, "Individual OKR Period")
    m = re.search(r"(\d{4})년\s*Q\d\s*[/_]\s*(\d+)월", t.cell(0, 1).text) if t is not None else None
    return (int(m.group(1)), int(m.group(2))) if m else None


def is_appendix(slide):
    return any(sh.has_text_frame and sh.text_frame.text.strip() == "Appendix" for sh in slide.shapes)


def classify(prs):
    """종류별 (오래된 보드, 새 보드). 다르면 멈춤"""
    boards = {"general": [], "challenge": []}
    for s in prs.slides:
        if board_table(s) is not None and kind_of(s) and period_of(s):
            boards[kind_of(s)].append(s)
    for k, lst in boards.items():
        if len(lst) != 2:
            sys.exit(f"템플릿 오류: {k} 보드가 {len(lst)}장입니다 (2장이어야 함). 번들 템플릿이나 직전 완성본을 쓰세요.")
        lst.sort(key=period_of)
    return boards


# --- 셀, 문단 쓰기

def _first_run(p):
    return p.find(A("r"))


def _ensure_run(p):
    """첫 런, 빈 문단이면 endParaRPr 로 생성"""
    r = _first_run(p)
    if r is None:
        r = etree.Element(A("r"))
        end = p.find(A("endParaRPr"))
        if end is not None:
            rpr = copy.deepcopy(end)
            rpr.tag = A("rPr")
            r.append(rpr)
            p.insert(list(p).index(end), r)
        else:
            p.append(r)
        etree.SubElement(r, A("t"))
    return r


def _set_para_text(p, text):
    r = _ensure_run(p)
    for other in p.findall(A("r"))[1:]:
        p.remove(other)
    for br in p.findall(A("br")):
        p.remove(br)
    r.find(A("t")).text = text


def set_cell_text(cell, text):
    """첫 런 서식 유지, 한 줄 셀 수정"""
    tb = cell._tc.txBody
    paras = tb.findall(A("p"))
    for extra in paras[1:]:
        tb.remove(extra)
    _set_para_text(paras[0], text)


def set_cell_lines(cell, lines):
    """첫 문단 복제, 여러 줄 셀 작성"""
    tb = cell._tc.txBody
    paras = tb.findall(A("p"))
    src = copy.deepcopy(paras[0])
    for p in paras:
        tb.remove(p)
    for line in lines:
        p = copy.deepcopy(src)
        _set_para_text(p, line)
        tb.append(p)


def _exemplar(cell):
    paras = [p._p for p in cell.text_frame.paragraphs if p.text.strip()]
    hdr = [p for p in paras if p.find(".//" + A("buNone")) is not None]
    bul = [p for p in paras if p.find(".//" + A("buChar")) is not None]
    labelled = bool(paras) and all(len(p.findall(A("r"))) == 2 and p.findall(A("r"))[0].find(A("rPr")).get("b") == "1" for p in paras)
    return {"paras": paras, "hdr": hdr, "bul": bul, "labelled": labelled}


def _para_from(src, texts):
    p = copy.deepcopy(src)
    runs = p.findall(A("r"))
    for t, r in zip(texts, runs):
        r.find(A("t")).text = t
    for r in runs[len(texts):]:
        p.remove(r)
    return p


def write_lines(cell, lines, ex):
    """Plan, Review 칸에 줄 목록 작성"""
    if not ex["paras"]:
        sys.exit("템플릿에 서식을 본뜰 문단이 없습니다 (Plan, Review 칸이 비어 있음). 번들 템플릿이나 직전 완성본을 쓰세요.")
    new = []
    if ex["labelled"]:
        for i, line in enumerate(lines):
            src = ex["paras"][0] if i == 0 else ex["paras"][min(1, len(ex["paras"]) - 1)]
            m = re.match(r"^([^:]+:)(.*)$", line)
            new.append(_para_from(src, [m.group(1), m.group(2)] if m else [line]))
    else:
        seen_hdr = 0
        prev_hdr = False
        for line in lines:
            if line.startswith("[") and ex["hdr"]:
                src = ex["hdr"][0] if seen_hdr == 0 or len(ex["hdr"]) < 2 else ex["hdr"][1]
                seen_hdr += 1
                prev_hdr = True
            else:
                pool = ex["bul"] or ex["paras"]
                src = pool[0] if prev_hdr or len(pool) < 2 else pool[1]
                prev_hdr = False
            new.append(_para_from(src, [line]))
    tb = cell._tc.txBody
    for p in tb.findall(A("p")):
        tb.remove(p)
    for p in new:
        tb.append(p)


def blank_cell(cell):
    tb = cell._tc.txBody
    paras = tb.findall(A("p"))
    for extra in paras[1:]:
        tb.remove(extra)
    p = paras[0]
    for r in p.findall(A("r"))[1:]:
        p.remove(r)
    r = _first_run(p)
    if r is not None:
        r.find(A("t")).text = ""


# --- Review 가리개

def find_cover(slide):
    for sh in slide.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and not sh.text_frame.text.strip() and sh.left > Emu(5000000):
            return sh
    return None


def add_cover(slide, el):
    new = copy.deepcopy(el)
    ids = [int(x) for x in slide.shapes._spTree.xpath("//p:cNvPr/@id")]
    new.xpath(".//p:cNvPr")[0].set("id", str(max(ids) + 1))
    slide.shapes._spTree.append(new)


# --- 보드 표: 행 수 조정, 채우기

def _tcs(tr):
    return tr.findall(A("tc"))


def _origin_of(trs, row_idx, col):
    """세로 병합 셀 그룹의 시작 셀"""
    for j in range(row_idx - 1, -1, -1):
        tc = _tcs(trs[j])[col]
        if tc.get("vMerge") != "1":
            return tc
    return None


def _drop_last_row(tbl):
    trs = tbl.findall(A("tr"))
    tr = trs[-1]
    for c in MERGE_COLS:
        tc = _tcs(tr)[c]
        if tc.get("vMerge") == "1":
            o = _origin_of(trs, len(trs) - 1, c)
            if o is not None and o.get("rowSpan"):
                o.set("rowSpan", str(max(1, int(o.get("rowSpan")) - 1)))
    tbl.remove(tr)


def _add_row(tbl):
    trs = tbl.findall(A("tr"))
    last = trs[-1]
    new = copy.deepcopy(last)
    for c, tc in enumerate(_tcs(new)):
        body = tc.find(A("txBody"))
        if body is not None:
            paras = body.findall(A("p"))
            for extra in paras[1:]:
                body.remove(extra)
            for r in paras[0].findall(A("r"))[1:]:
                paras[0].remove(r)
            r = _first_run(paras[0])
            if r is not None:
                r.find(A("t")).text = ""
        if c in MERGE_COLS:
            lt = _tcs(last)[c]
            if lt.get("vMerge") == "1":
                o = _origin_of(trs, len(trs) - 1, c)
                if o is not None and o.get("rowSpan"):
                    o.set("rowSpan", str(int(o.get("rowSpan")) + 1))
            elif lt.get("rowSpan"):
                lt.set("rowSpan", str(int(lt.get("rowSpan")) + 1))
            else:
                continue
            tc.set("vMerge", "1")
            if "rowSpan" in tc.attrib:
                del tc.attrib["rowSpan"]
    last.addnext(new)


def ensure_rows(table, n):
    if n < 1 or n > MAX_ROWS:
        sys.exit(f"보드 행은 1~{MAX_ROWS}개여야 합니다 (spec {n}개)")
    tbl = table._tbl
    while len(tbl.findall(A("tr"))) - 1 > n:
        _drop_last_row(tbl)
    while len(tbl.findall(A("tr"))) - 1 < n:
        _add_row(tbl)


def regroup_items(table, rows):
    """item 이 있는 행마다 세부항목 그룹 재분할"""
    if not any("item" in r for r in rows):
        return
    trs = table._tbl.findall(A("tr"))[1:]
    starts = [i for i in range(len(rows)) if i == 0 or "item" in rows[i]]
    first_body = copy.deepcopy(_tcs(trs[0])[1].find(A("txBody")))
    for k, s in enumerate(starts):
        end = starts[k + 1] if k + 1 < len(starts) else len(rows)
        for j in range(s, end):
            tc = _tcs(trs[j])[1]
            if j == s:
                if tc.get("vMerge"):
                    del tc.attrib["vMerge"]
                    old = tc.find(A("txBody"))
                    if old is not None:
                        tc.replace(old, copy.deepcopy(first_body))
                    else:  # 가려진 칸엔 txBody 가 없음. tcPr 앞에 넣음
                        tc.insert(0, copy.deepcopy(first_body))
                tc.set("rowSpan", str(end - s))
            else:
                tc.set("vMerge", "1")
                if "rowSpan" in tc.attrib:
                    del tc.attrib["rowSpan"]


def center_cell(cell):
    """가운데 정렬로 통일"""
    for p in cell._tc.txBody.findall(A("p")):
        pPr = p.find(A("pPr"))
        if pPr is None:
            pPr = etree.Element(A("pPr"))
            p.insert(0, pPr)
        pPr.set("algn", "ctr")


def fill_board_rows(slide, rows, clear_achievement):
    """행 키: company, item, kr, w, deliverable, ach"""
    sh = board_shape(slide)
    t = sh.table
    ensure_rows(t, len(rows))
    regroup_items(t, rows)
    sh.height = sum(r.height for r in t.rows)
    for r in range(1, len(rows) + 1):
        for c in range(3, 9):
            center_cell(t.cell(r, c))
    for i, row in enumerate(rows, start=1):
        if "company" in row and not t.cell(i, 0).is_spanned:
            set_cell_text(t.cell(i, 0), row["company"])
        if "item" in row and not t.cell(i, 1).is_spanned:
            set_cell_lines(t.cell(i, 1), row["item"].split("\n"))
        if "kr" in row:
            set_cell_text(t.cell(i, 2), row["kr"])
        for j, w in enumerate(row.get("w", [])):
            set_cell_text(t.cell(i, 3 + j), w)
        if "deliverable" in row:
            set_cell_text(t.cell(i, 7), row["deliverable"])
        if clear_achievement:
            set_cell_text(t.cell(i, 8), "")
        elif "ach" in row:
            set_cell_text(t.cell(i, 8), row["ach"])
            # 빈 칸은 서식이 커서 4W 칸 서식을 복제
            src = _first_run(t.cell(i, 6)._tc.txBody.find(A("p")))
            dst = _first_run(t.cell(i, 8)._tc.txBody.find(A("p")))
            if src is not None and src.find(A("rPr")) is not None:
                old = dst.find(A("rPr"))
                if old is not None:
                    dst.remove(old)
                dst.insert(0, copy.deepcopy(src.find(A("rPr"))))
        if "success" in row and not t.cell(i, 9).is_spanned:
            set_cell_lines(t.cell(i, 9), row["success"].split("\n"))


def set_period_and_headers(slide, period, month):
    set_cell_text(table_by_head(slide, "Individual OKR Period").cell(0, 1), period)
    c = continuation_of(slide)
    if c is not None and _own_table(c, "Individual OKR Period") is not None:
        set_cell_text(_own_table(c, "Individual OKR Period").cell(0, 1), period)
    set_cell_text(table_by_head(slide, "Monthly Plan").cell(0, 0), f"Monthly Plan ({month} 계획)")
    set_cell_text(table_by_head(slide, "Monthly Review").cell(0, 0), f"Monthly Review ({month} 실적)")


# --- 표지, Vision, Company OKR

def apply_author(prs, author, cover_date):
    """표지, Vision 의 작성자 및 날짜 작성"""
    for sh in prs.slides[0].shapes:
        if sh.has_text_frame and len(sh.text_frame.paragraphs) == 2 and re.fullmatch(r"\d{4}\.\d{2}", sh.text_frame.paragraphs[1].text.strip()):
            p0, p1 = sh.text_frame.paragraphs
            p0.runs[0].text = f"{author['team']} / {author['name']} {author['title']}"
            for r in p0.runs[1:]:
                r._r.getparent().remove(r._r)
            p1.runs[0].text = cover_date
            break
    else:
        sys.exit("표지에서 '소속 이름 / 날짜' 글상자를 찾지 못했습니다")
    for sh in prs.slides[1].shapes:
        if sh.has_table and sh.table.cell(0, 0).text.strip() == "작성자":
            set_cell_text(sh.table.cell(0, 1), author["name"])
            set_cell_text(sh.table.cell(0, 3), author["team"])
            return
    sys.exit("Vision 슬라이드에서 작성자 표를 찾지 못했습니다")


def apply_vision(prs, vision, mission):
    """vision: 문자열, mission: [{title, details}]"""
    for sh in prs.slides[1].shapes:
        if not sh.has_text_frame:
            continue
        txt = sh.text_frame.text
        tb = sh._element.txBody
        ps = tb.findall(A("p"))
        if txt.startswith("My Vision") and vision and len(ps) >= 2:
            _set_para_text(ps[1], vision)
        elif txt.startswith("My Mission") and mission and len(ps) >= 3:
            head, title_ex, detail_ex = ps[0], ps[1], ps[2]
            for p in ps[1:]:
                tb.remove(p)
            for m in mission:
                p = copy.deepcopy(title_ex)
                _set_para_text(p, m["title"])
                tb.append(p)
                for d in m.get("details", []):
                    p = copy.deepcopy(detail_ex)
                    _set_para_text(p, d)
                    tb.append(p)


def apply_company_okr(prs, c):
    """c: {title, items, success}"""
    for sh in prs.slides[2].shapes:
        if sh.has_table:
            t = sh.table
            cell = t.cell(1, 5)
            write_lines(cell, [c["title"]] + c.get("items", []), _exemplar(cell))
            if c.get("success"):
                set_cell_lines(t.cell(1, 6), c["success"])
            return
    sys.exit("Company OKR 슬라이드에서 표를 찾지 못했습니다")


# --- 장 순서, 삭제

def drop_slide(prs, slide):
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        if sid.id == slide.slide_id:
            prs.part.drop_rel(sid.rId)
            lst.remove(sid)
            return


def reorder(prs, ordered):
    lst = prs.slides._sldIdLst
    want = [s.slide_id for s in ordered]  # sldIdLst 를 비우기 전에 조회해야 함
    by_id = {sid.id: sid for sid in lst}
    for sid in list(lst):
        lst.remove(sid)
    for i in want:
        lst.append(by_id[i])


# --- Appendix

def _rpr(sz, bold, color):
    return (f'<a:rPr lang="ko-KR" sz="{sz}" b="{1 if bold else 0}" u="none" cap="none" strike="noStrike">'
            f'<a:solidFill><a:schemeClr val="{color}"/></a:solidFill>{FONTS}</a:rPr>')


def _esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _textbox(shape_id, name, x, y, w, h, text, sz, bold, color, algn, anchor, fill=None):
    return etree.fromstring(
        f'<p:sp xmlns:a="{NS_A}" xmlns:p="{NS_P}"><p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr anchor="{anchor}" lIns="91425" tIns="45700" rIns="91425" bIns="45700" wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
        f'<a:p><a:pPr marL="0" indent="0" algn="{algn}"><a:buNone/></a:pPr><a:r>{_rpr(sz, bold, color)}<a:t>{_esc(text)}</a:t></a:r></a:p></p:txBody></p:sp>')


def panel_xml(shape_id, x, y, w, h, bullets, heading="핵심 개요"):
    head = (f'<a:p><a:pPr marL="0" indent="0" algn="l"><a:lnSpc><a:spcPct val="115000"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:buNone/></a:pPr><a:r>{_rpr(1100, True, "dk2")}'
            f'<a:t>{_esc(heading)}</a:t></a:r></a:p>')
    body = "".join(
        '<a:p><a:pPr marL="171450" indent="-171450" algn="l"><a:lnSpc><a:spcPct val="130000"/></a:lnSpc>'
        '<a:spcBef><a:spcPts val="900"/></a:spcBef><a:buClr><a:schemeClr val="dk1"/></a:buClr>'
        '<a:buSzPts val="1000"/><a:buFont typeface="Noto Sans KR"/><a:buChar char="-"/></a:pPr>'
        f'<a:r>{_rpr(1000, False, "dk1")}<a:t>{_esc(b)}</a:t></a:r></a:p>' for b in bullets)
    return etree.fromstring(
        f'<p:sp xmlns:a="{NS_A}" xmlns:p="{NS_P}"><p:nvSpPr><p:cNvPr id="{shape_id}" name="핵심 개요 {shape_id}"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F7F9FC"/></a:solidFill>'
        '<a:ln w="9525"><a:solidFill><a:srgbClr val="BFC7D5"/></a:solidFill></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr anchorCtr="0" anchor="t" lIns="137160" tIns="109728" rIns="137160" bIns="109728" wrap="square">'
        f'<a:noAutofit/></a:bodyPr><a:lstStyle/>{head}{body}</p:txBody></p:sp>')


def build_appendix(prs, layout, entries, workdir):
    """이미지, [참고N] 캡션, 핵심 개요 장"""
    box_l, box_t, box_w, box_h = Inches(0.35), Inches(1.0), Inches(8.8), Inches(5.5)
    panel_l, panel_w = Inches(9.35), Inches(3.65)
    area_h = Inches(5.9)  # 블록이 낮으면 세로 가운데
    out = []
    for n, e in enumerate(entries, start=1):
        s = prs.slides.add_slide(layout)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        tree = s.shapes._spTree
        tree.append(_textbox(2, "Appendix 제목", 0, 203200, prs.slide_width, 461700, "Appendix", 2400, True, "dk2", "ctr", "t"))
        img = e["image"]
        if e.get("crop"):
            img_cropped = os.path.join(workdir, f"crop_{n}.png")
            Image.open(img).crop(tuple(e["crop"])).save(img_cropped)
            img = img_cropped
        iw, ih = Image.open(img).size
        scale = min(box_w / iw, box_h / ih)
        w, h = int(iw * scale), int(ih * scale)
        top = box_t + max(0, (area_h - (h + Inches(0.35))) // 2)
        s.shapes.add_picture(img, box_l + (box_w - w) // 2, top, w, h)
        ids = [int(x) for x in tree.xpath("//p:cNvPr/@id")]
        tree.append(_textbox(max(ids) + 1, "참고 캡션", box_l, top + h + Inches(0.05), box_w, 231300, f"[참고{n}] {e['title']}", 900, False, "dk1", "ctr", "ctr"))
        tree.append(panel_xml(max(ids) + 2, panel_l, top, panel_w, h + Inches(0.35), e["bullets"]))
        out.append(s)
    return out


# --- 빌드

def build(template, spec_path, out_path):
    spec = json.load(open(spec_path, encoding="utf-8-sig"))
    base = os.path.dirname(os.path.abspath(spec_path))
    prs = Presentation(template)
    boards = classify(prs)
    workdir = tempfile.mkdtemp(prefix="okr_deck_")
    covers = {k: find_cover(v[1]) for k, v in boards.items()}
    layout = boards["general"][1].slide_layout

    cur_spec, next_spec = spec["current"], spec["next"]
    for kind, (old, cur) in boards.items():
        cs = cur_spec[kind]
        if cur_spec.get("period"):
            set_period_and_headers(cur, cur_spec["period"], cur_spec.get("month") or month_of(cur_spec["period"]))
        # plan 이 없으면 전월의 익월 계획을 이월
        if "plan" in cs:
            plan_cell = table_by_head(cur, "Monthly Plan").cell(1, 0)
            write_lines(plan_cell, cs["plan"], _exemplar(plan_cell))
        review_ex = _exemplar(table_by_head(old, "Monthly Review").cell(1, 0))
        write_lines(table_by_head(cur, "Monthly Review").cell(1, 0), cs["review"], review_ex)
        fill_board_rows(cur, cs["rows"], clear_achievement=False)
        if covers[kind] is not None:
            covers[kind]._element.getparent().remove(covers[kind]._element)
        # 익월: 전월 장 재활용, Review 는 비움
        ns = next_spec[kind]
        plan_cell = table_by_head(old, "Monthly Plan").cell(1, 0)
        write_lines(plan_cell, ns["plan"], _exemplar(plan_cell))
        blank_cell(table_by_head(old, "Monthly Review").cell(1, 0))
        set_period_and_headers(old, next_spec["period"], next_spec.get("month") or month_of(next_spec["period"]))
        fill_board_rows(old, ns["rows"], clear_achievement=True)
        if covers[kind] is not None:
            add_cover(old, covers[kind]._element)

    if "vision" in spec or "mission" in spec:
        apply_vision(prs, spec.get("vision"), spec.get("mission"))
    if "company_okr" in spec:
        apply_company_okr(prs, spec["company_okr"])
    if "author" in spec:
        cur = period_of(boards["general"][1])  # 당월 (년, 월)
        apply_author(prs, spec["author"], f"{cur[0]}.{cur[1]:02d}")

    entries = copy.deepcopy(spec.get("appendix", []))
    for e in entries:
        if not os.path.isabs(e["image"]):
            e["image"] = os.path.join(base, e["image"])
    new_apps = build_appendix(prs, layout, entries, workdir)
    for s in [s for s in prs.slides if is_appendix(s) and s.slide_id not in {x.slide_id for x in new_apps}]:
        drop_slide(prs, s)

    new_ids = {s.slide_id for s in new_apps}
    front = [s for s in prs.slides if board_table(s) is None and s.slide_id not in new_ids]
    ordered = front + [boards["general"][1], boards["challenge"][1], boards["general"][0], boards["challenge"][0]] + new_apps
    reorder(prs, ordered)
    global SNAP
    SNAP = snapshot_boxes(prs)
    for line in fit_deck(prs):          # 1차: 추정으로 맞춤
        print("맞춤:", line)
    log, left = fit_measured(prs)       # 2차: 실제 배치로 재확인
    if BOX_OVER:                        # 3차: Plan/Review 를 통째로 다음 장으로
        for s, over in list(BOX_OVER):
            move_boxes_to_next(prs, s, SNAP)
            log.append(f"'{kind_of(s)} {period_of(s)[1]}월' Plan/Review를 통째로 다음 장으로 옮김")
        log2, left = fit_measured(prs)
        log += log2
    for line in settle_deck(prs):       # 4차: 테두리, 행 높이, 틀 크기 확정
        log.append(line)
    for i, s in enumerate(prs.slides, 1):
        sh = board_shape(s)
        bx = _boxes(s)
        if sh is None or not bx:
            continue
        below = [x for x in s.shapes if x is not sh and x.top > sh.top + 10 * E_PT
                 and x.left < sh.left + sh.width and x.left + x.width > sh.left]
        lack = (sh.top + sh.height + 8 * E_PT - min(x.top for x in below)) / E_PT
        if lack > 0:
            _shift_down(below, bx, lack)
            for x in bx:
                settle_table(x)
            log.append(f"슬라이드 {i}: 보드와 아래 칸 간격 확보 +{lack:.0f}pt")
    log += fit_pages_estimate(prs)      # 5차: 보드 밖 장도 꼬리말 위로
    plog, pleft = fit_pages_measured(prs)
    log += plog
    left += pleft
    for line in log:
        print("맞춤(실측):", line)
    for line in left:
        print("남은 넘침:", line)
    prs.save(out_path)
    print(f"saved {out_path} ({len(ordered)} slides, appendix {len(new_apps)})")


# --- 검사

def em(text):
    """글자 폭 추정 (전각 1.0)"""
    w = 0.0
    for ch in text:
        if ch == " ":
            w += 0.28
        elif ch.isascii():
            w += 0.55 if ch.isalnum() else 0.35
        else:
            w += 1.0
    return w


def _lines_of(shape):
    if shape.has_text_frame:
        yield from shape.text_frame.text.splitlines()
    elif shape.has_table:
        for row in shape.table.rows:
            for c in row.cells:
                yield from c.text.splitlines()


def _size_pt(cell_or_tf, default=9.0):
    body = getattr(cell_or_tf, "txBody", None)
    if body is None:
        body = getattr(cell_or_tf, "_txBody", None)
    for p in body.findall(A("p")) if body is not None else []:
        for r in p.findall(A("r")):
            rpr = r.find(A("rPr"))
            if rpr is not None and rpr.get("sz"):
                return int(rpr.get("sz")) / 100
    return default


def _wrap_lines(text, per_line):
    """단어 단위 줄 바꿈 추정"""
    if not text.strip():
        return 1
    lines, cur = 1, 0.0
    for w in text.split(" "):
        ww = em(w)
        if cur and cur + 0.28 + ww > per_line:
            lines += 1
            cur = 0.0
        while ww > per_line:
            lines += 1
            ww -= per_line
        cur += (0.28 if cur else 0) + ww
    return lines


def _text_height_pt(paras, width_pt, size_pt, spacing=1.25):
    """문단 목록 높이(pt) 추정"""
    per_line = max(1.0, width_pt / size_pt)
    return sum(_wrap_lines(t, per_line) for t in paras) * size_pt * spacing


def layout_problems(prs):
    """렌더 없이 잡는 넘침 목록"""
    out = []
    E = 12700.0  # EMU per pt
    for i, s in enumerate(prs.slides, 1):
        sh = board_shape(s)
        if sh is not None:
            total = _board_height(sh.table)
            below = [x.top for x in s.shapes if x is not sh and x.top > sh.top + 10 * E
                     and x.left < sh.left + sh.width and x.left + x.width > sh.left]
            limit = (min(below) - sh.top) / E if below else None
            if limit and total > limit + 2:
                out.append(f"슬라이드 {i}: 보드 표 예상 높이 {total:.0f}pt > 아래 Plan 표까지 {limit:.0f}pt — 1W~4W·주요결과물 문구를 짧게 (칸이 좁아 긴 문장은 세로로 늘어남)")
            for x in s.shapes:
                if x.has_table and x.table.cell(0, 0).text.strip().startswith(("Monthly Plan", "Monthly Review")):
                    avail = x.height / E - x.table.rows[0].height / E
                    need = _box_need(x)
                    if need > avail + 2:
                        name = x.table.cell(0, 0).text.strip().split("(")[0].strip()
                        out.append(f"슬라이드 {i}: {name} 칸 넘침 (예상 {need:.0f}pt > {avail:.0f}pt) — 줄 수를 줄이거나 문장을 짧게")
        for x in s.shapes:
            if x.has_text_frame and not x.has_table and x.text_frame.text.strip() and x.height:
                sz = _size_pt(x.text_frame)
                paras = [p.text for p in x.text_frame.paragraphs]
                need = _text_height_pt(paras, x.width / E - 14, sz, 1.3)
                bottom = (x.top + need * E) / E
                if bottom > prs.slide_height / E - 18 and need > x.height / E + 4:
                    out.append(f"슬라이드 {i}: '{x.text_frame.text.strip()[:16]}' 글상자가 슬라이드 아래로 넘침 — 항목 수를 줄이기")
    return out


# --- 화면에 맞추기
E_PT = 12700.0
MIN_BOARD_PT, MIN_BOX_PT = 7.5, 8.5   # 이보다 작으면 다음 장으로 넘김
BOX_OVER = []


def _set_size(body, sz_pt):
    """txBody 글자 크기를 sz_pt 로"""
    if body is None:
        return
    v = str(int(round(sz_pt * 100)))
    for tag in ("rPr", "endParaRPr"):
        for el in body.iter(A(tag)):
            el.set("sz", v)
    for r in body.iter(A("r")):
        if r.find(A("rPr")) is None:
            rpr = r.makeelement(A("rPr"), {"lang": "ko-KR", "sz": v})
            r.insert(0, rpr)


def _board_height(t, body_sz=None):
    """보드 표 예상 높이(pt)"""
    widths = [c.width / E_PT for c in t.columns]
    total = 0.0
    for r in range(len(t.rows)):
        h_need = t.rows[r].height / E_PT
        for c in range(len(t.columns)):
            cell = t.cell(r, c)
            if cell.is_spanned:
                continue
            span_w = sum(widths[c:c + (cell.span_width if cell.is_merge_origin else 1)])
            tcPr = cell._tc.find(A("tcPr"))
            g = (lambda k, dv: int(tcPr.get(k, dv)) / E_PT if tcPr is not None else dv / E_PT)
            need = _body_height(cell._tc.txBody, span_w - g("marL", 91440) - g("marR", 91440),
                                body_sz if (body_sz and r > 0) else None) * 1.03 + g("marT", 45720) + g("marB", 45720)
            if cell.is_merge_origin and cell.span_height > 1:
                need /= cell.span_height
            h_need = max(h_need, need)
        total += h_need
    return total


def _tw(x):
    """표 실제 폭(pt) = 열 너비 합"""
    return sum(c.width for c in x.table.columns) / E_PT


def _para_metrics(p):
    """줄 간격 배수, 앞뒤 간격, 왼쪽 여백"""
    ln, bef, aft, marl = 1.2, 0.0, 0.0, 0.0
    pPr = p.find(A("pPr"))
    if pPr is not None:
        marl = int(pPr.get("marL", "0")) / E_PT
        x = pPr.find(A("lnSpc"))
        if x is not None and x.find(A("spcPct")) is not None:
            ln = int(x.find(A("spcPct")).get("val")) / 100000 * 1.2
        for tag, attr in (("spcBef", "bef"), ("spcAft", "aft")):
            x = pPr.find(A(tag))
            if x is not None and x.find(A("spcPts")) is not None:
                v = int(x.find(A("spcPts")).get("val")) / 100
                if attr == "bef":
                    bef = v
                else:
                    aft = v
    return ln, bef, aft, marl


def _body_height(body, width_pt, sz=None):
    """txBody 높이(pt) 추정"""
    total = 0.0
    for p in body.findall(A("p")):
        text = "".join(t.text or "" for t in p.iter(A("t")))
        ln, bef, aft, marl = _para_metrics(p)
        s = sz or next((int(r.get("sz")) / 100 for r in p.iter(A("rPr")) if r.get("sz")), 9.0)
        per_line = max(1.0, (width_pt - marl) / s)
        total += bef + aft + _wrap_lines(text, per_line) * s * ln
    return total


def _box_need(x, sz=None):
    return _body_height(x.table.cell(1, 0)._tc.txBody, _tw(x) - 14.4, sz) * 1.03 + 7.2


def tighten_spacing(body, level):
    """문단 간격 줄이기 (level 1~3)"""
    cap_sp = {1: 600, 2: 300, 3: 0}[level]
    cap_ln = {1: 130000, 2: 115000, 3: 100000}[level]
    changed = False
    for pPr in body.iter(A("pPr")):
        for tag in ("spcBef", "spcAft"):
            x = pPr.find(A(tag))
            if x is not None and x.find(A("spcPts")) is not None and int(x.find(A("spcPts")).get("val")) > cap_sp:
                x.find(A("spcPts")).set("val", str(cap_sp))
                changed = True
        x = pPr.find(A("lnSpc"))
        if x is not None and x.find(A("spcPct")) is not None and int(x.find(A("spcPct")).get("val")) > cap_ln:
            x.find(A("spcPct")).set("val", str(cap_ln))
            changed = True
    return changed


def _shift_down(below, boxes, dpt):
    """보드 아래 도형을 내리고 바닥은 유지"""
    d = int(dpt * E_PT)
    bottom = max(x.top + x.height for x in boxes)
    for x in below:
        same_bottom = abs((x.top + x.height) - bottom) < E_PT * 6
        x.top = x.top + d
        if x in boxes or same_bottom:
            x.height = max(int(E_PT * 30), x.height - d)
            if x.has_table and len(x.table.rows) > 1:
                x.table.rows[1].height = max(int(E_PT * 20), x.height - x.table.rows[0].height)


def fit_board(slide):
    """보드 장을 화면 안에 맞춤"""
    sh = board_shape(slide)
    if sh is None:
        return []
    t = sh.table
    h0 = _board_height(t)
    h1 = rebalance_columns(t)
    pre = [f"열 너비 재배분(예상 높이 {h0:.0f}→{h1:.0f}pt)"] if h1 < h0 - 1 else []
    boxes = [x for x in slide.shapes if x.has_table and x.table.cell(0, 0).text.strip().startswith(("Monthly Plan", "Monthly Review"))]
    below = [x for x in slide.shapes if x is not sh and x.top > sh.top + 10 * E_PT
             and x.left < sh.left + sh.width and x.left + x.width > sh.left]
    if not boxes or not below:
        return []
    first_below = min(x.top for x in below) / E_PT
    limit = first_below - sh.top / E_PT
    box_bottom = max(x.top + x.height for x in boxes) / E_PT
    box_head = {id(x): x.table.rows[0].height / E_PT for x in boxes}
    box_top0 = {id(x): x.top / E_PT for x in boxes}

    b_sz0 = _cells_size(t)
    x_sz0 = {id(x): _size_pt(x.table.cell(1, 0)._tc) for x in boxes}
    changes = []
    # 글자보다 문단 간격 먼저 축소
    need_b0 = _board_height(t)
    tight = any(_box_need(x) > box_bottom - box_top0[id(x)] - box_head[id(x)] for x in boxes)
    clamp = (lambda v: v) if tight else (lambda v: max(0.0, v))
    shift0 = clamp(need_b0 + 6 - limit)
    for lv in (1, 2):
        if all(_box_need(x) <= box_bottom - (box_top0[id(x)] + shift0) - box_head[id(x)] for x in boxes):
            break
        if any(tighten_spacing(x.table.cell(1, 0)._tc.txBody, lv) for x in boxes):
            changes.append(f"Plan/Review 문단 간격 줄임({lv}단계)")
    best = None
    step = 0
    n_box = int(max(0, max(x_sz0.values()) - MIN_BOX_PT) / 0.5)
    while True:
        x_sz = {k: max(MIN_BOX_PT, v - 0.5 * step) for k, v in x_sz0.items()}
        b_sz = b_sz0   # 보드 글자는 실측에서 넘칠 때만 줄임
        need_b = _board_height(t, b_sz)
        shift = clamp(need_b + 6 - limit)
        ok = True
        for x in boxes:
            avail = box_bottom - (box_top0[id(x)] + shift) - box_head[id(x)]
            if _box_need(x, x_sz[id(x)]) > avail:
                ok = False
        best = (b_sz, x_sz, shift)
        if ok or all(v == MIN_BOX_PT for v in x_sz.values()):
            break
        step += 1
    b_sz, x_sz, shift = best
    if b_sz < b_sz0:
        for r in range(1, len(t.rows)):
            for c in range(len(t.columns)):
                _set_size(t.cell(r, c)._tc.txBody, b_sz)
        changes.append(f"보드 글자 {b_sz0:g}→{b_sz:g}pt")
    for x in boxes:
        if x_sz[id(x)] < x_sz0[id(x)]:
            _set_size(x.table.cell(1, 0)._tc.txBody, x_sz[id(x)])
            changes.append(f"{x.table.cell(0, 0).text.strip().split('(')[0].strip()} 글자 {x_sz0[id(x)]:g}→{x_sz[id(x)]:g}pt")
    if abs(shift) > 1:
        _shift_down(below, boxes, shift)
        changes.append(f"보드 공간 +{shift:.0f}pt (Plan/Review를 아래로)" if shift > 0
                       else f"보드 아래 빈 공간 {-shift:.0f}pt를 Plan/Review에 줌")
    return pre + changes


def fit_textboxes(slide, slide_h_pt):
    """글상자가 꼬리말 위로 넘치면 글자 축소"""
    changes = []
    for x in slide.shapes:
        if not (x.has_text_frame and x.text_frame.text.strip()) or x.has_table:
            continue
        paras = [p.text for p in x.text_frame.paragraphs]
        sz0 = sz = _size_pt(x.text_frame)
        room = max(x.height / E_PT, slide_h_pt - 24 - x.top / E_PT)
        while _text_height_pt(paras, x.width / E_PT - 14, sz, 1.3) > room and sz > MIN_BOX_PT:
            sz -= 0.5
        if sz < sz0:
            _set_size(x.text_frame._txBody, sz)
            changes.append(f"'{x.text_frame.text.strip()[:12]}' 글자 {sz0:g}→{sz:g}pt")
    return changes


def fit_deck(prs):
    log = []
    for i, s in enumerate(prs.slides, 1):
        ch = fit_board(s) + fit_textboxes(s, prs.slide_height / E_PT)
        if ch:
            log.append(f"슬라이드 {i}: " + ", ".join(ch))
    return log


# --- 실제 배치로 맞추기

def _tkey(x):
    return x.table.cell(0, 0).text.strip()[:14]


def measure_tables(prs):
    """LibreOffice 로 표 실제 높이(pt). 없으면 None"""
    import shutil
    import subprocess
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    tmp = tempfile.mkdtemp(prefix="okr_fit_")
    src = os.path.join(tmp, "in.pptx")
    prs.save(src)
    out = os.path.join(tmp, "out")
    try:
        subprocess.run([soffice, "--headless", "--convert-to", "pptx", "--outdir", out, src],
                       capture_output=True, timeout=180)
        laid = Presentation(os.path.join(out, "in.pptx"))
    except Exception:
        return None
    res = {}
    for i, s in enumerate(laid.slides, 1):
        for x in s.shapes:
            if x.has_table:
                res[(i, _tkey(x))] = x.height / E_PT
    return res


def _cells_size(t):
    """보드 본문 칸의 현재 글자 크기"""
    szs = [_size_pt(t.cell(r, c)._tc, None) for r in range(1, len(t.rows)) for c in range(len(t.columns))
           if t.cell(r, c).text.strip()]
    szs = [v for v in szs if v]
    return max(szs) if szs else 9.0


def _box_size(x):
    return max([_size_pt(x.table.cell(1, 0)._tc)] or [9.0])


def _cell_need(t, r, c, widths):
    cell = t.cell(r, c)
    tcPr = cell._tc.find(A("tcPr"))
    g = (lambda k, dv: int(tcPr.get(k, dv)) / E_PT if tcPr is not None else dv / E_PT)
    span_w = sum(widths[c:c + (cell.span_width if cell.is_merge_origin else 1)])
    return _body_height(cell._tc.txBody, span_w - g("marL", 91440) - g("marR", 91440)) * 1.03 + g("marT", 45720) + g("marB", 45720)


def rebalance_columns(t, groups=((2,), (3, 4, 5, 6), (7,)), step=4.0, min_w=30.0):
    """보드 열 너비를 묶음 단위로 재분배"""
    widths = [c.width / E_PT for c in t.columns]
    for gp in groups:
        avg = sum(widths[k] for k in gp) / len(gp)
        for k in gp:
            widths[k] = avg

    def score(ws):
        for k, c in enumerate(t.columns):
            c.width = int(ws[k] * E_PT)
        total, cells = 0.0, 0.0
        for r in range(len(t.rows)):
            row = t.rows[r].height / E_PT
            for c in range(len(t.columns)):
                cell = t.cell(r, c)
                if cell.is_spanned:
                    continue
                need = _cell_need(t, r, c, ws)
                if cell.is_merge_origin and cell.span_height > 1:
                    need /= cell.span_height
                row = max(row, need)
                if r > 0 and any(c in gp for gp in groups):
                    cells += need
            total += row
        return total + 0.3 * cells, total

    best, h = score(widths)
    for _ in range(300):
        cand = None
        for st in (step, step * 3, step * 8, step * 16):   # 줄 바꿈은 계단식이라 큰 걸음도 시도
            for a in groups:
                for b in groups:
                    if a is b or min(widths[k] for k in a) - st / len(a) < min_w:
                        continue
                    ws = widths[:]
                    for k in a:
                        ws[k] -= st / len(a)
                    for k in b:
                        ws[k] += st / len(b)
                    sc, hh = score(ws)
                    if sc < best - 0.3 and (cand is None or sc < cand[0]):
                        cand = (sc, hh, ws)
        if cand is None:
            break
        best, h, widths = cand
    score(widths)
    return h


def balance_boxes(slide, m, i):
    """Plan, Review 경계를 옮긴다. 옮기면 True"""
    boxes = sorted([x for x in slide.shapes if x.has_table and _tkey(x).startswith(("Monthly Plan", "Monthly Review"))],
                   key=lambda x: x.left)
    if len(boxes) != 2:
        return False
    L, R = boxes

    def ov(x):  # 넘침(+)은 실측, 여유(-)는 추정
        o = m.get((i, _tkey(x)), x.height / E_PT) - x.height / E_PT
        return o if o > 3 else _box_need(x) - (x.height / E_PT - x.table.rows[0].height / E_PT)
    over = {id(L): ov(L), id(R): ov(R)}
    if over[id(R)] > 1 and over[id(L)] < -30 and _tw(L) > 300:
        give = 24
    elif over[id(L)] > 1 and over[id(R)] < -30 and _tw(R) > 300:
        give = -24
    else:
        return False
    g = int(give * E_PT)
    L.table.columns[0].width -= g
    R.left -= g
    R.table.columns[0].width += g
    for x in slide.shapes:   # Review 머리 띠도 같이 이동
        if x is not R and not x.has_table and abs(x.left - (R.left + g)) < E_PT * 2 and abs(x.top - R.top) < E_PT * 30:
            x.left -= g
            x.width += g
    return True


def fit_measured(prs, max_iter=8):
    """실측 높이로 보드 장 맞춤"""
    log, left = [], []
    for it in range(max_iter):
        BOX_OVER.clear()
        m = measure_tables(prs)
        if m is None:
            return log, ["LibreOffice가 없어 실제 배치 확인을 건너뜀"]
        left = []
        changed = False
        for i, s in enumerate(prs.slides, 1):
            sh = board_shape(s)
            if sh is None:
                continue
            t = sh.table
            boxes = [x for x in s.shapes if x.has_table and _tkey(x).startswith(("Monthly Plan", "Monthly Review"))]
            below = [x for x in s.shapes if x is not sh and x.top > sh.top + 10 * E_PT
                     and x.left < sh.left + sh.width and x.left + x.width > sh.left]
            if not boxes or not below:
                continue
            B = m.get((i, _tkey(sh)), sh.height / E_PT)
            limit = min(x.top for x in below) / E_PT - sh.top / E_PT - 3
            over_b = B - limit
            over_x = max(m.get((i, _tkey(x)), x.height / E_PT) - x.height / E_PT for x in boxes)
            # 여유는 추정으로 잰다 (실측은 칸보다 안 작음)
            slack_x = min(x.height / E_PT - x.table.rows[0].height / E_PT - _box_need(x) for x in boxes) if over_x <= 3 else 0
            b_sz = _cells_size(t)
            x_sz = max(_box_size(x) for x in boxes)
            if over_b <= 1 and over_x <= 3:     # LibreOffice 반올림 2pt 는 넘침 아님
                continue
            if over_x > 3 and balance_boxes(s, m, i):
                log.append(f"슬라이드 {i}: Plan/Review 폭 다시 나눔")
                changed = True
                continue
            if over_b > 1 and slack_x > 8:
                dpt = min(over_b + 2, slack_x - 4)
                _shift_down(below, boxes, dpt)
                log.append(f"슬라이드 {i}: 보드 공간 +{dpt:.0f}pt (Plan/Review를 아래로)")
                changed = True
            elif over_b > 1 and b_sz > MIN_BOARD_PT:
                for r in range(1, len(t.rows)):
                    for c in range(len(t.columns)):
                        _set_size(t.cell(r, c)._tc.txBody, b_sz - 0.5)
                log.append(f"슬라이드 {i}: 보드 글자 {b_sz:g}→{b_sz - 0.5:g}pt")
                changed = True
            elif over_x > 3 and any(tighten_spacing(x.table.cell(1, 0)._tc.txBody, 2) for x in boxes):
                log.append(f"슬라이드 {i}: Plan/Review 문단 간격 줄임")
                changed = True
            elif over_x > 3 and x_sz > MIN_BOX_PT:
                for x in boxes:
                    _set_size(x.table.cell(1, 0)._tc.txBody, x_sz - 0.5)
                log.append(f"슬라이드 {i}: Plan/Review 글자 {x_sz:g}→{x_sz - 0.5:g}pt")
                changed = True
            else:
                if over_b > 1:
                    left.append(f"슬라이드 {i}: 보드 표가 최소 글자({MIN_BOARD_PT}pt)에서도 {over_b:.0f}pt 넘침 — 1W~4W·주요결과물 문구를 줄여야 함")
                if over_x > 3:
                    BOX_OVER.append((s, {_tkey(x): m.get((i, _tkey(x)), 0) - x.height / E_PT for x in boxes}))
                    left.append(f"슬라이드 {i}: Plan/Review가 최소 글자({MIN_BOX_PT}pt)에서도 {over_x:.0f}pt 넘침 — 줄 수를 줄여야 함")
        if not changed:
            break
    return log, left


# --- Plan/Review 를 다음 장으로

def _boxes(slide):
    return [x for x in slide.shapes if x.has_table and _tkey(x).startswith(("Monthly Plan", "Monthly Review"))]


def _clone_slide(prs, src):
    """같은 레이아웃 새 장에 도형 복사 (그림 제외)"""
    from pptx.opc.packuri import PackURI
    new = prs.slides.add_slide(src.slide_layout)
    used = {str(p.partname) for p in prs.part.package.iter_parts() if p is not new.part}
    n = 1
    while f"/ppt/slides/slide{n}.xml" in used:
        n += 1
    new.part.partname = PackURI(f"/ppt/slides/slide{n}.xml")
    tree = new.shapes._spTree
    for ph in list(new.placeholders):
        ph._element.getparent().remove(ph._element)
    for el in src.shapes._spTree.iterchildren():
        tag = el.tag.split("}")[1]
        if tag in ("nvGrpSpPr", "grpSpPr", "pic"):
            continue
        tree.append(copy.deepcopy(el))
    return new


def snapshot_boxes(prs):
    """줄이기 전 Plan/Review 원본 보관"""
    snap = {}
    for s in prs.slides:
        for x in _boxes(s):
            snap[(s.slide_id, _tkey(x))] = copy.deepcopy(x.table.cell(1, 0)._tc.txBody)
        if board_shape(s) is not None:
            snap[(s.slide_id, "board_sz")] = _cells_size(board_shape(s).table)
    return snap


def fit_box_alone(x, avail):
    """한 칸만 있는 장: 간격, 글자 순으로 줄임"""
    body = x.table.cell(1, 0)._tc.txBody
    lv = 1
    while _box_need(x) > avail:
        if lv <= 3 and tighten_spacing(body, lv):
            lv += 1
            continue
        lv += 1
        sz = _box_size(x)
        if sz <= MIN_BOARD_PT:
            break
        _set_size(body, sz - 0.5)


def move_boxes_to_next(prs, slide, snap):
    """Plan/Review 를 통째로 다음 장으로 이동"""
    cont = _clone_slide(prs, slide)
    sh = board_shape(cont)
    top = sh.top
    sh._element.getparent().remove(sh._element)
    boxes_c = _boxes(cont)
    box_top = min(b.top for b in boxes_c)
    bottom = max(b.top + b.height for b in boxes_c)
    d = box_top - top
    labels = [x for x in cont.shapes if x.top > top and x.top < box_top and not x.has_table]
    lab_h = max([x.height for x in labels] or [0])
    for x in labels:
        x.top = top
    for x in list(cont.shapes):
        if x.top >= box_top - E_PT * 3:
            same_bottom = abs((x.top + x.height) - bottom) < E_PT * 6
            x.top = top + lab_h + int(E_PT * 2) + (x.top - box_top)
            if same_bottom:
                x.height = bottom - x.top
    for x in _boxes(cont):
        x.table.rows[1].height = x.height - x.table.rows[0].height
        key = (slide.slide_id, _tkey(x))
        if key in snap:
            tc = x.table.cell(1, 0)._tc
            tc.replace(tc.txBody, copy.deepcopy(snap[key]))
        fit_box_alone(x, x.height / E_PT - x.table.rows[0].height / E_PT - 6)
    bt = board_shape(slide)
    b_bottom = bt.top + int(E_PT * 10)
    for x in list(slide.shapes):
        if x is not bt and x.top > b_bottom and x.left < bt.left + bt.width and x.left + x.width > bt.left:
            x._element.getparent().remove(x._element)
    t = bt.table
    sz0 = snap.get((slide.slide_id, "board_sz"))
    room = (bottom - bt.top) / E_PT
    while sz0 and _cells_size(t) < sz0 and _board_height(t, _cells_size(t) + 0.5) <= room:
        nsz = _cells_size(t) + 0.5
        for r in range(1, len(t.rows)):
            for c in range(len(t.columns)):
                _set_size(t.cell(r, c)._tc.txBody, nsz)
    order = list(prs.slides)
    order.remove(cont)
    order.insert(order.index(slide) + 1, cont)
    reorder(prs, order)
    return cont


# --- 표 정리 (표 깨짐 방지)
_BORDERS = ("lnL", "lnR", "lnT", "lnB")


def unify_borders(t, ref=(1, 2)):
    """기준 칸 테두리를 모든 칸에 똑같이"""
    src = t.cell(*ref)._tc.find(A("tcPr"))
    if src is None:
        return 0
    lines = {tag: src.find(A(tag)) for tag in _BORDERS}
    if any(v is None for v in lines.values()):
        return 0
    fixed = 0
    for r in range(len(t.rows)):
        for c in range(len(t.columns)):
            tcPr = t.cell(r, c)._tc.get_or_add_tcPr()
            for k, tag in enumerate(_BORDERS):
                if tcPr.find(A(tag)) is None:
                    # 스키마 순서: lnL, lnR, lnT, lnB 가 맨 앞
                    pos = sum(1 for t2 in _BORDERS[:k] if tcPr.find(A(t2)) is not None)
                    tcPr.insert(pos, copy.deepcopy(lines[tag]))
                    fixed += 1
    return fixed


def settle_table(x, need_rows=False):
    """틀 크기를 실제 표 크기와 같게"""
    t = x.table
    if need_rows:
        ws = [c.width / E_PT for c in t.columns]
        need = [t.rows[r].height / E_PT for r in range(len(t.rows))]
        for r in range(len(t.rows)):
            for c in range(len(t.columns)):
                cell = t.cell(r, c)
                if cell.is_spanned or (cell.is_merge_origin and cell.span_height > 1):
                    continue
                need[r] = max(need[r], _cell_need(t, r, c, ws))
        for r in range(len(t.rows)):        # 세로 병합 칸: 모자라면 마지막 행에 더함
            for c in range(len(t.columns)):
                cell = t.cell(r, c)
                if cell.is_merge_origin and cell.span_height > 1:
                    h = sum(need[r:r + cell.span_height])
                    n = _cell_need(t, r, c, ws)
                    if n > h:
                        need[r + cell.span_height - 1] += n - h
        for r, h in enumerate(need):
            t.rows[r].height = int(math.ceil(h) * E_PT)
    x.width = sum(c.width for c in t.columns)
    x.height = sum(r.height for r in t.rows)


def settle_deck(prs):
    log = []
    for i, s in enumerate(prs.slides, 1):
        sh = board_shape(s)
        if sh is not None:
            n = unify_borders(sh.table)
            if n:
                log.append(f"슬라이드 {i}: 보드 테두리 {n}곳 보완")
            settle_table(sh, need_rows=True)
        for x in _boxes(s):
            settle_table(x)
    return log


def table_problems(prs, gap_pt=6):
    """테두리 빠짐, 틀 크기, 간격 부족 검사"""
    out = []
    for i, s in enumerate(prs.slides, 1):
        sh = board_shape(s)
        if sh is None:
            continue
        t = sh.table
        miss = sum(1 for r in range(len(t.rows)) for c in range(len(t.columns))
                   if not t.cell(r, c).is_spanned and any(
                       (t.cell(r, c)._tc.find(A("tcPr")) is None) or t.cell(r, c)._tc.find(A("tcPr")).find(A(tag)) is None
                       for tag in _BORDERS))
        if miss:
            out.append(f"슬라이드 {i}: 보드 테두리 빠진 칸 {miss}개")
        for x in [sh] + _boxes(s):
            gw = sum(c.width for c in x.table.columns)
            gh = sum(r.height for r in x.table.rows)
            if abs(x.width - gw) > E_PT * 2 or abs(x.height - gh) > E_PT * 2:
                out.append(f"슬라이드 {i}: '{_tkey(x)}' 표 틀 크기가 실제 표와 다름")
        below = [x for x in s.shapes if x is not sh and x.top > sh.top + 10 * E_PT
                 and x.left < sh.left + sh.width and x.left + x.width > sh.left]
        if below and sh.top + sh.height + gap_pt * E_PT > min(x.top for x in below) + E_PT:
            out.append(f"슬라이드 {i}: 보드 표와 아래 칸 간격 {(min(x.top for x in below) - sh.top - sh.height) / E_PT:.0f}pt (<{gap_pt}pt)")
    return out


# --- 모든 장: 꼬리말 위로
FOOTER_WORDS = ("Standard", "Form", "INTERX", "IX-SF-001(R0)")
MIN_TEXT_PT = 7.5


def _content_limit(prs):
    return prs.slide_height / E_PT - 22      # 꼬리말 선 위


def _text_shapes(slide):
    return [x for x in slide.shapes if x.has_text_frame and x.text_frame.text.strip() and not x.has_table]


def _shrink_text(x, step=0.5):
    body = x.text_frame._txBody
    sizes = [int(r.get("sz")) / 100 for r in body.iter(A("rPr")) if r.get("sz")]
    if not sizes or max(sizes) - step < MIN_TEXT_PT:
        return False
    for el in list(body.iter(A("rPr"))) + list(body.iter(A("endParaRPr"))):
        if el.get("sz"):
            el.set("sz", str(int(max(MIN_TEXT_PT, int(el.get("sz")) / 100 - step) * 100)))
    for el in body.iter(A("buSzPts")):
        el.set("val", str(int(max(MIN_TEXT_PT * 100, int(el.get("val")) - step * 100))))
    return True


def fit_pages_estimate(prs):
    """추정으로 1차 맞춤"""
    log = []
    limit = _content_limit(prs)
    for i, s in enumerate(prs.slides, 1):
        for x in _text_shapes(s):
            if x.top / E_PT > limit:
                continue
            room = limit - x.top / E_PT - 7.2
            w = x.width / E_PT - 14.4
            body = x.text_frame._txBody
            acts = []
            lv = 1
            while _body_height(body, w) * 0.9 > room:      # 추정은 실제보다 약 10% 크게 나옴
                if lv <= 2 and tighten_spacing(body, lv):
                    acts.append(f"간격 {lv}단계")
                    lv += 1
                    continue
                lv += 1
                if not _shrink_text(x):
                    break
                acts.append("글자 -0.5")
            if acts:
                log.append(f"슬라이드 {i}: '{x.text_frame.text.strip()[:12]}' " + ", ".join(dict.fromkeys(acts)))
    return log


def _pdf_overflow(prs, workdir):
    """실제 렌더(PDF)에서 넘친 글자 위치"""
    import subprocess
    src = os.path.join(workdir, "fit.pptx")
    pdf = os.path.join(workdir, "fit.pdf")
    prs.save(src)
    rpy = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "scripts", "render.py")
    if os.path.exists(pdf):
        os.remove(pdf)   # 렌더가 실패하면 지난 회차 PDF 를 읽지 않게
    try:
        import pdfplumber
        r = subprocess.run([sys.executable, "-P", rpy, src, os.path.join(workdir, "png"), "--slides", "1", "--pdf", pdf],
                           capture_output=True, timeout=240)
        if r.returncode:
            return None
        doc = pdfplumber.open(pdf)
    except Exception:
        return None
    limit = _content_limit(prs)
    out = {}
    for n, page in enumerate(doc.pages, 1):
        sc = (prs.slide_height / E_PT) / float(page.height)
        for wd in page.extract_words():
            if wd["text"] in FOOTER_WORDS:
                continue
            top, bottom = wd["top"] * sc, wd["bottom"] * sc
            if bottom > limit + 1 or wd["x1"] * sc > prs.slide_width / E_PT + 1:
                out.setdefault(n, []).append((wd["x0"] * sc, top))
    doc.close()
    return out


def fit_pages_measured(prs, max_iter=5):
    """실제 렌더로 확인하며 글자 축소"""
    log, left = [], []
    work = tempfile.mkdtemp(prefix="okr_page_")
    for _ in range(max_iter):
        over = _pdf_overflow(prs, work)
        if over is None:
            return log, ["렌더 실패나 pdfplumber 없음: 화면 넘침 실측을 건너뜀"]
        if not over:
            return log, []
        changed = False
        left = []
        slides = list(prs.slides)
        for n, words in over.items():
            s = slides[n - 1]
            owners = {}
            for wx, wy in words:
                cands = [x for x in s.shapes if (x.has_text_frame or x.has_table)
                         and x.left / E_PT - 2 <= wx <= (x.left + max(x.width, sum(c.width for c in x.table.columns) if x.has_table else 0)) / E_PT + 2
                         and x.top / E_PT <= wy]
                if cands:
                    o = max(cands, key=lambda x: x.top)
                    owners[o.shape_id] = o
            for x in owners.values():
                name = _tkey(x) if x.has_table else x.text_frame.text.strip()[:12]
                if x.has_table:
                    ok = False
                    for row in x.table.rows:
                        for c in row.cells:
                            sz = _size_pt(c._tc, None)
                            if sz and sz - 0.5 >= MIN_TEXT_PT:
                                _set_size(c._tc.txBody, sz - 0.5)
                                ok = True
                else:
                    ok = tighten_spacing(x.text_frame._txBody, 2) or _shrink_text(x)
                if ok:
                    changed = True
                    log.append(f"슬라이드 {n}: '{name}' 화면 넘침 → 글자/간격 줄임")
                else:
                    left.append(f"슬라이드 {n}: '{name}' 최소 글자에서도 화면 넘침 — 내용을 줄여야 함")
        if not changed:
            break
    return log, left


def check(path, style="plain"):
    prs = Presentation(path)
    problems, notes = [], []
    boards = [(i, s) for i, s in enumerate(prs.slides, 1) if board_table(s) is not None]
    notes.append(f"슬라이드 {len(prs.slides)}장, 보드 {len(boards)}장, Appendix {sum(is_appendix(s) for s in prs.slides)}장")
    if len(boards) != 4:
        problems.append(f"보드가 4장이 아님 ({len(boards)})")
    periods = [period_of(s) for _, s in boards]
    if periods != sorted(periods):
        problems.append(f"보드가 월 순서가 아님: {periods}")
    for i, s in boards:
        per = period_of(s)
        rev = table_by_head(s, "Monthly Review")
        t = board_table(s)
        ach = [t.cell(r, 8).text.strip() for r in range(1, len(t.rows))]
        body = rev.cell(1, 0).text.strip()
        for r in range(1, len(t.rows)):
            for c in range(3, 9):
                for p in t.cell(r, c)._tc.txBody.findall(A("p")):
                    pPr = p.find(A("pPr"))
                    if pPr is None or pPr.get("algn") != "ctr":
                        problems.append(f"슬라이드 {i}: 보드 {r}행 {c + 1}열 가운데 정렬 아님")
                        break
        is_next = per == max(periods)
        label = f"슬라이드 {i} ({kind_of(s)}, {per[1]}월)"
        if is_next:
            if body:
                problems.append(f"{label}: 익월인데 Review 가 비어 있지 않음")
            if any(ach):
                problems.append(f"{label}: 익월인데 달성현황이 비어 있지 않음")
        else:
            if not body:
                problems.append(f"{label}: 당월인데 Review 가 비어 있음")
            if not all(ach) and style != "own":      # own: 달성현황을 비워 두는 문체 허용
                problems.append(f"{label}: 달성현황 빈칸 {ach}")
            if find_cover(s) is not None:
                problems.append(f"{label}: Review 가리개가 남아 있음")
            cap_em = (5797500 / 914400 - 0.7) * 72 / 9
            for para in rev.cell(1, 0).text_frame.paragraphs:
                if style != "own" and kind_of(s) == "general" and para.text.strip() and not para.text.startswith("[") and em(para.text) > cap_em:
                    problems.append(f"{label}: 한 줄을 넘음 ({em(para.text):.0f}em > {cap_em:.0f}em): {para.text[:30]}")
    # 티켓 번호, 미완 낱말, 달성현황 16자
    for i, s in boards:
        per = period_of(s)
        if per == max(periods):
            continue
        body = table_by_head(s, "Monthly Review").cell(1, 0).text
        m = UNDONE.search(body) if style != "own" else None
        if m:
            problems.append(f"슬라이드 {i}: 실적(Review)에 '{m.group(0)}' — 한 일만 쓰고 못 한 일은 익월 계획에 이월 표시")
        t = board_table(s)
        for r in range(1, len(t.rows)):
            a = t.cell(r, 8).text.strip()
            if len(a) > ACH_MAX:
                problems.append(f"슬라이드 {i}: 달성현황 {r}행 {len(a)}자 > {ACH_MAX}자 ({a})")
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            for ln in _lines_of(sh):
                m = TICKET.search(ln)
                if m and not re.match(r"^(?:IX-SF|KR\d-\d|CO\d)", m.group(0)):
                    problems.append(f"슬라이드 {i}: 티켓 번호로 보이는 표기 '{m.group(0)}' — 업무 언어로 풀어 쓰기")
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            for ln in _lines_of(sh):
                if any(tok in ln for tok in PLACEHOLDERS):
                    problems.append(f"슬라이드 {i}: 자리표시자 남음 -> {ln.strip()[:60]}")
        if is_appendix(s):
            pics = [sh for sh in s.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]
            if len(pics) != 1:
                problems.append(f"슬라이드 {i}: Appendix 그림 {len(pics)}개")
            else:
                iw, ih = pics[0].image.size
                notes.append(f"슬라이드 {i}: 이미지 {iw}x{ih}px")
                if iw < 1600:
                    problems.append(f"슬라이드 {i}: 해상도 낮음 ({iw}px)")
            panel = next((sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.startswith("핵심 개요")), None)
            if panel is None:
                problems.append(f"슬라이드 {i}: 핵심 개요 패널 없음")
            else:
                inner = (panel.width / 914400 - 0.3 - 0.19) * 72 / 10
                lines = sum(max(1, math.ceil(em(p.text) / inner)) for p in panel.text_frame.paragraphs[1:])
                if lines * 10 * 1.3 + len(panel.text_frame.paragraphs) * 9 + 30 > panel.height / 12700:
                    problems.append(f"슬라이드 {i}: 핵심 개요 글이 패널을 넘을 수 있음 ({lines}줄)")
    problems += layout_problems(prs)
    problems += table_problems(prs)
    print("\n".join(notes))
    if problems:
        print("\nFAIL")
        print("\n".join(" - " + p for p in problems))
        return 1
    print("\nPASS")
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build")
    b.add_argument("template")
    b.add_argument("spec")
    b.add_argument("out")
    c = sub.add_parser("check")
    c.add_argument("deck")
    c.add_argument("--style", default="plain", help="own = 사용자 덱 문체(상태 표시 허용)")
    a = ap.parse_args()
    sys.exit(build(a.template, a.spec, a.out) or 0 if a.cmd == "build" else check(a.deck, a.style))
