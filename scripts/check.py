"""덱 검사. 사용: check.py <덱.pptx> [장 수] [--soft]"""
import os
import re
import sys
import zipfile

from lxml import etree

import math

sys.stdout.reconfigure(errors="replace")  # cp949 콘솔에서 멈추지 않게

args = [a for a in sys.argv[1:] if not a.startswith("--")]
SOFT = "--soft" in sys.argv
path = args[0]
want = int(args[1]) if len(args) > 1 else None
ALLOWED = {5.44, 6.35, 7.26, 8.0, 8.16, 9.0, 9.07, 10.0, 10.5, 11.0, 12.0, 14.0, 16.32, 20.0}
FONT = "Noto Sans KR"
FOOTER_Y = 7.215      # 꼬리말 선 (바닥)
LINE_H = 1.2 * 1.25   # 한 줄 높이 = 글자 x 1.2 x 줄 간격 1.25


def _em(ch):  # 굵은 글자 폭 어림 (렌더로 잼)
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


def tw(text, size):
    return sum(_em(ch) for ch in text) * size / 72


def lines_needed(tf, width, default):
    n, top = 0, 0.0
    for para in tf.paragraphs:
        sz = next((r.font.size.pt for r in para.runs if r.font.size), default)
        ls = para.line_spacing if isinstance(para.line_spacing, float) else 1.25
        whole = para.text.replace("\x0b", "\n").replace("\v", "\n")   # a:br 도 줄로 셈
        for txt in whole.split("\n"):
            k = max(1, math.ceil(tw(txt, sz) / max(width * 1.03, 0.05))) if txt else 1
            n += k
            top += k * sz * 1.2 * ls / 72
    return n, top


def widow(txt, size, width, k=1.0):
    """마지막 줄에 짧은 낱말 하나만 남으면 그 낱말"""
    words, lines, cur = txt.split(), [], ""
    for wd in words:
        cand = (cur + " " + wd).strip()
        if cur and tw(cand, size) * k > width * 1.03:
            lines.append(cur)
            cur = wd
        else:
            cur = cand
    return cur if lines and " " not in cur and tw(cur, size) * k < width * 0.25 else None


def frame_widows(tf, width):
    for para in tf.paragraphs:
        sz = next((r.font.size.pt for r in para.runs if r.font.size), 7.26)
        k = 1.0 if any(r.font.bold for r in para.runs) else 0.88   # 보통 글자는 굵은 글자보다 좁음
        for txt in para.text.replace("\x0b", "\n").replace("\v", "\n").split("\n"):
            w = widow(txt, sz, width, k)
            if w:
                yield txt, w


z = zipfile.ZipFile(path)
bad = []
pres = z.read("ppt/presentation.xml").decode("utf8")
sz = re.search(r"<p:sldSz[^>]*>", pres).group(0)
cx, cy = (int(re.search(rf'{k}="(\d+)"', sz).group(1)) for k in ("cx", "cy"))
if (cx, cy) != (12192000, 6858000):
    bad.append(f"장 크기 {cx / 914400:.2f} x {cy / 914400:.2f} in (13.33 x 7.5 아님)")
slides = sorted((n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                key=lambda n: int(re.findall(r"\d+", n)[0]))
_stem = os.path.splitext(os.path.basename(path))[0]
if not re.match(r"^\d{8}_[^_\s]+_[^_\s]+_v\d+\.\d+\.\d+$", _stem):
    bad.append(f"파일 이름 '{_stem}' 이 규칙(YYYYMMDD_파일명_작성자_v0.1.0) 아님: d.filename(파일명, 버전)")
if want is not None and len(slides) != want:
    bad.append(f"장 수 {len(slides)} (기대 {want})")
for n in slides:
    x = z.read(n).decode("utf8")
    tag = n.split("/")[-1]
    c, st, en = x.count("<p:cxnSp>"), x.count("<a:stCxn"), x.count("<a:endCxn")
    if not (c == st == en):
        bad.append(f"{tag}: 연결선 {c}, 시작 붙음 {st}, 끝 붙음 {en}")
    ids = re.findall(r'<p:cNvPr id="(\d+)"', x)
    if len(ids) != len(set(ids)):
        bad.append(f"{tag}: 도형 번호 중복 {sorted({i for i in ids if ids.count(i) > 1})}")
    fonts = {f for k, f in re.findall(r'<a:(latin|ea|cs|sym|buFont|font)[^>]*typeface="([^"+][^"]*)"', x)
             if f != FONT and not (k in ("sym", "buFont") and f.startswith(("Wingdings", "Symbol", "Webdings")))}
    if fonts:
        bad.append(f"{tag}: {FONT} 아닌 글꼴 {sorted(fonts)} (굵기 이름 글꼴도 안 됨, 굵게 켬으로)")
    A_ = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    off = 0
    for p_ in etree.fromstring(x.encode("utf8")).iter(A_ + "p"):
        v = p_.find(f"{A_}pPr/{A_}lnSpc/{A_}spcPct")
        off += v is None or v.get("val") not in ("125000", "150000")
    if off:
        bad.append(f"{tag}: 줄 간격이 1.25 또는 1.5 가 아닌 문단 {off}개")
    for body in re.findall(r'<p:cNvSpPr txBox="1"/>.*?<a:bodyPr([^>]*)>', x):
        if 'anchor="ctr"' not in body:
            bad.append(f"{tag}: 세로 중간 맞춤이 아닌 글상자 있음")
            break
    svg_only = len(re.findall(r"<a:blip(?![^>]*r:embed)[^>]*>", x))
    if svg_only:
        bad.append(f"{tag}: SVG 만 있는 그림 {svg_only}개. make_template.py --png-fallback 로 고침")
    for leftover in ("동도금 1호기", "MeasurementCurrentDensity", "Asset Administration Shell(AAS) 기반"):
        if leftover in x:
            bad.append(f"{tag}: 템플릿 예시 글자 남음 '{leftover}'")
from pptx import Presentation  # noqa: E402

E = 914400
sections = {}
_all = list(Presentation(path).slides)
_names = [sl.slide_layout.name for sl in _all]
if len(_all) < 3 or _names[0] != "3_사용자 지정 레이아웃" or _names[1] != "5_contents-02" or _names[-1] != "앤드페이지_다크네이비":
    bad.append(f"장 순서: 표지, 목차, 본문, Thank you 여야 함 (지금 {_names[:2]} ... {_names[-1:]})")
_toc_text = _all[1].shapes and " ".join(sh.text_frame.text for sh in _all[1].shapes if sh.has_text_frame) if len(_all) > 1 else ""
for _sl in _all[2:-1]:   # 본문 절 이름이 목차에 있어야 함
    _big = next((q.text_frame.text.strip() for q in _sl.placeholders if q.placeholder_format.idx == 10), "")
    _nm = _big.split(" ", 1)[1] if " " in _big else _big
    if _nm and _nm not in _toc_text:
        bad.append(f"목차에 절 '{_nm}' 이 없음 (d.section, d.page 로 만든다)")
        break
for i, sl in enumerate(_all, 1):
    for sh in list(sl.shapes) + list(sl.slide_layout.shapes):  # 16:9 틀 안
        if sh.width is None:
            continue
        if sh.left < -E * 0.005 or sh.top < -E * 0.005 or sh.left + sh.width > cx + E * 0.005 or sh.top + sh.height > cy + E * 0.005:
            bad.append(f"slide{i}: 도형 '{sh.name}' 이 장 밖으로 나감 ({sh.left / E:.2f}, {sh.top / E:.2f}, "
                       f"{(sh.left + sh.width) / E:.2f}, {(sh.top + sh.height) / E:.2f} in)")
    if sl.slide_layout.name in ("3_사용자 지정 레이아웃", "5_contents-02", "앤드페이지_다크네이비"):  # 표지, 목차, 끝 장
        continue
    for sh in sl.shapes:  # 번호 원 지름
        if (sh.has_text_frame and sh.text_frame.text.strip().isdigit() and sh.width is not None
                and "Oval" in sh.name and min(abs(sh.width / E - d) for d in (0.18, 0.15)) > 0.006):
            bad.append(f"slide{i}: 번호 원 '{sh.text_frame.text}' 지름 {sh.width / E:.2f} in (0.18 또는 0.15 만)")
    for sh in sl.shapes:  # 글자 크기, 넘침
        frames = [sh.text_frame] if sh.has_text_frame and not sh.is_placeholder else []
        if sh.has_table:
            frames += [c.text_frame for row in sh.table.rows for c in row.cells]
        for f_ in frames:
            for para in f_.paragraphs:
                for r in para.runs:
                    if r.font.size and r.text.strip() and round(r.font.size.pt, 2) not in ALLOWED and r.font.size.pt < 28:
                        bad.append(f"slide{i}: 글자 크기 {r.font.size.pt:g}pt '{r.text[:16]}' (기본 크기 아님)")
        if sh.has_text_frame and sh.text_frame.text.strip() and not sh.is_placeholder:
            tf = sh.text_frame
            w = sh.width / E - (tf.margin_left + tf.margin_right) / E
            n, need = lines_needed(tf, w, 7.26)
            h = sh.height / E
            if need > h + 0.06 and tf.word_wrap is not False:
                bad.append(f"slide{i}: 글자가 칸을 넘침 ({n}줄, 칸 높이 {h:.2f} in) '{tf.text[:24]}'")
            for txt, w in frame_widows(tf, w) if tf.word_wrap is not False else ():
                print(f"주의: slide{i} 마지막 줄에 '{w}' 하나만 남을 수 있음 '{txt[:24]}'")
        if sh.has_table:
            tb = sh.table
            for ri, row in enumerate(tb.rows):
                for ci, cell in enumerate(row.cells):
                    if not cell.text_frame.text.strip():
                        continue
                    w = tb.columns[ci].width / E - (cell.margin_left + cell.margin_right) / E
                    n, need = lines_needed(cell.text_frame, w, 7.26)
                    if need > row.height / E - (cell.margin_top + cell.margin_bottom) / E + 0.04:
                        bad.append(f"slide{i}: 표 칸 글자가 넘침 ({n}줄) '{cell.text_frame.text[:24]}'")
                    for txt, wd in frame_widows(cell.text_frame, w):
                        print(f"주의: slide{i} 표 칸 마지막 줄에 '{wd}' 하나만 남을 수 있음 '{txt[:24]}'")
            cws = [cl.width / E - 0.12 for cl in tb.columns]
            tall = [row.height / E for row in list(tb.rows)[1:]
                    if row.height / E > 0.42 and max(lines_needed(cl.text_frame, cws[k], 7.26)[1]
                                                     for k, cl in enumerate(row.cells)) * 2.6 < row.height / E]
            if len(tall) > len(tb.rows) // 2:  # 글보다 훨씬 높은 행은 빈 자리를 메운 티
                bad.append(f"slide{i}: 표 행이 글자보다 너무 높음 ({tall[0]:.2f} in). 행을 낮추고 블록을 더함")
    ph = {q.placeholder_format.idx: q.text_frame.text.strip() for q in sl.placeholders if q.has_text_frame}
    big, sub = ph.get(10, ""), ph.get(11, "")
    m1, m2 = re.match(r"(\d+)\. \S", big), re.match(r"(\d+)-(\d+)\. \S", sub)
    if not m1:
        bad.append(f"slide{i}: 대제목 '{big}' 이 '번호. 절 이름' 꼴 아님 (d.section, d.page)")
    if not m2:
        bad.append(f"slide{i}: 소제목 '{sub}' 이 '번호-번호. 장 주제' 꼴 아님")
    if m1 and m2 and m1.group(1) != m2.group(1):
        bad.append(f"slide{i}: 대제목 번호 {m1.group(1)} 과 소제목 번호 {m2.group(1)} 이 다름")
    if m1:
        if sections.setdefault(m1.group(1), big) != big:
            bad.append(f"slide{i}: 같은 절 {m1.group(1)} 인데 대제목이 다름 ('{sections[m1.group(1)]}' / '{big}')")
    for q in sl.placeholders:
        if q.placeholder_format.idx in (10, 11) and q.has_text_frame and q.text_frame.text:
            sz = 24 if q.placeholder_format.idx == 10 else 14
            wd = tw(q.text_frame.text, sz)
            if wd > q.width / E - 0.1:
                bad.append(f"slide{i}: 제목 '{q.text_frame.text}' 이 자리표시자 폭({q.width / E:.2f} in)을 넘어 접힘")
    if re.search(r"(에서|때|에게|으로|로)$", sub.split(" (")[0]):
        bad.append(f"slide{i}: 소제목 '{sub}' 이 끝맺지 않은 구절 (명사구로)")
    lead = [q.text_frame.text.strip() for q in sl.shapes if q.has_text_frame and not q.is_placeholder
            and abs(q.top / E - 1.35) < 0.02]
    if lead and lead[0].endswith(")"):
        bad.append(f"slide{i}: 요약 줄이 괄호로 끝남 '{lead[0][-20:]}'")
    if lead and tw(lead[0], 12) > 12.1:
        bad.append(f"slide{i}: 요약 줄이 한 줄을 넘음 (어림 {tw(lead[0], 12):.1f} in, 칸 12.17 in)")
    if sl.slide_layout.name != "본문_소제목 O":
        bad.append(f"slide{i}: 레이아웃 '{sl.slide_layout.name}' (본문은 모두 '본문_소제목 O')")
    bx = [(sh.left / E, sh.top / E, (sh.left + sh.width) / E, (sh.top + sh.height) / E) for sh in sl.shapes
          if not sh.is_placeholder and sh.width is not None and sh.width / E < 13.0 and sh.height / E < 6.5]
    if not bx:
        continue
    L, R, B = min(b[0] for b in bx), max(b[2] for b in bx), max(b[3] for b in bx)
    c = (L + R) / 2
    note = print if SOFT else bad.append
    if abs(c - 6.67) > 0.35:
        note(f"slide{i}: 내용 중심 {c:.2f} (장 가운데 6.67), 왼 여백 {L:.2f}, 오른 여백 {13.33 - R:.2f}")
    if 7.5 - B > 0.95:
        note(f"slide{i}: 아래 여백 {7.5 - B:.2f} in (도식을 아래로 늘리거나 블록을 더함)")
    if B > FOOTER_Y - 0.25:   # 꼬리말 선 위 0.3 in 비움
        bad.append(f"slide{i}: 내용 아래 {B:.2f} in 가 바닥 선({FOOTER_Y}) 위 0.3 in 여백을 침범")
    for sh in sl.shapes:   # 한 문단에 괄호 부연 셋 이상
        if sh.has_text_frame and not sh.is_placeholder:
            for para in sh.text_frame.paragraphs:
                if para.text.count("(") >= 3:
                    bad.append(f"slide{i}: 괄호 부연 {para.text.count('(')}개인 문단 '{para.text[:24]}' (핵심만 남김)")
if 'latinLnBrk="1"' in "".join(z.read(n).decode("utf8") for n in z.namelist()
                               if n.endswith(".xml") and n.startswith("ppt/") and "/media/" not in n):
    bad.append('낱말 중간 줄 바꿈(latinLnBrk="1")이 남음: 한글 낱말이 줄 끝에서 잘림')
for n in slides:  # lang 이 ko-KR 아니면 낱말 중간 줄 바꿈
    k = len(re.findall(r'<a:rPr(?![^>]*lang="ko-KR")[^>]*>', z.read(n).decode("utf8")))
    if k:
        bad.append(f"{n.split('/')[-1]}: lang 이 ko-KR 이 아닌 글 조각 {k}개 (한글 낱말이 줄 끝에서 잘림)")
for n in z.namelist():  # 레이아웃, 마스터, 테마, 차트 글꼴도 하나
    if re.match(r"ppt/(slideLayouts|slideMasters|notesMasters|notesSlides|handoutMasters|theme|charts)/[^/]+\.xml$", n):
        xx = z.read(n).decode("utf8")
        other = {f for k, f in re.findall(r'<a:(latin|ea|cs|sym|buFont|font)[^>]*typeface="([^"+]*)"', xx)
                 if f != FONT and not (k in ("sym", "buFont") and f.startswith(("Wingdings", "Symbol", "Webdings")))}
        if other:
            bad.append(f"{n.split('/')[-1]}: {FONT} 아닌 글꼴 {sorted(other)}")
for n in z.namelist():
    if re.match(r"ppt/(slideLayouts|slideMasters)/[^/]+\.xml$", n):
        k = len(re.findall(r"<a:blip(?![^>]*r:embed)[^>]*>", z.read(n).decode("utf8")))
        if k:
            bad.append(f"{n.split('/')[-1]}: SVG 만 있는 그림 {k}개 (표지 배경, 머리 장식이 깨짐). "
                       "make_template.py --png-fallback 로 고침")
body = [sl for sl in _all if sl.slide_layout.name == "본문_소제목 O"]


def shape_of(sl):
    """장 구도: (패널 수, 표, 차트, 그림, 연결선)"""
    sh = list(sl.shapes)
    return (sum(1 for s in sh if s.has_text_frame and not s.is_placeholder and "252E3F" in s._element.xml
                and s.height is not None and s.height / E < 0.32),
            any(s.has_table for s in sh), any(s.has_chart for s in sh),
            any(s.shape_type == 13 for s in sh), sum(1 for s in sh if s._element.tag.endswith("}cxnSp")) > 1)


tables = sum(1 for sl in body if any(s.has_table for s in sl.shapes))
if len(body) >= 4 and tables > len(body) * 0.35:
    print(f"주의: 표가 있는 본문 장 {tables}/{len(body)} (35% 넘음). 규칙과 이유는 글, 관계는 그림으로")
kinds = [shape_of(sl) for sl in body]
for k in range(len(kinds) - 2):
    if kinds[k] == kinds[k + 1] == kinds[k + 2]:
        print(f"주의: 본문 {k + 1}~{k + 3}장 구도가 같음. 내용 관계에 맞게 바꿈")
        break
print("\n".join(bad) if bad else f"OK {path} 장 {len(slides)}")
sys.exit(1 if bad else 0)
