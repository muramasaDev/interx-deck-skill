"""xlsx 정의서 검사 및 관문 (읽기 전용)
사용: python -P scripts/sheet_check.py 파일.xlsx [파일.xlsx ...]   (오류가 있으면 종료 코드 1, 열 수 없으면 2)
ixsheet.Book.save() 도 같은 check() 를 거친다."""
import collections
import datetime
import re
import sys
import unicodedata

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, range_boundaries
from openpyxl.utils.cell import coordinate_from_string

FONT = "맑은 고딕"
DOC_SHEETS = ["문서 정보", "개정 이력"]
LONG = 300
NUM_TEXT = re.compile(r"-?[\d,]+(\.\d+)?%?")
KINDS = ("empty", "font", "height", "numtext", "datefmt", "long", "clip", "tall", "formula", "huge")


def dw(s):
    """표시 폭 (한글 등 넓은 글자는 2칸)"""
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in str(s))


def line_count(text, width):
    """자동 줄 바꿈 때 줄 수 (단어 단위, 긴 단어는 글자 단위)"""
    usable = max(width - 2, 2)
    total = 0
    for para in str(text).split("\n"):
        n, cur = 1, 0
        for w in para.split(" "):
            need = dw(w) + (1 if cur else 0)
            if cur + need <= usable:
                cur += need
                continue
            if cur:
                n, cur = n + 1, 0
            for ch in w:
                c = dw(ch)
                if cur + c > usable:
                    n, cur = n + 1, c
                else:
                    cur += c
        total += n
    return total


def row_height(lines):
    return max(18, lines * 15 + 3)


def shown(cell):
    """셀에 보이는 글 (숫자 서식과 날짜 반영)"""
    v, nf = cell.value, cell.number_format
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        d = nf.partition(".")[2].count("0")
        if "," in nf:
            return f"{v:,.{d}f}"
        return str(v) if nf == "General" else f"{v:.{d}f}"
    return "" if v is None else str(v)


def col_widths(ws):
    """열 번호 -> 너비 (같은 너비 열이 묶여 읽히는 경우 포함)"""
    out = {}
    for cd in ws.column_dimensions.values():
        if cd.width:
            lo = cd.min or column_index_from_string(cd.index)
            for i in range(lo, (cd.max or lo) + 1):
                out[i] = cd.width
    return out


def _add(lst, where, msg, locs=()):
    more = f" 외 {len(locs) - 5}" if len(locs) > 5 else ""
    lst.append(f"[{where}] {msg}" + (f" ({', '.join(locs[:5])}{more})" if locs else ""))


def _sheet(ws, E, W):
    name = ws.title
    if len(name) > 31 or re.search(r"[\\/?*\[\]:]", name):
        _add(W, name, "시트 이름이 31자를 넘거나 금지 문자를 포함")
    if ws.auto_filter.ref:
        _, hr, nc, _ = range_boundaries(ws.auto_filter.ref)
    else:
        hr, nc = 2, ws.max_column
        _add(E, name, "필터 없음")
    fp = ws.freeze_panes
    if not fp or coordinate_from_string(fp)[1] != hr + 1:
        _add(E, name, f"틀 고정이 머리행 아래(A{hr + 1})가 아님")
    for m in ws.merged_cells.ranges:
        if not (m.min_row == m.max_row == hr - 1 and hr >= 3):
            _add(E, name, f"병합 {m.coord} (머리행 바로 위 그룹 줄 외 병합 금지)")
    widths, last = col_widths(ws), ws.max_row
    maxc = max(nc, ws.max_column)   # 행마다 다시 세면 행 수의 제곱으로 느려짐
    loc = {k: [] for k in KINDS}
    has_num = {c for c in range(1, nc + 1) if any(
        isinstance(ws.cell(r, c).value, (int, float)) and not isinstance(ws.cell(r, c).value, bool)
        for r in range(hr + 1, last + 1))}
    for r in range(1, last + 1):
        if r >= hr and ws.row_dimensions[r].height is None:
            loc["height"].append(f"{r}행")
        elif r >= hr and ws.row_dimensions[r].height > 409.5:
            loc["huge"].append(f"{r}행")
        for c in range(1, maxc + 1):
            cell = ws.cell(r, c)
            v, where = cell.value, cell.coordinate
            if v is None or (isinstance(v, str) and not v.strip()):
                if r > hr and c <= nc:
                    loc["empty"].append(where)
                continue
            if cell.data_type == "f":
                loc["formula"].append(where)
            if cell.font.name != FONT:
                loc["font"].append(where)
            if isinstance(v, str) and len(v) > LONG:
                loc["long"].append(where)
            if r < hr or c > nc:
                continue
            if r > hr and isinstance(v, datetime.date) and cell.number_format.lower() != "yyyy-mm-dd":
                loc["datefmt"].append(where)
            if r > hr and isinstance(v, str) and NUM_TEXT.fullmatch(v) and v != "-" and (
                    c in has_num or cell.alignment.horizontal == "right"):
                loc["numtext"].append(where)
            w, h = widths.get(c, 8.43), ws.row_dimensions[r].height
            if cell.alignment.wrap_text:
                if h is not None and row_height(line_count(shown(cell), w)) > h + 0.5:
                    loc["tall"].append(where)
            elif dw(shown(cell)) > w - 1:
                loc["clip"].append(where)
    ids = [ws.cell(r, 1).value for r in range(hr + 1, last + 1)]
    head = [ws.cell(hr, c).value for c in range(1, nc + 1)]
    h1 = str(head[0] or "")
    if h1.upper().endswith("ID") or h1.endswith("번호"):
        dup = sorted(str(k) for k, n in collections.Counter(ids).items() if n > 1 and k not in (None, "-"))
        if dup:
            _add(E, name, "중복 ID", dup)
    dup = sorted(str(k) for k, n in collections.Counter(head).items() if n > 1)
    if dup:
        _add(E, name, "머리행 이름 중복", dup)
    for k, msg in (("empty", "빈 셀 (없으면 '-')"), ("font", "글꼴이 맑은 고딕이 아님"),
                   ("height", "행 높이 미지정"), ("numtext", "숫자가 글로 저장됨"),
                   ("datefmt", "날짜 서식이 YYYY-MM-DD 가 아님"), ("formula", "수식 셀 ('=' 로 시작하는 글)"),
                   ("huge", "행 높이가 Excel 상한 409.5pt 초과 (글을 줄이거나 열을 넓힘)")):
        if loc[k]:
            _add(E, name, msg, loc[k])
    if loc["long"]:
        _add(W, name, f"{LONG}자 초과 셀 (여러 값이면 하위 표 시트로 나눔)", loc["long"])
    if loc["clip"]:
        _add(W, name, "열 너비가 내용보다 좁아 잘림", loc["clip"])
    if loc["tall"]:
        _add(W, name, "행 높이가 줄 수보다 낮아 잘림", loc["tall"])


def check(wb):
    """(오류 목록, 경고 목록)"""
    E, W = [], []
    if wb.sheetnames[:2] != DOC_SHEETS:
        _add(E, "책", "첫 두 시트가 '문서 정보', '개정 이력' 이 아님")
    for ws in wb.worksheets:
        _sheet(ws, E, W)
    return E, W


def main(argv):
    if not argv:
        print(__doc__)
        return 2
    code = 0
    for path in argv:
        try:
            E, W = check(load_workbook(path))
        except Exception as e:
            print(f"{path}: 열 수 없음 ({e})")
            return 2
        print(f"{path}: 오류 {len(E)}, 경고 {len(W)}")
        for m in E:
            print("  오류", m)
        for m in W:
            print("  경고", m)
        code = 1 if E else code
    print("결과:", "실패" if code else "통과")
    return code


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main(sys.argv[1:]))
