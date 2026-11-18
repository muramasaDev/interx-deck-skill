"""xlsx 정의서 엔진 자체 시험: 규칙 위반이 각각 막히는지, 정상 예제가 통과하는지
사용: python -P scripts/selfcheck_sheet.py   (임시 폴더만 씀)"""
import os
import subprocess
import sys
import tempfile

from openpyxl import load_workbook
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.realpath(__file__))
sys.path.insert(0, HERE)
from example_sheet import build_example  # noqa: E402
from ixsheet import Book, Field  # noqa: E402
from sheet_check import check, line_count, row_height  # noqa: E402

tmp = tempfile.mkdtemp(prefix="selfcheck_sheet_")
good = os.path.join(tmp, "good.xlsx")
build_example(good)
E, W = check(load_workbook(good))
assert not E and not W, (E, W)
assert subprocess.run([sys.executable, "-P", os.path.join(HERE, "sheet_check.py"), good],
                      capture_output=True).returncode == 0

# 정상 책의 구조
wb = load_workbook(good)
assert wb.sheetnames[:2] == ["문서 정보", "개정 이력"], wb.sheetnames
ws = wb["API 목록"]
assert ws.freeze_panes == "A4" and ws.auto_filter.ref == "A3:H9" and ws.print_title_rows.replace("$", "") == "1:3"
assert ws.page_setup.orientation == "landscape" and ws.page_setup.fitToWidth == 1
assert ws["A4"].alignment.horizontal == "center" and ws["D4"].alignment.horizontal == "right"
assert ws["D4"].number_format == "#,##0" and ws["E4"].number_format == "#,##0.0"
assert ws["B4"].alignment.horizontal == "left" and ws["B4"].alignment.wrap_text
assert ws["G4"].number_format == "yyyy-mm-dd" and ws["H4"].fill.fgColor.rgb.endswith("FFF2CC")
assert ws["A3"].fill.fgColor.rgb.endswith("002060") and ws["A3"].font.name == "맑은 고딕"
assert ws["H1"].value is None and wb["처리 현황"]["D1"].value == "(단위: 건)"
assert ws.row_dimensions[6].height == row_height(line_count(ws["F6"].value, 50)) > 18
assert wb["요구사항 추적"]["B3"].value == "●" and wb["요구사항 추적"]["C3"].value == "-"

# 파일 규칙 위반: 정상 책을 고쳐 검사가 잡는지 (키워드, 오류 또는 경고)
def row_of(ws):
    return int(ws.auto_filter.ref.split(":")[0][1:]) + 1


def merge(ws):
    ws.merge_cells(f"A{row_of(ws) + 1}:B{row_of(ws) + 1}")


def blank(ws):
    ws[f"B{row_of(ws)}"].value = None


def dup_id(ws):
    ws[f"A{row_of(ws) + 1}"].value = ws[f"A{row_of(ws)}"].value


def dup_head(ws):
    ws[f"C{row_of(ws) - 1}"].value = ws[f"B{row_of(ws) - 1}"].value


def num_text(ws):
    ws[f"D{row_of(ws)}"].value = "1,234"


def date_fmt(ws):
    ws[f"G{row_of(ws)}"].number_format = "General"


def font(ws):
    ws[f"B{row_of(ws)}"].font = Font(name="Arial")


def height(ws):
    ws.row_dimensions[row_of(ws)].height = None


def freeze(ws):
    ws.freeze_panes = None


def no_filter(ws):
    ws.auto_filter.ref = None


def long_cell(ws):
    ws[f"F{row_of(ws)}"].value = "가" * 301


def narrow(ws):
    ws.column_dimensions["G"].width = 8


cases = [("병합", merge, True), ("빈 셀", blank, True), ("중복 ID", dup_id, True), ("머리행 이름 중복", dup_head, True),
         ("숫자가 글로", num_text, True), ("날짜 서식", date_fmt, True), ("글꼴", font, True),
         ("행 높이 미지정", height, True), ("틀 고정", freeze, True), ("필터 없음", no_filter, True),
         ("300자 초과", long_cell, False), ("열 너비", narrow, False)]
for key, fn, is_error in cases:
    wb = load_workbook(good)
    fn(wb["API 목록"])
    E, W = check(wb)
    hits = [m for m in (E if is_error else W) if key in m]
    assert hits, (key, E, W)
    assert not (W if is_error else E) or key == "300자 초과", (key, E, W)
wb = load_workbook(good)
wb.move_sheet("개정 이력", offset=2)
assert any("첫 두 시트" in m for m in check(wb)[0])

# 만들 때 막힘 (ValueError)
b = Book("시험", history=[("v0.1.0", "2026-11-17", "신규", "시험")])
fields = [Field("항목 ID", "id"), Field("값", "num"), Field("기준일", "date"), Field("상태", "status")]
ok = ("A-1", 1, "2026-11-17", "완료")
b.table("정상", fields, [ok])
for why, f, rows in [("중복 ID", fields, [ok, ok]), ("날짜 오류", fields, [("A-2", 1, "2026-13-01", "완료")]),
                     ("날짜 형식", fields, [("A-2", 1, "2026/11/17", "완료")]),
                     ("숫자 아님", fields, [("A-2", "1,000", "2026-11-17", "완료")]),
                     ("상태 값", fields, [("A-2", 1, "2026-11-17", "몰라")]),
                     ("수식 글", [Field("항목 ID", "id")], [("=1+1",)]),
                     ("열 수", fields, [("A-2", 1)]),
                     ("열 이름 중복", [Field("값", "num")] * 2, [(1, 2)])]:
    try:
        b.table(why, f, rows)
    except ValueError:
        continue
    raise AssertionError(f"{why} 가 막히지 않음")
for why, args in [("행렬 없는 ID", ("행렬", ["A-1"], ["B-1"], [("A-1", "B-9")])),
                  ("행렬 중복 행", ("행렬", ["A-1", "A-1"], ["B-1"], []))]:
    try:
        b.matrix(*args)
    except ValueError:
        continue
    raise AssertionError(f"{why} 가 막히지 않음")
try:
    b.table("정상", fields, [ok])
    raise AssertionError("시트 이름 중복이 막히지 않음")
except ValueError:
    pass

# 저장 관문: 규칙을 어긴 책은 저장되지 않음
b.wb["정상"].merge_cells("A3:B3")
try:
    b.save(os.path.join(tmp, "bad.xlsx"))
    raise AssertionError("병합된 책이 저장됨")
except ValueError:
    assert not os.path.exists(os.path.join(tmp, "bad.xlsx"))

# CLI: 오류 파일은 종료 코드 1
wb = load_workbook(good)
merge(wb["API 목록"])
bad = os.path.join(tmp, "bad_file.xlsx")
wb.save(bad)
assert subprocess.run([sys.executable, "-P", os.path.join(HERE, "sheet_check.py"), bad],
                      capture_output=True).returncode == 1
# 검토 반영: 입력 검증, 큰 표 속도
import datetime  # noqa: E402
import time  # noqa: E402


def _bad(fn, why):
    try:
        fn()
    except ValueError:
        return
    raise AssertionError(why)


_h = [("v0.1.0", "2026-11-18", "신규", "시험")]
_F = [Field("ID", "id"), Field("글", "text")]
_bad(lambda: Book("=1+1", history=_h), "문서명 수식")
_bad(lambda: Book("시험", history=_h).table("=1+1", _F, [("A", "b")]), "표 이름 수식")
_bad(lambda: Field("=열", "text"), "열 이름 수식")
_bad(lambda: Book("시험", history=_h).table("표", _F, [{"ID": "A", "글": "b", "여분": 1}]), "여분 키")
_bad(lambda: Book("시험", history=_h).table("표", _F, [{"ID": "A"}]), "빠진 키")
_bad(lambda: Book("시험", history=_h).table("", _F, [("A", "b")]), "빈 표 이름")
_bad(lambda: Book("시험", history=_h).table("History", _F, [("A", "b")]), "예약 시트 이름")
_bk = Book("시험", history=_h)
_bk.table("Test", _F, [("A", "b")])
_bad(lambda: _bk.table("test", _F, [("A", "b")]), "대소문자만 다른 시트 이름")
_bad(lambda: Book("시험", history=_h).table("표", [Field("ID", "id"), Field("날짜", "date")], [("A", "1899-12-31")]), "날짜 범위")
_bad(lambda: Book("시험", history=_h).table("표", _F, [("A", "b")]), "제어 문자")
_bad(lambda: Book("시험", history=_h).table("표", [Field("ID", "id"), Field("수", "num")], [("A", 10 ** 400)]), "큰 수")
_bad(lambda: Book("시험", history=_h).table("표", [Field("ID", "id"), Field("글", "text", 8)], [("A", "가" * 299)]), "행 높이 상한")
_big = Book("시험", history=_h)
_t0 = time.time()
_big.table("큰 표", [Field("ID", "id"), Field("글", "text", 30), Field("수", "num")], [(f"R-{i}", "샘플", i) for i in range(3000)])
assert time.time() - _t0 < 20, "3000행 표가 느림"
print("selfcheck_sheet OK")
