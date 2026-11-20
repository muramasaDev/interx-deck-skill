"""xlsx 정의서 엔진: 목록형 정의서(테이블, API, 기능 항목, 화면)를 정석 서식으로 만든다
사용 예는 example_sheet.py. 저장 전에 sheet_check 규칙을 통과해야 한다 (오류는 ValueError)."""
import collections
import datetime
import math
import os
import re
import sys
from dataclasses import dataclass

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.properties import PageSetupProperties

sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))
from sheet_check import FONT, check, dw, line_count, row_height  # noqa: E402

HEAD, GRID = "D9D9D9", "A6A6A6"   # 공공기관 정의서식: 연한 회색 머리, 검정 굵은 글씨, 가는 회색 격자
_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
MIN_DATE = datetime.date(1900, 3, 1)   # Excel 날짜 일련번호가 안전한 범위


def _label(s, what):
    """이름, 머리, 단위, 그룹 라벨: 비거나 '=' 로 시작하거나 제어 문자가 있으면 거부"""
    t = str(s)
    if not t.strip() or t.lstrip().startswith("=") or _CTRL.search(t):
        raise ValueError(f"{what} {s!r}: 비었거나 '=' 로 시작하거나 제어 문자가 있음")
    return t
STATUS = {"완료": "E2EFDA", "진행": "FFF2CC", "계획": "EDEDED", "미정": "FCE4E4"}
ALIGN = {"id": "center", "ref": "center", "code": "center", "text": "left", "num": "right", "date": "center", "status": "center"}
_S = Side(style="thin", color=GRID)
BORDER = Border(left=_S, right=_S, top=_S, bottom=_S)


def _font(**kw):
    return Font(name=FONT, size=kw.pop("size", 10), **kw)


@dataclass
class Field:
    """열 하나. kind: id(유일), ref(하위 표의 상위 ID, 첫 두 열 조합이 유일), code, text, num(decimals, unit), date, status(mapping)"""
    name: str
    kind: str
    width: float = None
    decimals: int = 0
    unit: str = None
    mapping: dict = None
    indent: bool = False   # True 면 줄 맨 앞 공백을 지우지 않음 (코드, JSON 예시의 들여쓰기)

    def __post_init__(self):
        _label(self.name, "열 이름")
        if self.unit:
            _label(self.unit, "단위")
        if self.kind not in ALIGN:
            raise ValueError(f"{self.name}: 알 수 없는 종류 {self.kind!r} ({', '.join(ALIGN)})")

    @property
    def head(self):
        return f"{self.name} ({self.unit})" if self.unit else self.name

    def cell(self, v):
        """검증한 셀 값. 비어 있으면 '-'"""
        if v is None or (isinstance(v, str) and v.strip() in ("", "-")):
            return "-"
        if self.kind == "num":
            if isinstance(v, bool) or not isinstance(v, (int, float)) or abs(v) > 1e15 or not math.isfinite(v):
                raise ValueError(f"{self.name}: 숫자가 아님 {v!r}")
            return v
        if self.kind == "date":
            if isinstance(v, datetime.datetime):
                v = v.date()
            if isinstance(v, str):
                try:
                    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
                        raise ValueError
                    v = datetime.date.fromisoformat(v)
                except ValueError:
                    raise ValueError(f"{self.name}: 날짜는 YYYY-MM-DD 여야 함 {v!r}") from None
            if not isinstance(v, datetime.date):
                raise ValueError(f"{self.name}: 날짜가 아님 {v!r}")
            if not MIN_DATE <= v <= datetime.date(9999, 12, 31):
                raise ValueError(f"{self.name}: 날짜는 {MIN_DATE} 부터 {v!r}")
            return v
        v = str(v).rstrip() if self.indent else str(v).strip()
        if _CTRL.search(v):
            raise ValueError(f"{self.name}: 제어 문자가 있음 {v!r}")
        if self.kind == "status" and v not in (self.mapping or STATUS):
            raise ValueError(f"{self.name}: 상태 값 {v!r} 이 매핑에 없음 ({', '.join(self.mapping or STATUS)})")
        if v.startswith("="):
            raise ValueError(f"{self.name}: '=' 로 시작하는 글은 수식이 되므로 쓸 수 없음 {v!r}")
        return v

    def text(self, v):
        """셀에 보이는 글 (폭 계산용)"""
        if isinstance(v, datetime.date):
            return "0000-00-00"
        return f"{v:,.{self.decimals}f}" if isinstance(v, (int, float)) else str(v)


class Book:
    def __init__(self, title, info=None, history=None):
        if not title:
            raise ValueError("문서명이 필요함")
        if not history:
            raise ValueError("개정 이력이 필요함: [(판, 일자, 구분, 내용), ...]")
        self.title, self.wb = title, Workbook()
        self.wb.remove(self.wb.active)
        self.wb.properties.title, self.wb.properties.creator = title, ""
        info = {"문서명": title, **(info or {})}
        if "기준일" in info:
            d = Field("기준일", "date").cell(info["기준일"])
            info["기준일"] = d.isoformat() if hasattr(d, "isoformat") else d
        self.table("문서 정보", [Field("항목", "text", 18), Field("내용", "text", 70)],
                   [(k, v) for k, v in info.items()])
        for r in range(3, self.wb["문서 정보"].max_row + 1):   # 문서 정보의 항목 열은 연한 회색 칸에 굵게 가운데: 값 칸과 구분
            c = self.wb["문서 정보"].cell(r, 1)
            c.font, c.fill = _font(bold=True), PatternFill("solid", fgColor="F2F2F2")
            c.alignment = Alignment(horizontal="center", vertical="center")
        self.table("개정 이력", [Field("판", "code"), Field("일자", "date"), Field("구분", "code"), Field("내용", "text", 60)],
                   history)   # 작성자 열은 두지 않음 (개정 이력 한 줄: 판, 일자, 구분, 내용)

    def table(self, name, fields, rows, unit=None, group=None):
        _label(name, "표 이름")
        for label in list(group or {}) + ([unit] if unit else []):
            _label(label, "그룹 라벨 또는 단위")
        names = [f.name for f in fields]
        if len(set(names)) != len(names):
            raise ValueError(f"{name}: 열 이름 중복")
        vals = []
        for i, row in enumerate(rows, 1):
            if isinstance(row, dict):
                if set(row) != set(names):
                    raise ValueError(f"{name}: {i}번째 행의 키가 열 이름과 다름 (빠진 {sorted(set(names) - set(row))}, "
                                     f"여분 {sorted(set(row) - set(names))})")
                row = [row[n] for n in names]
            else:
                row = list(row)
            if len(row) != len(fields):
                raise ValueError(f"{name}: {i}번째 행의 값이 {len(row)}개 (열 {len(fields)}개)")
            vals.append([f.cell(v) for f, v in zip(fields, row)])
        key = (lambda r: r[0]) if fields[0].kind == "id" else (
            (lambda r: tuple(r[:2])) if fields[0].kind == "ref" and len(fields) > 1 else None)   # ref: 상위 ID 를 되풀이하는 하위 표
        if key:
            dup = [k for k, c in collections.Counter(key(r) for r in vals).items() if c > 1]
            if dup:
                raise ValueError(f"{name}: 중복 ID {dup}")
        spans = {}
        for label, (a, b) in (group or {}).items():
            i, j = names.index(a), names.index(b)
            if i > j or any(set(range(i, j + 1)) & set(range(p, q + 1)) for p, q in spans.values()):
                raise ValueError(f"{name}: 그룹 {label!r} 범위가 거꾸로이거나 겹침")
            spans[label] = (i, j)
        if unit and len(fields) < 2:
            raise ValueError(f"{name}: 단위는 열이 둘 이상일 때만 둠")
        self._write(name, fields, vals, unit, spans)

    def matrix(self, name, row_ids, col_ids, links, corner="ID"):
        """행 ID 와 열 ID 의 연결을 ● 로 표시. links: (행 ID, 열 ID) 쌍"""
        links = set(map(tuple, links))
        bad = [p for p in links if p[0] not in row_ids or p[1] not in col_ids]
        if bad:
            raise ValueError(f"{name}: 없는 ID 를 잇는 연결 {sorted(bad)[:3]}")
        fields = [Field(corner, "id")] + [Field(c, "code") for c in col_ids]
        self.table(name, fields, [[r] + ["●" if (r, c) in links else "-" for c in col_ids] for r in row_ids])

    def _write(self, name, fields, vals, unit, spans):
        title = re.sub(r"[\\/?*\[\]:]", "_", name)[:31]
        if title.strip("'") != title or title.lower() == "history" or title.lower() in map(str.lower, self.wb.sheetnames):
            raise ValueError(f"시트 이름 '{title}' 이 겹치거나 Excel 이 거부하는 꼴 (대소문자 무시, 앞뒤 ', History)")
        ws = self.wb.create_sheet(title)
        nc, hr = len(fields), 3 if spans else 2
        widths = [max(f.width or min(60, max(8, 2 + max([dw(f.head)] + [dw(f.text(r[j])) for r in vals]))), dw(f.head) + 5)
                  for j, f in enumerate(fields)]   # 머리글은 필터 단추(약 3칸) 때문에 글자 수 + 5 이상: 머리글이 두 줄로 꺾이지 않게
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[L(j)].width = w
        ws.cell(1, 1, name).font = _font(size=12, bold=True)
        ws.cell(1, 1).alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 24
        if unit:
            ws.cell(1, nc, f"(단위: {unit})").font = _font()
            ws.cell(1, nc).alignment = Alignment(horizontal="right", vertical="center")
        mid = Alignment(horizontal="center", vertical="center", wrap_text=True)

        def style_head(c):
            c.font, c.border, c.alignment = _font(bold=True), BORDER, mid
            c.fill = PatternFill("solid", fgColor=HEAD)

        if spans:
            for j in range(1, nc + 1):
                style_head(ws.cell(hr - 1, j))
            for label, (i, j) in spans.items():
                ws.cell(hr - 1, i + 1, label)
                ws.merge_cells(start_row=hr - 1, start_column=i + 1, end_row=hr - 1, end_column=j + 1)
            ws.row_dimensions[hr - 1].height = 24
        for j, f in enumerate(fields, 1):
            style_head(ws.cell(hr, j, f.head))
        ws.row_dimensions[hr].height = max(24, row_height(max(line_count(f.head, w) for f, w in zip(fields, widths))))
        for r, row in enumerate(vals, hr + 1):
            lines = 1
            for j, (f, v) in enumerate(zip(fields, row), 1):
                c = ws.cell(r, j, v)
                c.font, c.border = _font(), BORDER
                c.alignment = Alignment(horizontal=ALIGN[f.kind], vertical="center", wrap_text=f.kind == "text")   # 세로 가운데: 한 줄 칸과 여러 줄 칸이 같은 선에 놓임
                if isinstance(v, datetime.date):
                    c.number_format = "yyyy-mm-dd"
                elif isinstance(v, (int, float)):
                    c.number_format = "#,##0" + ("." + "0" * f.decimals if f.decimals else "")
                if f.kind == "status" and v != "-":
                    c.fill = PatternFill("solid", fgColor=(f.mapping or STATUS)[v])
                if f.kind == "text":
                    lines = max(lines, line_count(v, widths[j - 1]))
            h = row_height(lines)
            if h > 409:
                raise ValueError(f"{name}: {r - hr}번째 행 높이 {h}pt 가 Excel 상한 409pt 를 넘음: 글을 줄이거나 열을 넓힘")
            ws.row_dimensions[r].height = h
        ws.auto_filter.ref = f"A{hr}:{L(nc)}{hr + len(vals)}"
        ws.freeze_panes = f"A{hr + 1}"
        ws.page_setup.orientation, ws.page_setup.paperSize = "landscape", ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.print_title_rows = f"1:{hr}"
        for part, text in ((ws.oddFooter.left, self.title.replace("&", "&&")),
                           (ws.oddFooter.center, "&P / &N"), (ws.oddFooter.right, "&A")):
            part.text, part.size = text, 9

    def save(self, path):
        errs, warns = check(self.wb)
        for w in warns:
            print("경고", w)
        if errs:
            raise ValueError("sheet_check 오류:\n" + "\n".join(errs))
        self.wb.save(path)
