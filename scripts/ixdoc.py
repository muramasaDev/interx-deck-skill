"""설계 문서 부품: 기능 상세, 화면 상세 장을 표 중심으로 (ixdeck.py 맨 아래에서 불러 Col 과 Deck 에 붙임)
- 표 부품: field_table(입출력), limit_table(한계 값), error_table(오류 코드), element_table(화면 요소), ref_table(관련 번호)
- 흐름: flow (번호 원과 이은 선)
- 장: Deck.func_page(기능 상세), Deck.screen_page(화면 상세). 내용이 있는 구역만 패널이 되고 너비는 내용 양으로 정함"""
import json
import re

from ixdeck import BLACK, CAP_H, GRAY_L, NUM_D, Col, Deck, _lh, _nl, box, num_badge, report_panels, table, table_caption, text

TS = 8.0   # 설계 문서 표 글자 (보고 표 7.26 보다 큼)
_PLAIN = re.compile(r"[가-힣]니다|(한다|된다|이다|했다|하였다)(\.|$)")   # 서술체 검사에 걸리는 화면 문구만 예외로 등록
_CODE = re.compile(r"^(\d{3}|[A-Z][A-Z0-9_-]{2,})$")


def _table(self, head, rows, ratio, title, unit, align, hl=(), row_h=0.27):
    """머리행 고정 표. 행마다 줄 수에 맞춰 높이를 따로 정함 (PowerPoint 표는 칸이 줄 바꿈되면 자동으로 늘어 장 밖으로 넘기 때문)"""
    rows = [list(r) for r in rows]
    if not rows:
        raise ValueError(f"{head[0]} 표 행이 없음")
    for r in rows:
        if len(r) != len(head) or any(c is None or not str(c).strip() for c in r):
            raise ValueError(f"표 행은 {head} 칸을 모두 채움 (없으면 -): {r}")
    cw = [self.w * k / sum(ratio) for k in ratio]
    wide = lambda j: 1.12 if j in hl else 1.0   # 남색 굵은 글씨 열은 글자가 더 넓음
    hs = [max(row_h, max(_nl(str(c), (cw[j] - 0.12) / wide(j), TS) for j, c in enumerate(r)) * _lh(TS) + 0.07) for r in rows]
    head_h = min(row_h, 0.27)

    def draw(y, r):
        table_caption(self.s, self.x, y, self.w, title, unit, [head, *rows], key=self._ck(y))
        return table(self.s, self.x, y + CAP_H, cw, [head, *rows], row_h=row_h, row_hs=hs, size=TS, align=align,
                     cell_hl=[(i, j) for i in range(len(rows)) for j in hl])
    return self._add(CAP_H + head_h + sum(hs), draw)


def field_table(self, rows, title="입출력 항목", unit=None, row_h=0.27):
    """rows=[(항목, 입력 또는 출력, 형식, 필수 또는 선택 또는 -, 설명)]"""
    for r in rows:
        if len(r) != 5 or r[1] not in ("입력", "출력") or r[3] not in ("필수", "선택", "-"):
            raise ValueError(f"field_table 행은 (항목, 입력 또는 출력, 형식, 필수 또는 선택 또는 -, 설명): {r}")
    return _table(self, ["항목", "방향", "형식", "필수", "설명"], rows, [1.5, 0.55, 0.8, 0.55, 2.4], title, unit, ["l", "c", "c", "c", "l"], row_h=row_h)


def limit_table(self, rows, title="한계 값", unit=None, row_h=0.27):
    """rows=[(항목, 값, 단위, 비고)]. 값 열은 남색 굵은 글씨 (한계값이 일반 글에 묻히지 않게)"""
    return _table(self, ["항목", "값", "단위", "비고"], rows, [1.7, 1.35, 0.65, 1.8], title, unit, ["l", "r", "c", "l"], hl=(1,), row_h=row_h)


def error_table(self, rows, title="오류 코드", unit=None, row_h=0.27):
    """rows=[(코드, 조건, 화면 문구 및 처리)]. 코드는 숫자 3자리 또는 대문자 식별자"""
    for r in rows:
        if len(r) != 3 or not _CODE.match(str(r[0])):
            raise ValueError(f"error_table 행은 (코드, 조건, 화면 문구 및 처리), 코드는 400 또는 E_AUTH 꼴: {r}")
    return _table(self, ["코드", "조건", "화면 문구 및 처리"], rows, [0.6, 1.8, 2.2], title, unit, ["c", "l", "l"], hl=(0,), row_h=row_h)


def element_table(self, rows, title="화면 요소", unit=None, row_h=0.27):
    """rows=[(번호, 요소, 동작, 검증 및 표시)]. 번호는 캡처의 번호 원과 같은 수"""
    for r in rows:
        if len(r) != 4 or not str(r[0]).isdigit():
            raise ValueError(f"element_table 행은 (번호, 요소, 동작, 검증 및 표시), 번호는 숫자: {r}")
    return _table(self, ["번호", "요소", "동작", "검증 및 표시"], rows, [0.55, 1.1, 1.8, 1.7], title, unit, ["c", "l", "l", "l"], hl=(0,), row_h=row_h)


def ref_table(self, refs, title="관련 번호", unit=None, row_h=0.27):
    """refs={구분: 번호 글}. 예: {"기능": "F-001 ~ 003", "화면": "SCR-02"}"""
    if not refs:
        raise ValueError("ref_table 이 비었음")
    return _table(self, ["구분", "번호"], list(refs.items()), [0.8, 3.0], title, unit, ["l", "l"], row_h=row_h)


def flow(self, steps, row_h=0.27):
    """처리 흐름: 번호 원을 가는 선으로 이음. steps=[글]. 단계마다 글 길이에 맞춰 높이가 달라짐"""
    if len(steps) < 2:
        raise ValueError("flow 는 두 단계 이상 (한 단계면 문장으로)")
    gap, w = max(row_h - 0.17, 0.1), self.w - NUM_D - 0.12
    hs = [max(NUM_D, _nl(t, w, TS) * _lh(TS), row_h - 0.06) + 0.04 for t in steps]

    def draw(y, r):
        yy, tops = y, []
        for t, h in zip(steps, hs):
            tops.append(yy + (h - NUM_D) / 2)
            text(self.s, self.x + NUM_D + 0.12, yy, w, h, t, size=TS, color=BLACK, align="l")
            yy += h + gap
        for a, b in zip(tops, tops[1:]):   # 원 사이 선 (원 뒤에 그려지도록 먼저)
            box(self.s, self.x + NUM_D / 2 - 0.004, a + NUM_D, 0.008, b - a - NUM_D, fill=GRAY_L)
        for n, top in enumerate(tops, 1):
            num_badge(self.s, self.x, top, n).name = "IXFLOW"
    return self._add(sum(hs) + gap * (len(hs) - 1), draw)


for _f in (field_table, limit_table, error_table, element_table, ref_table, flow):
    setattr(Col, _f.__name__, _f)


def _comfort(n):
    """행 수가 적은 장은 행을 넉넉하게 (장 아래가 비지 않게), 많으면 촘촘하게"""
    return 0.27 if n >= 22 else 0.34 if n >= 14 else 0.42


def func_page(self, subtitle, lead, *, flow, limits=None, io=None, errors=None, refs=None, notes=None):
    """기능 상세 장. 있는 구역만 패널이 된다 (흐름은 항상).
    flow=[글], limits=[(항목, 값, 단위, 비고)], io=[(항목, 방향, 형식, 필수, 설명)],
    errors=[(코드, 조건, 화면 문구 및 처리)], refs={구분: 번호}"""
    s = self.page(subtitle, lead=lead, notes=notes)
    self.allow_text(*(e[2] for e in errors or () if len(e) == 3 and isinstance(e[2], str)))
    rh = _comfort(len(flow) + len(limits or ()) + len(io or ()) + len(errors or ()) + len(refs or ()))
    titles, builders, wt = ["처리 흐름"], [], [1.0]

    def left(c):
        c.flow(flow, row_h=rh)
        if refs:
            c.ref_table(refs, row_h=rh)
    builders.append(left)
    if limits:
        titles.append("한계 및 규칙")
        builders.append(lambda c: c.limit_table(limits, row_h=rh))
        wt.append(1.2 + 0.04 * len(limits))
    if io or errors:
        titles.append("입출력 및 오류" if io and errors else ("입출력" if io else "오류"))

        def right(c):
            if io:
                c.field_table(io, row_h=rh)
            if errors:
                c.error_table(errors, row_h=rh)
        builders.append(right)
        wt.append(1.5 + 0.08 * (len(io or ()) + len(errors or ())))
    cols = report_panels(s, titles, widths=wt if len(wt) > 1 else None, fit=True)
    for c, b in zip(cols, builders):
        b(c)
    for c in cols:
        c.done()
    return s


def screen_page(self, subtitle, lead, *, image, marks, elements, rules=None, refs=None, caption=None, notes=None):
    """화면 상세 장: 왼쪽에 캡처를 크게, 오른쪽에 번호별 요소 표.
    image=캡처 경로, marks=[(x0, y0, x1, y1, 번호)] (그림 비율 0 ~ 1), elements=[(번호, 요소, 동작, 검증 및 표시)],
    rules=[(항목, 값, 단위, 비고)] 화면 한계 값, refs={구분: 번호}"""
    s = self.page(subtitle, lead=lead, notes=notes)
    seen = set()
    for m in marks:
        if len(m) != 5 or not (0 <= m[0] < m[2] <= 1 and 0 <= m[1] < m[3] <= 1) or str(m[4]) in seen:
            raise ValueError(f"marks 는 (x0, y0, x1, y1, 번호): 0 <= x0 < x1 <= 1, 0 <= y0 < y1 <= 1, 번호는 서로 다름: {m}")
        seen.add(str(m[4]))
    if len({str(e[0]) for e in elements}) != len(elements):
        raise ValueError("elements 의 번호가 겹침")
    nums = {str(m[4]) for m in marks}
    if nums != {str(e[0]) for e in elements}:
        raise ValueError(f"캡처 번호 {sorted(nums)} 와 요소 표 번호 {sorted(str(e[0]) for e in elements)} 가 다름")
    left, right = report_panels(s, ["화면", "요소 설명"], widths=(1.7, 1.0), fit=True)
    left.image(image, left.bottom - left.top - CAP_H - 0.14, marks=marks, title=caption or subtitle)   # 0.14: 번호 원 여유
    right.element_table(elements)
    if rules:
        right.limit_table(rules)
    if refs:
        right.ref_table(refs)
    left.done()
    right.done()
    return s


def allow_text(self, *texts):
    """화면 문구를 그대로 옮긴 글을 문장 검사에서 뺌 (서술체 등). 오류 표의 화면 문구는 자동"""
    self.allow += [re.escape(t) for t in texts if _PLAIN.search(t) and re.escape(t) not in self.allow]
    self.prs.core_properties.keywords = "ix-allow:" + json.dumps(self.allow, ensure_ascii=False)


Deck.func_page, Deck.screen_page, Deck.allow_text = func_page, screen_page, allow_text
