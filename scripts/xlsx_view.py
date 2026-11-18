"""xlsx 그림 확인: 시트마다 앞 N 행을 PNG 로 만든다
사용: python -P scripts/xlsx_view.py 파일.xlsx 출력폴더 [--rows 40]
openpyxl 로 읽은 서식을 HTML 표로 그려 Playwright(msedge, 없으면 chrome)로 찍는다. Excel 은 열지 않는다.
실제 Excel 화면과 똑같지는 않으며 줄 바꿈과 잘림은 근사이다. 불가능하면 이유를 출력하고 종료 코드 2."""
import html
import os
import re
import sys

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))
from sheet_check import col_widths, dw, shown  # noqa: E402

GRID = "E1E1E1"
BW = {"thin": 1, "hair": 1, "medium": 2, "thick": 3, "double": 3}


def rgb(color, default=None):
    v = getattr(color, "rgb", None)
    return "#" + v[-6:] if isinstance(v, str) and len(v) in (6, 8) else default


def px(width):
    return round(width * 7 + 5)


def td(ws, r, c, span, rspan, hpx):
    cell = ws.cell(r, c)
    f, a, b = cell.font, cell.alignment, cell.border
    v = shown(cell)
    h = a.horizontal if a.horizontal not in (None, "general") else ("right" if isinstance(cell.value, (int, float)) else "left")
    css = [f"font-family:'{html.escape(f.name or '맑은 고딕')}',sans-serif", f"font-size:{(f.size or 10) * 1.333:.1f}px",
           f"color:{rgb(f.color, '#000')}", f"font-weight:{'bold' if f.b else 'normal'}"]
    if cell.fill.fill_type == "solid":
        css.append(f"background:{rgb(cell.fill.fgColor, '#fff')}")
    for side in ("left", "right", "top", "bottom"):
        s = getattr(b, side)
        css.append(f"border-{side}:{BW.get(s.style, 1) if s.style else 1}px solid {rgb(s.color, '#000') if s.style else '#' + GRID}")
    box = {"top": "flex-start", "center": "center", "bottom": "flex-end"}.get(a.vertical or "bottom")
    inner = (f"display:flex;align-items:{box};height:{hpx}px;overflow:hidden;padding:0 3px;box-sizing:border-box;")
    text = (f"width:100%;text-align:{h};line-height:20px;"
            + ("white-space:pre-wrap;word-break:break-all;" if a.wrap_text else "white-space:nowrap;overflow:hidden;"))
    return (f'<td colspan="{span}" rowspan="{rspan}" style="padding:0;{";".join(css)}">'
            f'<div style="{inner}"><div style="{text}">{html.escape(v)}</div></div></td>')


def sheet_html(ws, nrows):
    nc, widths = ws.max_column, col_widths(ws)
    nrows = min(nrows, ws.max_row)
    wpx = {c: px(widths.get(c, 8.43)) for c in range(1, nc + 1)}
    hpx = {r: round((ws.row_dimensions[r].height or 15) * 1.333) for r in range(1, nrows + 1)}
    start, covered = {}, set()
    for m in ws.merged_cells.ranges:
        start[(m.min_row, m.min_col)] = (m.max_row - m.min_row + 1, m.max_col - m.min_col + 1)
        covered |= {(r, c) for r in range(m.min_row, m.max_row + 1) for c in range(m.min_col, m.max_col + 1)}
        covered.discard((m.min_row, m.min_col))
    blocked = covered | set(start)
    rows = []
    for r in range(1, nrows + 1):
        plan, absorbed = {}, set()
        for c in range(1, nc + 1):
            cell, v = ws.cell(r, c), ws.cell(r, c).value
            if isinstance(v, str) and not cell.alignment.wrap_text and (r, c) not in start and c not in absorbed:
                need, have, right = dw(v) * 6.7 + 10, wpx[c], cell.alignment.horizontal == "right"
                k, step = c + (-1 if right else 1), -1 if right else 1
                while need > have and 1 <= k <= nc and ws.cell(r, k).value is None and (r, k) not in blocked and k not in absorbed:
                    have += wpx[k]
                    absorbed.add(k)
                    k += step
                if have > wpx[c]:
                    plan[c] = abs(k - step - c) + 1
        cells = []
        for c in range(1, nc + 1):
            if (r, c) in covered or c in absorbed:
                continue
            rs, cs = start.get((r, c), (1, 1))
            span = plan.get(c, cs)
            if r + rs - 1 > nrows:
                rs = nrows - r + 1
            h = sum(hpx[x] for x in range(r, r + rs))
            cells.append(td(ws, r, c, span, rs, h))
        rows.append(f'<tr style="height:{hpx[r]}px">{"".join(cells)}</tr>')
    cols = "".join(f'<col style="width:{wpx[c]}px">' for c in range(1, nc + 1))
    total = sum(wpx.values())
    cap = f"{ws.title}  (앞 {nrows}행 / 전체 {ws.max_row}행)"
    page = (f'<html><body style="margin:12px;background:#fff;font-family:\'맑은 고딕\',sans-serif">'
            f'<div style="font-size:12px;color:#666;margin-bottom:6px">{html.escape(cap)}</div>'
            f'<table style="border-collapse:collapse;table-layout:fixed;width:{total}px">'
            f'<colgroup>{cols}</colgroup>{"".join(rows)}</table></body></html>')
    return page, total


def main(argv):
    rows = 40
    if "--rows" in argv:
        i = argv.index("--rows")
        rows = int(argv[i + 1])
        del argv[i:i + 2]
    if len(argv) != 2:
        print(__doc__)
        return 2
    src, out = argv
    try:
        from playwright.sync_api import sync_playwright
        wb = load_workbook(src)
    except Exception as e:
        print(f"불가능: {e} (playwright 필요: pip install playwright)")
        return 2
    os.makedirs(out, exist_ok=True)
    with sync_playwright() as p:
        browser = None
        for ch in ("msedge", "chrome"):
            try:
                browser = p.chromium.launch(channel=ch)
                break
            except Exception as e:
                err = e
        if browser is None:
            print(f"불가능: msedge 와 chrome 을 열지 못함 ({str(err).splitlines()[0]})")
            return 2
        for i, ws in enumerate(wb.worksheets, 1):
            page_html, total = sheet_html(ws, rows)
            page = browser.new_page(viewport={"width": min(total + 30, 4000), "height": 400})
            page.set_content(page_html)
            path = os.path.join(out, f"{i:02d}_{re.sub(r'[^0-9A-Za-z가-힣-]+', '_', ws.title)}.png")
            page.screenshot(path=path, full_page=True)
            page.close()
            print(path)
        browser.close()
    print("참고: 이 그림은 openpyxl 로 읽은 서식을 HTML 로 그린 근사이며 실제 Excel 화면과 똑같지는 않음")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main(sys.argv[1:]))
