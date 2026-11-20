"""덱의 숫자와 그림이 원본 자료에 있는지 대조한다 (원본 자료를 발표 덱으로 옮길 때).
사용: python -P source_trace.py 덱.pptx 원본1 [원본2 ...] [--out 표.md]
원본: pdf, docx, pptx, xlsx, txt, md, csv.
1. 덱의 숫자(소수, 단위 붙은 수, 두 자리 이상 정수)가 원본 글에 같은 꼴로 있는지. '없음' 은 계산한 값이거나 옮기다 틀린 값이므로 사람이 확인
2. 원본의 Figure, Table, Algorithm, Equation, 그림, 표, 수식 번호가 덱(장 글, 노트)에 나오는지. '없음' 은 덱에서 빠졌거나 번호를 안 적은 것
관문이 아니라 참고 도구: 결과를 보고 원본 대조 검토(사실 정확성 1건)에 넘긴다. 종료 코드는 항상 0."""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.realpath(__file__))
sys.path.insert(0, HERE)
sys.stdout.reconfigure(errors="replace")

CAP = re.compile(r"\b(Figure|Fig\.?|Table|Algorithm|Equation|Eq\.?)\s*(\d+(?:\.\d+)?)|(그림|표|수식|알고리즘)\s*(\d+(?:[.-]\d+)?)", re.I)
KIND = {"fig": "Figure", "figure": "Figure", "table": "Table", "algorithm": "Algorithm", "equation": "Equation", "eq": "Equation",
        "그림": "Figure", "표": "Table", "수식": "Equation", "알고리즘": "Algorithm"}
NUM = re.compile(r"(?<![\w.])(\d[\d,]*(?:\.\d+)?)\s?(%|배|건|장|명|개|ms|초|분|시간|mm|µm|um|nm|MB|GB|KB|배|점|회|일|개월|년|원|만|억)?")
DATE = re.compile(r"\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}|\d{4}\.\d{1,2}|\d{1,2}[.\-]\d{1,2}[.\-]\d{1,2}")


def read_source(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        try:
            import fitz
            with fitz.open(path) as d:
                return "\n".join(p.get_text() for p in d)
        except ImportError:
            from pypdf import PdfReader
            return "\n".join((p.extract_text() or "") for p in PdfReader(path).pages)
    if ext == ".docx":
        import docx
        d = docx.Document(path)
        return "\n".join([p.text for p in d.paragraphs] + [c.text for t in d.tables for r in t.rows for c in r.cells])
    if ext == ".pptx":
        from pptx import Presentation
        from check_ban import notes_of, slide_texts
        return "\n".join(t for s in Presentation(path).slides for t in list(slide_texts(s)) + [notes_of(s)])
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        return "\n".join(str(c) for ws in openpyxl.load_workbook(path, data_only=True) for r in ws.iter_rows(values_only=True) for c in r if c is not None)
    return open(path, encoding="utf8", errors="replace").read()


def deck_slides(path):
    from pptx import Presentation
    from check_ban import notes_of, slide_texts
    return [(i, "\n".join(slide_texts(s)), notes_of(s)) for i, s in enumerate(Presentation(path).slides, 1)]


def norm(s):
    return re.sub(r"[\s,]", "", s)


def labels(text):
    out = set()
    for m in CAP.finditer(text):
        k = (m.group(1) or m.group(3)).lower().rstrip(".")
        n = m.group(2) or m.group(4)
        out.add((KIND.get(k, k), n.replace("-", ".")))
    return out


def trace(deck, sources):
    stext = "\n".join(read_source(p) for p in sources)
    sflat = norm(stext)
    slides = deck_slides(deck)
    rows, seen = [], set()
    for i, body, note in slides:
        for src_name, t in (("장", body), ("노트", note)):
            t = DATE.sub(" ", t)
            for ln in t.split("\n"):
                for m in NUM.finditer(ln):
                    num, unit = m.group(1), m.group(2) or ""
                    digits = re.sub(r"\D", "", num)
                    if not unit and "." not in num and (len(digits) < 2 or (num.startswith("0") and len(digits) <= 2)):
                        continue   # 한 자리 정수와 목차 번호(01, 02)는 번호나 개수라 뺌
                    key = (i, norm(num) + unit)
                    if key in seen:
                        continue
                    seen.add(key)
                    found = norm(num) in sflat
                    ctx = ln[max(0, m.start() - 14):m.end() + 10].strip()
                    rows.append((i, src_name, num + unit, found, ctx))
    dtext = "\n".join(b + "\n" + n for _, b, n in slides)
    miss_lab = sorted(labels(stext) - labels(dtext), key=lambda x: (x[0], float(x[1])))
    return rows, miss_lab, len(slides)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    out = sys.argv[sys.argv.index("--out") + 1] if "--out" in sys.argv else None
    if out and out in args:
        args.remove(out)
    if len(args) < 2:
        sys.exit("사용: source_trace.py 덱.pptx 원본1 [원본2 ...] [--out 표.md]")
    rows, miss_lab, n = trace(args[0], args[1:])
    lines = [f"# 원본 대조 ({os.path.basename(args[0])}, 원본 {len(args) - 1}개)", ""]
    bad = [r for r in rows if not r[3]]
    lines += [f"- 덱의 숫자 {len(rows)}개 중 원본에서 못 찾은 것 {len(bad)}개 (계산한 값이거나 옮기다 틀린 값: 하나씩 확인)",
              f"- 원본의 Figure, Table, Algorithm, Equation 번호 중 덱에 안 나오는 것 {len(miss_lab)}개 (덱에서 빠졌거나 출처 줄에 번호를 안 적음)", ""]
    if bad:
        lines += ["## 원본에서 못 찾은 숫자", "| 장 | 위치 | 숫자 | 앞뒤 글 |", "|---|---|---|---|"]
        lines += [f"| {i} | {w} | {num} | {ctx.replace('|', '/')} |" for i, w, num, _, ctx in bad]
        lines.append("")
    if miss_lab:
        lines += ["## 덱에 안 나오는 원본 번호", ", ".join(f"{k} {num}" for k, num in miss_lab), ""]
    lines.append("※ 숫자는 글자 대조이며 원본 PDF 에서 표나 그림 속 숫자는 글로 뽑히지 않을 수 있음. 차이 있음이 곧 오류는 아님")
    text = "\n".join(lines)
    if out:
        open(out, "w", encoding="utf8", newline="\n").write(text + "\n")
    print(text)


if __name__ == "__main__":
    main()
