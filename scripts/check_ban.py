"""금지 표현 검사. 종료 코드 = 발견 건수"""
import re
import sys
from pptx import Presentation

BAN = re.compile(r"축으로|까닭|뜻|몫|(?<![빈])칸|[가-힣]봄\b|씀\b|뺌\b|나옴\b|·|—|–"
                 r"|(?<![가-힣])(?:늘|딱|쭉|좀|그냥|되게|엄청|그만큼|이만큼)(?![가-힣])"  # 구어 부사
                 r"|(?<![가-힣])[새옛](?= )"  # 새 ~, 옛 ~ 대신 역할 이름
                 r"|(?<=[\d)가-힣] )x(?= [\d(가-힣])")  # 곱하기는 ×

def texts(shape):
    if shape.has_text_frame:
        yield shape.text_frame.text
    if getattr(shape, "has_table", False) and shape.has_table:
        for r in shape.table.rows:
            for c in r.cells:
                yield c.text
    if getattr(shape, "has_chart", False) and shape.has_chart:  # 차트 항목, 계열 이름
        for plot in shape.chart.plots:
            yield from (str(c) for c in plot.categories)
            yield from (sr.name or "" for sr in plot.series)
        for ax in ("category_axis", "value_axis"):   # 축 이름
            try:
                a_ = getattr(shape.chart, ax)
                if a_.has_title:
                    yield a_.axis_title.text_frame.text
            except ValueError:
                pass
    if shape.shape_type == 6:
        for s in shape.shapes:
            yield from texts(s)

def allowed(prs):
    """Deck(allow=) 예외를 지우는 함수 (없으면 None)"""
    import json
    k = prs.core_properties.keywords or ""
    if not k.startswith("ix-allow:"):
        return None
    pats = [re.compile(p) for p in json.loads(k[len("ix-allow:"):])]

    def clean(t):
        for p in pats:   # 하나씩: (?i) 같은 전역 깃발이 섞여도 됨
            t = p.sub(" ", t)
        return t
    return clean


def slide_texts(slide):
    """도형 글 및 수식 상자 글 (mc:Fallback)"""
    for sh in slide.shapes:
        yield from texts(sh)
    mc, a = "{http://schemas.openxmlformats.org/markup-compatibility/2006}", \
        "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for fb in slide._element.iter(mc + "Fallback"):
        yield "\n".join("".join(t.text or "" for t in p.iter(a + "t")) for p in fb.iter(a + "p"))


if __name__ == "__main__":
    n = 0
    for f in sys.argv[1:]:
        prs = Presentation(f)
        ok = allowed(prs)
        for i, s in enumerate(prs.slides, 1):
            for t in slide_texts(s):
                t = ok(t) if ok else t
                for m in BAN.finditer(t):
                    n += 1
                    print(f"{f} {i}쪽: '{m.group()}' ... {t[max(0, m.start()-15):m.end()+15]!r}")
    print("금지 표현", n)
    sys.exit(min(n, 255))
