"""부품을 모두 그려 보는 자체 시험"""
import datetime
import json
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ixdeck import *  # noqa: E402,F403

out = sys.argv[1]
d = Deck(author="AX Development / 점검")
d.cover("부품 자체 점검", "2026.11.05")
d.section("부품 자체 점검")

s = d.report_slide("부품 자체 점검", "보고형 3열", lead="세 패널을 화살표로 이은 보고형 장")
c1, c2, c3 = report_panels(s, ["원천", "규칙", "결과"])
c1.table([["Tag명", "Description"], ["CU1_F_DEN", "전면 전류 밀도"], ["CU1_B_DEN", "후면 전류 밀도"]],
         ratio=[1.2, 2.15], highlight=(0,)).arrow().notes(["같은 Tag명 다른 Description"])
c2.rule_list("ISH 데이터 생성 규칙", [("doc_id_green", "규칙 하나"), ("checklist_green", "규칙 둘")]).notes(["규칙 결론"])
c3.chip_tree([(0, "SM", "OperationData"), (1, "SMC", "Recipe"), (2, "Prop", "Front_Side"), (2, "Prop", "Back_Side"),
              (1, "SML", "Rectifier"), (2, "Ent", "Machine1"), (2, "Rel", "HasPart")], value_w=1.6)
bottoms = [c.done() for c in (c1, c2, c3)]
assert all(BODY_TOP < y <= BODY_BOTTOM + 1e-6 for y in bottoms), bottoms
try:
    Deck(author=" ")
    raise AssertionError("작성자 없이 Deck 이 만들어짐")
except ValueError:
    pass

s = d.doc_slide("1-1. 부품 자체 점검 (설계문서형)", "구성도")
src = frame(s, 0.75, 2.21, 1.1, 3.3, "Data Source", kind="source")
icon(s, "plc", 1.05, 2.8, w=0.4, name="PLC")
edge = frame(s, 2.1, 2.21, 4.12, 3.6, "Edge.PC(Machine.AX)", kind="edge", sub="On-Premise")
a = server(s, 2.5, 2.68, 1.5, 0.27, "DA Server")
area(s, 2.4, 3.4, 1.7, 1.3)
b = db(s, 2.55, 3.55, 1.4, 0.22, "PostgreSQL")
c = db(s, 2.55, 3.95, 1.4, 0.22, "InfluxDB")
plat = frame(s, 6.71, 2.20, 4.67, 3.9, "SDF Platform(Factory.AX)", kind="platform")
e = server(s, 9.3, 2.68, 1.6, 0.27, "Data Process Server")
f = db(s, 9.4, 3.5, 1.4, 0.22, "MongoDB")
zone(s, 8.9, 3.3, 2.3, 1.0, color=GREEN, alpha=0.08)
connect(s, a, "bottom", b, "top")
connect(s, a, "right", e, "left", color=ORANGE)
connect(s, c, "right", f, "bottom", color=BLUE)
connect(s, e, "bottom", f, "top", head="both")
g = db(s, 9.4, 4.6, 1.4, 0.22, "Jena Fuseki")
connect(s, c, "bottom", g, "left", color=GREEN, dash="sysDot")
connect(s, src, "right", a, "left")
tag(s, 6.3, 2.75)
step(s, 7.2, 2.6, 1, "파일 수집")
step(s, 7.2, 3.0, 2, "파싱 및 검증")
caption(s, 7.0, 5.3, 4, "Machine AAS 전체 Graph 데이터 정의 예시")
d.clone(s, 4, text="Data Lake", dy=-0.1)

# 설계문서형 자동 배치
s = d.doc_slide("1-2. 자동 배치 점검", "구성도")
g = DocGrid(s, right="source", rows=3)
src = g.source("source", [("plc", "PLC"), ("camera", "Camera")])
e = g.group("edge", "Edge.PC", kind="edge", sub="On-Premise", header="DA Server", cols=2, rows=(0, 1))
x1, x2 = e.box(0, 0, "PostgreSQL"), e.box(1, 1, "InfluxDB")
stack_bars(s, e.body.left / 914400, e.body.top / 914400 + e.body.height / 914400 + 0.1,
           e.body.width / 914400, "Edge.PC", "On-자산", n=1)
p = g.group("platform", "SDF Platform", header="Data Process Server", cols=2)
y1, y2 = p.box(0, 1, "MongoDB"), p.box(2, 0, "Jena Fuseki", kind="server", span=2)
assert abs(x1.top - y1.top) < 2, "같은 줄 상자 y 가 다름"
focus(y1)
p.zone(0, 0, 1, 1, color=GREEN)
connect(s, src, "right", x1, "left", text="수집")
connect(s, x1, "right", y1, "left", color=ORANGE, text="저장")
link(s, x2, y2, color=BLUE)
step_by(s, x1, 1, "수집", where="below_r", cond="변경사항 없는 경우")
legend(s, 11.7, 5.4, [(BLUE, "AAS Node"), (GREEN, "SemanticId Node")])
more(s, 12.1, 6.3)

# 검토에서 나온 경계 사례
s = d.doc_slide("1-3. 경계 사례", "구성도")
E = 914400
u, v = db(s, 2, 3, 1.4, 0.24, "U"), db(s, 5, 5, 1.4, 0.24, "V")
connect(s, u, "bottom", v, "top", text="LBL")
lab = s.shapes[-1]
assert abs(lab.top / E - (4.12 - 0.19)) < 0.05, lab.top / E
w1, w2 = db(s, 2, 1, 1.4, 0.24, "W1"), db(s, 6, 2.5, 1.4, 0.24, "W2")
connect(s, w1, "right", w2, "top", text="L")
lab = s.shapes[-1]
assert abs(lab.left / E - 6.75) < 0.05, lab.left / E
g = DocGrid(s, rows=4)
g.group("platform", "위", rows=(0, 1))
for bad in (lambda: g.group("platform", "아래", rows=(2, 3)),
            lambda: g.group("edge", "칸", cols=2).box(0, 1, "넘침", span=2),
            lambda: report_panels(s, [])):
    try:
        bad()
        raise AssertionError("막혀야 할 입력이 통과함")
    except ValueError:
        pass
one = Col(s, 1, 2)
one.box("하나", h=one.bottom - one.top)
one.done()
try:
    one.done()
    raise AssertionError("done 두 번")
except ValueError:
    pass

# link 자동
s = d.doc_slide("1-4. link 자동 연결", "구성도")
g = DocGrid(s, right="source", rows=3)
src = g.source("source", [("plc", "PLC"), ("camera", "Camera")])
e = g.group("edge", "Edge.PC", kind="edge", header="DA Server", cols=2)
a1, a2, a3 = e.box(0, 0, "A"), e.box(1, 1, "B"), e.box(2, 0, "C")
b1 = g.group("platform", "P", cols=2).box(1, 1, "D")
fan(s, a1, g.icons, step=(1, "읽기"))
diag = link(s, a1, a2)
vert = link(s, a1, a3)
link(s, a2, b1, text="HTTP", step=(2, "전송"), cond="변경 없는 경우")
prst = lambda c: c._element.spPr.find(qn("a:prstGeom")).get("prst")
assert prst(diag) == "bentConnector2", prst(diag)
assert prst(vert) == "straightConnector1", prst(vert)
link_slide = s

s = d.doc_slide("1-5. 겹침 검출", "구성도")
x1, x2 = db(s, 1, 3, 1, 0.24, "X"), db(s, 1.5, 3.1, 1, 0.24, "Y")
link(s, x1, db(s, 4, 3, 1, 0.24, "Z"))
bad_slide = s

# 2차 검토 재현 사례
s = d.doc_slide("1-6. 반올림, 대체 경로, 조건 문구", "구성도")
A = db(s, round(2.5 - 0.8 / 2, 3), 2.5, 0.8, 0.24, "A")
B = db(s, round(2.5 - 0.9 / 2, 3), 4.5, 0.9, 0.24, "B")
c = link(s, A, B)                                       # EMU 반올림 오차에도 곧은 선
assert c._element.spPr.find(qn("a:prstGeom")).get("prst") == "straightConnector1"
P, Q = db(s, 4.4, 2.7, 1.4, 0.24, "P"), db(s, 6.0, 3.0, 1.4, 0.24, "Q")
link(s, P, Q, step=(1, "짧"), cond="변경사항이 없는 경우에는 이 조건 문구가 매우 길게 이어진다")
M, N = db(s, 8, 2.5, 1, 0.24, "M"), db(s, 10, 2.5, 1, 0.24, "N")
connect(s, M, "right", N, "left")
link(s, M, N)
extra_slide = s

# 점선 다발
s = d.report_slide("점선 다발 점검", "표 행과 칩", lead="표 행마다 링크 아이콘을 거쳐 칩으로")
k1, k2, k3 = cols = report_panels(s, ["표", "칩", "기타"], hub_gap=1)
assert abs((cols.gaps[1][1] - cols.gaps[1][0]) - 1.38) < 0.02 and abs((cols.gaps[0][1] - cols.gaps[0][0]) - 0.67) < 0.01
k1.notes(["왼쪽"])
k2.table([["용어", "SemanticId"], ["Rectifier", "/ids/1"], ["Frontside", "/ids/2"], ["Backside", "/ids/3"]])
k3.chip_tree([(0, "SM", "표준화된 데이터 구조"), (1, "SML", "Rectifier"), (2, "SMC", "Frontside"), (3, "Prop", "MeasurementCurrent"),
              (2, "SMC", "Backside")])
for c in cols:
    c.done()
chips = [k for k, _ in k3[0]]
assert all(abs((v.left + v.width) - (k3[0][0][1].left + k3[0][0][1].width)) < 914400 * 0.01 for _, v in k3[0]), "칩 오른쪽 끝"
n0 = len([e for e in s.shapes._spTree.iter(qn("p:cxnSp"))])
bundle(s, row_anchors(s, k2[0]), [chips[1], chips[2], chips[4]], cols.gaps[1], "SemanticId 기반 연결")
cx = [e for e in s.shapes._spTree.iter(qn("p:cxnSp"))][n0:]
assert len(cx) == 6, len(cx)
assert all(e.find(".//" + qn("a:prstDash")).get("val") == "sysDot" for e in cx)
assert all(e.find(".//" + qn("a:headEnd")).get("type") == "arrow" for e in cx)
bundle_slide = s

# 범용 장 부품과 복잡한 구성도
s = d.report_slide("범용 장 점검", "차트, 수치, 목록", lead="범용 장 부품")
a, b, c = cols = report_panels(s, ["수치", "차트", "목록"])
a.kpis([("15", "/ 22", "이름", "설명")]).kpis([("1", "건", "이름")])
b.chart("col", ["가", "나"], {"값": [1, 2]}, h=1.5, highlight=(1,), number_format='0"개"') \
 .chart("line", ["1월", "2월", "3월"], {"A": [1, 3, 2], "B": [2, 1, 3]}, h=1.5)
c.section("제목", ["줄 하나", "줄 둘"], highlight=(1,)).steps([(1, "제목", "설명"), (2, "제목", ["줄 1", "줄 2"])])
for col in cols:
    col.done()
cx_ = [e for e in s.shapes._spTree.iter(qn("p:graphicFrame"))]
assert len(cx_) == 2, len(cx_)
s = d.report_slide("일정 및 그림 점검", "gantt, image", lead="일정")
a, b = cols = report_panels(s, ["일정", "그림"], arrows=False)
a.gantt(["1", "2", "3"], [("가", 0, 1, "done"), ("나", 1, 2.5, "now", "진행"), ("다", 2.9, 0, "mile", "끝")], today=1.5)
b.image(os.path.join(ICONS, "_catalog.png"), h=2.0, marks=[(0.5, 0.5, 1)])
for col in cols:
    col.done()
s = d.doc_slide("1-7. 복잡한 구성도 점검", "구성도")
g = DocGrid(s, right=None, rows=6, bottom=6.3)
g.source("source", [("plc", "PLC"), ("camera", "Camera"), ("folder", "CSV")])
e = g.group("edge", "Edge", kind="edge", header="DA", cols=2, area=False)
sa = e.sub(0, 0, 3, header="수집")
x1 = sa.box(1, "가")
sb = e.sub(1, 0, 3, header="저장")
x2 = sb.box(3, "나")
p = g.group("platform", "Platform", cols=2, area=False)
y1 = p.sub(0, 0, 2, header="서버", caption="설명").box(2, "다")
bar_ = g.bar("platform", "Data Lake", tag_="EAC")
g.stack("edge", "Edge.PC", "On-자산", n=1)
fan(s, x1, g.icons, color=ORANGE, step=(1, "읽기"))
link(s, x1, y1, step=(2, "보냄"))
link(s, x2, bar_, color=BLUE, step=(3, "저장"))
complex_slide = s

ids = [int(e.get("id")) for e in s.shapes._spTree.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")]
assert len(ids) == len(set(ids)), "도형 번호 겹침"

d.save(out, strict=False)
z = zipfile.ZipFile(out)
slides = [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
x = "".join(z.read(n).decode("utf8") for n in slides)
assert len(slides) == 14, slides   # 표지, 목차, 본문 11, Thank you
assert x.count("<p:cxnSp>") == x.count("<a:stCxn") == x.count("<a:endCxn"), "연결선이 도형에 안 붙음"
assert "Pretendard" not in x and "Malgun" not in x
from pptx import Presentation  # noqa: E402
assert len(Presentation(out).slides) == 14
_ly = [sl.slide_layout.name for sl in Presentation(out).slides]
assert _ly[1] == "5_contents-02" and _ly[-1] == "앤드페이지_다크네이비", _ly
assert x.count('<a:spcPct val="125000"/>') > 100
_lp = layout_problems
layout_problems = lambda sl: [m for m in _lp(sl) if not m.startswith("패널 칸")]  # 패널을 일부러 비움
assert layout_problems(link_slide) == [], layout_problems(link_slide)
assert layout_problems(bad_slide), "겹침을 못 잡음"
# 둘째 선은 피할 길 없음 경고만 나와야 함
assert all(m.startswith("상자를 피하는 선을 못 찾음") for m in layout_problems(extra_slide)), layout_problems(extra_slide)
assert layout_problems(bundle_slide) == [], layout_problems(bundle_slide)
assert layout_problems(complex_slide) == [], layout_problems(complex_slide)
d.save(out, strict=False)
_d = Deck(author="소속 / 홍길동 연구원")   # 파일 이름 규칙
_n = _d.filename("과제 이름_예시", "v0.9.0", "x", datetime.date(2026, 1, 2))
assert os.path.basename(_n) == "20260102_과제이름예시_홍길동_v0.9.0.pptx", _n
f = Deck(author="시험", forbid=[r"192\.168\.\d+\.\d+"])   # forbid 가 저장을 막아야 함
f.cover("시험", "2026.11.10")
f.section("시험")
text(f.page("시험", lead="요약"), 1, 2, 3, 0.3, "서버 192.168.0.1")
try:
    f.save(out + ".f.pptx")
    raise AssertionError("forbid 를 못 잡음")
except ValueError as e:
    assert "forbid" in str(e), e
_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # 판 번호 셋이 같아야 함
_v = [json.load(open(os.path.join(_root, f), encoding="utf8")) for f in ("package.json", ".claude-plugin/plugin.json",
                                                                       ".claude-plugin/marketplace.json")]
assert _v[0]["version"] == _v[1]["version"] == _v[2]["plugins"][0]["version"], "판 번호가 서로 다름"
print("OK", out, "cxn", x.count("<p:cxnSp>"))
