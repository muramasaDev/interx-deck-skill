"""부품을 모두 그려 보는 자체 시험"""
import datetime
import hashlib
import json
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))
from ixdeck import *  # noqa: E402,F403
from ixdeck import _linear, _math, _reg, etree  # noqa: E402

out = sys.argv[1]
d = Deck(author="AX Development / 점검")
d.cover("부품 자체 점검", "2026.11.05")
d.section("부품 자체 점검")

s = d.report_slide("부품 자체 점검", "보고형 3열", lead="세 패널을 화살표로 이은 보고형 장")
c1, c2, c3 = report_panels(s, ["원천", "규칙", "결과"])
c1.table([["Tag명", "Description"], ["CU1_F_DEN", "전면 전류 밀도"], ["CU1_B_DEN", "후면 전류 밀도"]],
         ratio=[1.2, 2.15], highlight=(0,), title="원천 Tag 설명").arrow().notes(["같은 Tag명 다른 Description"])
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
_reg(s)["figure"] = "부품 구성도"
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
connect(s, a, "right", e, "left", color=ACCENT)
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
_reg(s)["figure"] = "자동 배치 구성도"
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
connect(s, x1, "right", y1, "left", color=ACCENT, text="저장")
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
_reg(s)["figure"] = "link 연결 구성도"
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
_reg(s)["figure"] = "경로 및 조건 구성도"
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
_reg(s)["figure"] = "점선 다발"
assert abs((cols.gaps[1][1] - cols.gaps[1][0]) - 1.38) < 0.02 and abs((cols.gaps[0][1] - cols.gaps[0][0]) - 0.67) < 0.01
k1.notes(["왼쪽"])
k2.table([["용어", "SemanticId"], ["Rectifier", "/ids/1"], ["Frontside", "/ids/2"], ["Backside", "/ids/3"]],
         title="용어 및 SemanticId")
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
b.chart("col", ["가", "나"], {"값": [1, 2]}, h=1.5, highlight=(1,), number_format='0"개"', title="막대 점검") \
 .chart("line", ["1월", "2월", "3월"], {"A": [1, 3, 2], "B": [2, 1, 3]}, h=1.5, title="선 점검")
c.section("제목", ["줄 하나", "줄 둘"], highlight=(1,)).steps([(1, "제목", "설명"), (2, "제목", ["줄 1", "줄 2"])])
for col in cols:
    col.done()
cx_ = [e for e in s.shapes._spTree.iter(qn("p:graphicFrame"))]
assert len(cx_) == 2, len(cx_)
s = d.report_slide("일정 및 그림 점검", "gantt, image", lead="일정")
a, b = cols = report_panels(s, ["일정", "그림"], arrows=False)
a.gantt(["1", "2", "3"], [("가", 0, 1, "done"), ("나", 1, 2.5, "now", "진행"), ("다", 2.9, 0, "mile", "끝")], today=1.5, title="일정 점검")
b.image(os.path.join(ICONS, "_catalog.png"), h=2.0, marks=[(0.5, 0.5, 1)], title="아이콘 목록 화면")
for col in cols:
    col.done()
s = d.doc_slide("1-7. 복잡한 구성도 점검", "구성도")
_reg(s)["figure"] = "복잡한 구성도"
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
fan(s, x1, g.icons, color=ACCENT, step=(1, "읽기"))
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
assert "Pretendard" in x and "Noto Sans" not in x and "Malgun" not in x   # 글꼴은 Pretendard 하나
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
_s = d.doc_slide("1-9. note_by 색", "구성도")   # 머리 글은 실패 경로 빨강과 달라야 함
note_by(_s, db(_s, 3, 3, 1.4, 0.24, "상자"), "머리 글")
assert _reg(_s)["pending"][-1][4] != RED, "note_by 머리 글이 빨강"
_d = Deck(author="소속 / 홍길동 연구원")   # 파일 이름 규칙
_n = _d.filename("과제 이름_예시", "v0.9.0", "x", datetime.date(2026, 1, 2))
assert os.path.basename(_n) == "과제_이름_예시_홍길동_v0.9.0_260102.pptx", _n
f = Deck(author="시험", forbid=[r"192\.168\.\d+\.\d+"])   # forbid 가 저장을 막아야 함
f.cover("시험", "2026.11.10")
f.section("시험")
text(f.page("시험", lead="요약"), 1, 2, 3, 0.3, "서버 192.168.0.1")
try:
    f.save(out + ".f.pptx")
    raise AssertionError("forbid 를 못 잡음")
except ValueError as e:
    assert "forbid" in str(e), e
_x = etree.tostring(_math(r"J_{보정} = J_{0} \times \frac{T + E}{\hat{d}}", 12, True), encoding="unicode")   # 수식
assert all(t in _x for t in ("<m:f>", "<m:sSub>", "<m:acc>", "\U0001D43D", ">보정<", ">×<")), _x
assert _linear(r"J_{0} \times \frac{T}{E}") == "J_0 × (T) / (E)", _linear(r"J_{0} \times \frac{T}{E}")
_root = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))   # 판 번호 셋이 같아야 함
_vs = {f: json.load(open(os.path.join(_root, f), encoding="utf8")) for f in
       ("package.json", ".claude-plugin/plugin.json", ".claude-plugin/marketplace.json")
       if os.path.exists(os.path.join(_root, f))}   # 설치본에는 package.json 이 없음
_vs = {f: (v["plugins"][0]["version"] if "plugins" in v else v["version"]) for f, v in _vs.items()}
assert len(set(_vs.values())) == 1, f"판 번호가 서로 다름: {_vs}"
# 배치 1: 검사와 발표자 노트, 도장, 숫자 단위 붙임
import subprocess  # noqa: E402
import tempfile  # noqa: E402
import check_ban  # noqa: E402
import check_text  # noqa: E402
import deck_text  # noqa: E402
import update_imnotai  # noqa: E402

_here = os.path.dirname(os.path.realpath(__file__))
_tmp = tempfile.mkdtemp(prefix="ixself_")


def _chart(s, x, y, w, h, *a, **kw):
    """chart + 아래 가운데 그림 제목"""
    g = chart(s, x, y, w, h, *a, **kw)
    figure_caption(s, "점검 차트", x, y + h + 0.02, w)
    return g


def _gantt(s, x, y, w, periods, tasks, **kw):
    """gantt + 아래 가운데 그림 제목"""
    g = gantt(s, x, y, w, periods, tasks, **kw)
    figure_caption(s, "점검 일정", x, y + 0.27 + 0.3 * len(tasks) + 0.02, w)
    return g


def _mini(name, build, **kw):
    """장 몇 개짜리 시험 덱 (파일 이름 규칙을 지킴)"""
    m = Deck(author="소속 / 점검", **kw)
    m.cover("점검", "2026.11.16")
    m.section("점검")
    build(m)
    p_ = os.path.join(_tmp, f"점검_{name}_v0.1.0_261116.pptx")
    m.save(p_, strict=False)
    return p_


def _py(script, *args, env=None):
    r = subprocess.run([sys.executable, "-P", os.path.join(_here, script), *args], capture_output=True,
                       env={**os.environ, "PYTHONIOENCODING": "utf-8", **(env or {})})
    return r.returncode, (r.stdout + r.stderr).decode("utf8", "replace")


def _p1(m):
    s = m.page("숫자와 단위", lead="요약", notes="말할 내용 하나")
    text(s, 1, 2, 6, 0.3, "정확도 95.7 %, 24 시간, 범위 1 ~ 4 개, 3 index")
    stamp(s, "예시")
    try:
        stamp(s, "또")
        raise AssertionError("도장 두 개가 통과함")
    except ValueError:
        pass
    return s


_p = _mini("a", _p1)
_zp = zipfile.ZipFile(_p)
_x = "".join(_zp.read(n).decode("utf8") for n in _zp.namelist() if n.startswith("ppt/slides/slide"))
assert "95.7 %" in _x and "24 시간" in _x and "4 개" in _x, "숫자 단위 사이가 NBSP 가 아님"
assert "1 ~ 4" in _x and "3 index" in _x, "범위와 영어 낱말은 그대로여야 함"
_sl = Presentation(_p).slides[2]
assert _sl.notes_slide.notes_text_frame.text == "말할 내용 하나", "발표자 노트가 없음"
_st = [sh for sh in _sl.shapes if sh.name == "도장"]
assert len(_st) == 1 and abs(_st[0].rotation - 352) < 0.01, "도장이 없거나 기울지 않음"
for _bad in (lambda: stamp(_sl, " "), lambda: Deck(author="a / b", purpose="없는 쓰임새"),
             lambda: Deck(author="a / b", minutes=0)):
    try:
        _bad()
        raise AssertionError("막혀야 할 입력이 통과함")
    except ValueError:
        pass
assert deck_text.titles(_p).strip() == "- 3쪽 1-1. 숫자와 단위: 요약", deck_text.titles(_p)
_rc, _o = _py("check_ban.py", _p)
assert _rc == 0, _o

# 금지 표현: 자리표시자
for _t in ("TODO 정리", "TBD", "xxx 값", "lorem ipsum", "[Insert 이름]", "여기에 텍스트를 입력", "예시 값", "요약 한 줄"):
    assert check_ban.BAN.search(_t), _t
for _t in ("Mastodon 서버", "xx 와 xxxl", "예시 장"):
    assert not check_ban.BAN.search(_t), _t


# check.py: 출처 줄, 첫 본문 장 요약, 발표 시간
def _p2(m):
    _chart(m.page("차트", lead="요약"), 1, 2, 5, 3, "col", ["가", "나"], {"값": [1, 2]}, number_format='0"개"')


_rc, _o = _py("check.py", _mini("b", _p2))
assert "출처 줄이 없음" in _o, _o


def _p3(m):
    s = m.page("차트", lead="요약")
    _chart(s, 1, 2, 5, 3, "col", ["가", "나"], {"값": [1, 2]}, number_format='0"개"')
    notes(s, 1, 5.5, 6, ["출처: 시험 결과서 (2026.05)"])


_rc, _o = _py("check.py", _mini("c", _p3))
assert "출처 줄이 없음" not in _o, _o
_rc, _o = _py("check.py", _mini("d", _p1, purpose="사내보고"))
assert "'요약' 이 없음" in _o, _o


def _p4(m):
    m.page("요약 및 현황", lead="요약")
    m.page("둘째", lead="요약")
    m.page("셋째", lead="요약")


_pe = _mini("e", _p4, purpose="제안", minutes=3)
_rc, _o = _py("check.py", _pe)
assert "'요약' 이 없음" not in _o and "주의: 본문 3장이 발표 3분" in _o, _o
assert Presentation(_pe).core_properties.category == "ix-purpose:제안"


# check_text: 어휘 사전과 노트
assert len(check_text.load_lexicon()) > 50, "im-not-ai 어휘 사전을 못 읽음"


def _p5(m):
    s = m.page("어휘", lead="요약", notes="다양한 사례를 든다. 팀 역량 잠금 해제")
    text(s, 1, 2, 6, 0.3, "팀 역량 잠금 해제 방법")


_f, _w = check_text.check(_mini("f", _p5))
assert any("3쪽 어휘 AST-01" in x for x in _w), _w
_f2, _w2 = check_text.check(_mini("g", lambda m: text(m.page("어휘", lead="요약"), 1, 2, 6, 0.3, "화면 잠금 해제 방법")))
assert not any("AST-01" in x for x in _w2), ("화면 잠금은 제외 규칙에 걸려야 함", _w2)
assert any("3쪽 노트 어휘 AST-01" in x for x in _w), _w
assert any("3쪽 노트 '다양한'" in x for x in _f), _f
assert not any("서술체" in x for x in _f), "노트에는 서술체 규칙을 쓰지 않음"

# render.py: 글꼴 없으면 종료 코드 3 (폴더, 등록 값, fc-list)
import render  # noqa: E402

_rf = (render._font_dirs, render._registry_fonts, render._fc_families)
_fd = tempfile.mkdtemp(prefix="ixfont_")
render._font_dirs = lambda: [_fd]
render._registry_fonts, render._fc_families = (lambda: ["Arial (TrueType)"]), (lambda: "DejaVu Sans")
assert not render.font_installed()
render._registry_fonts = lambda: ["Arial (TrueType)", "Pretendard Regular (TrueType)"]
assert render.font_installed(), "등록 값 Pretendard 를 못 봄"
render._registry_fonts, render._fc_families = (lambda: []), (lambda: "Pretendard:style=Regular")
assert render.font_installed(), "fc-list 의 Pretendard 를 못 봄"
render._fc_families = lambda: ""
open(os.path.join(_fd, "Pretendard-Regular.otf"), "w").close()
assert render.font_installed(), "글꼴 폴더 파일을 못 봄"
os.remove(os.path.join(_fd, "Pretendard-Regular.otf"))
assert not render.font_installed()
_argv = sys.argv
sys.argv = ["render.py", _p, os.path.join(_tmp, "r")]
try:
    render.main()
    raise AssertionError("글꼴이 없는데 렌더가 시작됨")
except SystemExit as _ex:
    assert _ex.code == 3, _ex.code
sys.argv = _argv
render._font_dirs, render._registry_fonts, render._fc_families = _rf

# update_imnotai.py: 받기 실패면 기존 파일 유지, 종료 코드 1
_k = tempfile.mkdtemp(prefix="ixlex_")
open(os.path.join(_k, "lexicon.json"), "w").write("keep")
update_imnotai.RAW = "https://127.0.0.1:1/"
assert update_imnotai.main(_k) == 1 and open(os.path.join(_k, "lexicon.json")).read() == "keep"

# 배치 2: 새 장 유형과 검사
import contextlib  # noqa: E402
import io  # noqa: E402


def _strict(build, **kw):
    """strict 저장 결과: 실패 메시지 (통과면 빈 글)"""
    m = Deck(author="소속 / 점검", **kw)
    m.cover("점검", "2026.11.16")
    m.section("점검")
    build(m)
    try:
        m.save(os.path.join(_tmp, "점검_strict_v0.1.0_261116.pptx"))
    except ValueError as e:
        return str(e)
    return ""


_m0 = Deck(author="소속 / 점검")
_m0.section("점검")
_s0 = _m0.page("블록", lead="요약")

# 글자 한도: 한도 + 1 자는 실패, 한도 자는 통과 (블록 이름이 메시지에 있음)
for _k, _fn in (("num_cards 제목", lambda c, n: c.num_cards([("가" * n, "설명")])),
                ("num_cards 설명", lambda c, n: c.num_cards([("제목", "가" * n)])),
                ("tab_box 줄", lambda c, n: c.tab_box("결론", ["가" * n])),
                ("kpis 이름", lambda c, n: c.kpis([("1", "건", "가" * n)])),
                ("section 행", lambda c, n: c.section("제목", ["가" * n])),
                ("rule_list 글", lambda c, n: c.rule_list("제목", [("doc_id_green", "가" * n)]))):
    for _n, _hit in ((LIMITS[_k] + 1, True), (LIMITS[_k], False)):
        _e = _strict(lambda m, fn=_fn, n=_n: fn(Col(m.page("한도", lead="요약"), 1, LIMIT_REF_W), n))
        assert (f"{_k} " in _e) == _hit, (_k, _n, _e)
    for _w, _lim in ((LIMIT_REF_W * 2, LIMITS[_k] * 2), (LIMIT_REF_W / 2, LIMITS[_k] // 2)):   # 칸 폭에 비례
        for _n, _hit in ((_lim + 1, True), (_lim, False)):
            _e = _strict(lambda m, fn=_fn, n=_n, w=_w: fn(Col(m.page("한도", lead="요약"), 1, w), n))
            assert (f"{_k} " in _e) == _hit, (_k, _w, _n, _e)

# 차트 단위: y_title 이나 number_format 의 단위 글자가 있어야 저장됨
for _kw, _ok in (({}, False), ({"y_title": "건수 (건)"}, True), ({"number_format": '0"건"'}, True),
                 ({"number_format": "0.0%"}, True)):
    _e = _strict(lambda m, kw=_kw: _chart(m.page("차트", lead="요약"), 1, 2, 5, 3, "col", ["가", "나"], {"값": [1, 2]}, **kw))
    assert ("y_title" in _e) == (not _ok), (_kw, _e)
_buf = io.StringIO()
with contextlib.redirect_stdout(_buf):
    _mini("h", lambda m: _chart(m.page("차트", lead="요약"), 1, 2, 5, 3, "col", ["가", "나"],
                               {k_: [1, 2] for k_ in "abcde"}, number_format='0"건"'))
assert "계열 5개" in _buf.getvalue(), _buf.getvalue()
_ch = _mini("i", lambda m: _chart(m.page("차트", lead="요약"), 1, 2, 5, 3, "bar", ["가", "나"], {"값": [1, 2]},
                                y_title="처리 건수 (건)"))
assert "처리 건수 (건)" in "".join(texts for sh_ in Presentation(_ch).slides[2].shapes if sh_.has_chart
                              for texts in [sh_.chart.value_axis.axis_title.text_frame.text]), "y_title 이 축에 없음"

# KPI 표, 데이터 수집 포인트 표
_kr = [("수집 성공률", "측정 안 함", "99.5% 이상", "저장 회차 / 수집 회차 × 100", "4차 시험 결과서")]
for _bad in (lambda: Col(_s0, 1, 11).kpi_table([]), lambda: Col(_s0, 1, 11).kpi_table([("가", "나", "다", "라", " ")]),
             lambda: Col(_s0, 1, 11).kpi_table([("가", "나", "다", "라")]),
             lambda: Col(_s0, 1, 11).collect_table([("도금", "전류", "가끔")]),
             lambda: Col(_s0, 1, 11).collect_table([("도금", "", "자동")])):
    try:
        _bad()
        raise AssertionError("막혀야 할 표 입력이 통과함")
    except ValueError:
        pass
_c0 = Col(_s0, 1, 11).collect_table([("도금", "전류", "자동"), ("검사", "두께", "반자동"), ("출하", "Lot", "수동")], title="수집 항목")
_c0.done()
assert [c.text for c in _c0[0].table.rows[0].cells] == ["공정", "데이터", "집계 방법"]


def _kt(src):
    def b(m):
        s = m.page("KPI", lead="요약")
        Col(s, 1, 11, top=2, bottom=2.8).kpi_table(_kr, title="수집 KPI").done()
        if src:
            notes(s, 1, 3, 10, [src if isinstance(src, str) else "출처: 4차 시험 결과서 (2026.05)"])
    return b


assert "KPI 표" in _strict(_kt(False)) and "KPI 표" not in _strict(_kt(True))
assert "KPI 표" in _strict(_kt("기준: 4차 시험 결과서")), "'기준:' 은 출처가 아님"
assert "KPI 표" not in _strict(_kt("자료: 4차 시험 결과서"))


# 평가항목 대응표
def _cr(both):
    def b(m):
        m.criteria(["기술 우수성", "수행 일정"])
        m.page("가", lead="요약", covers=["기술 우수성"])
        if both:
            m.page("나", lead="요약", covers=["수행 일정", "기술 우수성"])
    return b


assert "수행 일정" in _strict(_cr(False))
assert _strict(_cr(True)) == ""
_cp = os.path.join(_tmp, "점검_strict_v0.1.0_261116.pptx")
_cs = Presentation(_cp).slides
_ct = next(sh_ for sh_ in _cs[-2].shapes if sh_.has_table).table
assert _cs[-1].slide_layout.name == "앤드페이지_다크네이비" and "평가항목 대응표" in _cs[-2].placeholders[11].text_frame.text
assert [[c.text for c in r.cells] for r in _ct.rows][1:] == [["기술 우수성", "3쪽, 4쪽"], ["수행 일정", "4쪽"]]
_rc, _o = _py("check.py", _cp)
assert "slide5:" not in _o, _o   # 자동으로 만든 장은 check.py 를 통과
for _bad in (lambda: _m0.criteria([]), lambda: _m0.criteria(["가", "가"]),
             lambda: _m0.page("없는 항목", covers=["없음"])):
    try:
        _bad()
        raise AssertionError("막혀야 할 평가항목 입력이 통과함")
    except ValueError:
        pass


# 일정 및 산출물
def _dl(names, deliv):
    def b(m):
        _gantt(m.page("일정", lead="요약"), 1, 2, 10, ["1월", "2월"], [(n, 0, 1, "done") for n in names])
        m.deliverables(deliv)
    return b


assert _strict(_dl(["설계", "구현"], {"설계": ["설계서"], "구현": ["프로그램"]})) == ""
_e = _strict(_dl(["설계", "구현"], {"설계": ["설계서"], "시험": ["결과서"]}))
assert "'시험'" in _e and "'구현'" in _e, _e
assert _m0.deliverables({"설계": ["설계서", "정의서"]}) == [["단계", "산출물"], ["설계", "설계서, 정의서"]]
try:
    _m0.deliverables({"설계": []})
    raise AssertionError("빈 산출물이 통과함")
except ValueError:
    pass

# 영업, 문서 쓰임새
_rc, _o = _py("check.py", _mini("j", lambda m: [m.page(n, lead="요약") for n in ("회사 소개", "변화")], purpose="영업"))
assert "첫 본문 장" in _o and "6 ~ 15" in _o, _o
_rc, _o = _py("check.py", _mini("k", lambda m: [m.page(n, lead="요약") for n in ("변화", "회사 소개", "회사소개 상세")], purpose="영업"))
assert "회사 소개 장 2" in _o and "첫 본문 장" not in _o, _o
_rc, _o = _py("check.py", _mini("l", lambda m: [m.page(n, lead="요약") for n in
                                                ("변화", "승자와 패자", "목표 상태", "기능", "고객 사례", "회사 소개")],
                                purpose="영업"))
assert "6 ~ 15" not in _o and "회사 소개 장" not in _o and "첫 본문 장" not in _o, _o


def _tabs(m):
    for k_ in range(5):
        s_ = m.page(f"표 {k_}", lead="요약")
        if k_ < 3:
            table(s_, 1, 2, [3, 3], [["가", "나"], ["1", "2"]])
            notes(s_, 1, 3, 6, ["출처: 시험 결과서 (2026.05)"])


_rc, _o = _py("check.py", _mini("m", _tabs))
assert "3/5 (절반 넘음)" in _o, _o
_rc, _o = _py("check.py", _mini("n", _tabs, purpose="문서"))
assert "절반 넘음" not in _o, _o
_long = "수집 성공률은 4차 시험에서 99.8%로 검수 기준을 충족했고 " * 3
assert 70 < len(_long) < 110
_f, _w = check_text.check(_mini("o", lambda m: text(m.page("긴 줄", lead="요약"), 1, 2, 11, 0.3, _long)))
assert any("자 문장" in x for x in _w), _w
_f, _w = check_text.check(_mini("p", lambda m: text(m.page("긴 줄", lead="요약"), 1, 2, 11, 0.3, _long), purpose="문서"))
assert not any("자 문장" in x for x in _w), _w
assert "승리 주제" in open(os.path.join(_root, "skills", "using", "references", "review_prompt.md"), encoding="utf8").read()

# Mermaid: 편집 가능한 도형과 붙은 선
_FLOW = "graph TD" + chr(10) + chr(10).join([" A[시작] --> B[읽기]", " B --> C[검증]", " B -- 실패 --> D[재시도]",
                                              " C --> E[저장]", " C --> F[알림]", " D --> G[기록]"])
_SEQ = chr(10).join(["sequenceDiagram", " participant A as 설비", " participant B as 수집기", " participant C as 서버",
                    " A->>B: 값 요청", " B-->>A: 현재 값", " B->>C: 값 전송", " C-->>B: 저장 완료"])


def _mm(m):
    global _mm_shapes
    _mm_shapes = mermaid(m.page("흐름", lead="요약"), _FLOW, 1.0, 2.0, 11.3, 4.8, title="읽기 흐름")
    mermaid(m.page("순서", lead="요약"), _SEQ, 1.0, 2.0, 11.3, 4.8, title="값 전송 순서")


assert _strict(_mm) == ""
_mp = os.path.join(_tmp, "점검_strict_v0.1.0_261116.pptx")
_rc, _o = _py("check.py", _mp)   # 연결선 양 끝 붙음, 번호 중복, 장 밖 도형 포함
assert _rc == 0 and _o.startswith("OK"), _o
_x3, _x4 = (etree.tostring(sl_._element, encoding="unicode") for sl_ in list(Presentation(_mp).slides)[2:4])
assert _x3.count("<p:cxnSp>") == 6 == _x3.count("<a:stCxn") == _x3.count("<a:endCxn"), "flowchart 선 6개가 붙어야 함"
assert _x4.count("<p:cxnSp>") == 4 == _x4.count("<a:endCxn") and "활성 막대" in _x4 and 'prst="line"' in _x4
assert set(_mm_shapes) == set("ABCDEFG") and _mm_shapes["A"].text_frame.text == "시작"
for _bad_src in ("graph TD|A --- B", "graph TD|A{결정} --> B", "graph RL|A --> B", "pie|A", "graph TD|A --> A",
                 "graph TD|A -.-> B", "sequenceDiagram|loop 반복", "sequenceDiagram|A->>A: 자기", "sequenceDiagram|participant A",
                 "graph TD|A -->", ""):
    _bad_src = _bad_src.replace("|", chr(10))
    try:
        mermaid(_m0.page("지원 밖", lead="요약"), _bad_src, 1.0, 2.0, 11.3, 4.8)
        raise AssertionError(f"지원 밖 Mermaid 가 통과함: {_bad_src!r}")
    except ValueError:
        pass
try:   # 영역에 안 들어가는 크기
    mermaid(_m0.page("좁음", lead="요약"), _SEQ, 1.0, 2.0, 3.0, 1.0)
    raise AssertionError("좁은 영역이 통과함")
except ValueError:
    pass

# 기존 덱 고치기: 지정한 장 글만 바뀌고 나머지 파트는 바이트 그대로
_eo = os.path.join(_tmp, "edit_out.pptx")
_sha0 = hashlib.sha256(open(_mp, "rb").read()).hexdigest()
_rc, _o = _py("edit_deck.py", _mp, _eo, "--replace", "3", "시작", "개시", "3", "읽기", "값 20 개 읽기")
assert _rc == 0 and "3쪽" in _o and "중 다른 것 0개" in _o, _o
with zipfile.ZipFile(_mp) as _za, zipfile.ZipFile(_eo) as _zb:
    _diff = [n_ for n_ in _za.namelist() if _za.read(n_) != _zb.read(n_)]
    assert _za.namelist() == _zb.namelist() and len(_diff) == 1, _diff
    _x = _zb.read(_diff[0]).decode("utf8")
    assert "개시" in _x and "시작" not in _x and "20" + chr(160) + "개 읽기" in _x, "바뀐 글과 줄바꿈 방지 공백"
assert hashlib.sha256(open(_mp, "rb").read()).hexdigest() == _sha0, "원본이 바뀜"
for _args in (["3", "없는 글", "새 글"], ["99", "시작", "새 글"], ["3", "시작"]):
    _rc, _o = _py("edit_deck.py", _mp, os.path.join(_tmp, "edit_bad.pptx"), "--replace", *_args)
    assert _rc != 0 and not os.path.exists(os.path.join(_tmp, "edit_bad.pptx")), _o
_rc, _o = _py("edit_deck.py", _mp, _mp, "--replace", "3", "시작", "개시")
assert _rc != 0 and hashlib.sha256(open(_mp, "rb").read()).hexdigest() == _sha0, "원본을 덮음"

# 배치 4: 검토 반영
import re  # noqa: E402
import shutil  # noqa: E402
import edit_deck  # noqa: E402,F401
import gate  # noqa: E402

_mm_dir = os.path.join(_tmp, "mm0")   # 위 Mermaid 시험 덱 보관 (이후 _strict 가 _mp 를 덮음)
os.makedirs(_mm_dir)
_mp0 = os.path.join(_mm_dir, os.path.basename(_mp))
shutil.copy(_mp, _mp0)

# allow 정규식(보통 공백)이 NBSP 로 바뀐 글에도 맞음
_ta = "24 시간 늘 수집"
_rc, _o = _py("check_ban.py", _mini("r", lambda m: text(m.page("허용", lead="요약"), 1, 2, 8, 0.3, _ta), allow=[_ta]))
assert _rc == 0, ("allow 가 NBSP 글에 안 먹음", _o)
_rc, _o = _py("check_ban.py", _mini("s", lambda m: text(m.page("허용", lead="요약"), 1, 2, 8, 0.3, _ta)))
assert _rc == 1 and "늘" in _o, _o

# 출처 줄 (줄 머리 또는 괄호 안). '기준:' 은 아님
for _t, _ok in (("※ 출처: 시험 결과서", True), ("출처：시험 결과서", True), ("(출처: 시험 결과서)", True),
                ("그림 1 (자료: 시험 결과서 2026.05.27)", True), ("데이터 출처: MES", True), ("Source: MES", True),
                ("※ 기준: 본문 장 쪽 번호", False), ("기준: 4차 시험", False), ("자료실: 열람", False),
                ("출처 없음", False), ("시험 결과 출처를 밝힘", False)):
    assert bool(check_ban.SRC_LINE.search(_t)) == _ok, (_t, _ok)


def _chart_note(line):
    def b(m):
        s_ = m.page("차트", lead="요약")
        _chart(s_, 1, 2, 5, 3, "col", ["가", "나"], {"값": [1, 2]}, number_format='0"개"')
        notes(s_, 1, 5.5, 6, [line])
    return b


_rc, _o = _py("check.py", _mini("t", _chart_note("기준: 4차 시험 결과서")))
assert "출처 줄이 없음" in _o, "'기준:' 만 있는 장이 통과함"
_rc, _o = _py("check.py", _mini("u", _chart_note("시험 결과 (자료: 4차 시험 결과서)")))
assert "출처 줄이 없음" not in _o, _o

# 차트 단위: 통화 기호도 단위
for _nf in ("$#,##0", "₩#,##0", "€0.0", "¥0"):
    assert _strict(lambda m, nf=_nf: _chart(m.page("차트", lead="요약"), 1, 2, 5, 3, "col", ["가", "나"], {"값": [1, 2]},
                                         number_format=nf)) == "", _nf


# 일정 및 산출물: 마일스톤 행은 대조에서 제외
def _ms(kind):
    def b(m):
        _gantt(m.page("일정", lead="요약"), 1, 2, 10, ["1월", "2월"], [("설계", 0, 1, "done"), ("현장 검수", 1.5, 1, kind, "검수")])
        m.deliverables({"설계": ["설계서"]})
    return b


assert _strict(_ms("mile")) == ""
assert "'현장 검수'" in _strict(_ms("plan"))

# 간트 막대 글자색: 끝남은 진한 회색, 진행과 예정은 흰색
_gs = _m0.page("간트", lead="요약")
gantt(_gs, 1, 2, 8, ["1월", "2월"], [("가", 0, 1, "done", "완료막대"), ("나", 1, 1.5, "now", "진행막대"),
                                       ("다", 1.5, 2, "plan", "예정막대")])
_gc = {sh_.text_frame.text: str(sh_.text_frame.paragraphs[0].runs[0].font.color.rgb) for sh_ in _gs.shapes
       if sh_.has_text_frame and sh_.text_frame.text.endswith("막대")}
assert _gc == {"완료막대": SLATE9, "진행막대": WHITE, "예정막대": WHITE}, _gc


# 카드 목록: numbered=False 면 번호 원 없이 굵은 제목과 설명
def _cards(numbered):
    s_ = _m0.page("카드", lead="요약")
    c_ = Col(s_, 1, 4, top=2, bottom=6)
    c_.num_cards([("제목 하나", "설명 하나"), ("제목 둘", "설명 둘")], numbered=numbered)
    c_.done()
    return s_


_s1, _s2 = _cards(True), _cards(False)
_ov = lambda s_: [sh_.text_frame.text for sh_ in s_.shapes if sh_.has_text_frame and "Oval" in sh_.name]
assert _ov(_s1) == ["1", "2"] and _ov(_s2) == [], "numbered 가 번호 원을 정하지 못함"
assert any(sh_.has_text_frame and sh_.text_frame.text.startswith("제목 하나")
           and sh_.text_frame.paragraphs[0].runs[0].font.bold for sh_ in _s2.shapes)

# 비교 표: 한 표, 바뀐 행의 개선 칸만 남색
_cc = Col(_m0.page("비교", lead="요약"), 1, 8, top=2, bottom=4)
_cc.compare_table([("기록 방식", "수기 입력", "자동 수집"), ("기록 간격", "1일", "1초"), ("위치", "PC", "서버")], highlight=(1,), title="현행 및 개선")
_cc.done()
_ct2 = _cc[0].table
assert [c.text for c in _ct2.rows[0].cells] == ["항목", "현행", "개선"] and len(_ct2.rows) == 4
_clr = lambda r_, c_: str(_ct2.cell(r_, c_).text_frame.paragraphs[0].runs[0].font.color.rgb)
assert _clr(2, 2) == ACCENT and _clr(2, 1) == BLACK and _clr(1, 2) == BLACK and _clr(3, 2) == BLACK, "개선 칸 강조"
for _bad in (lambda: Col(_m0.page("비교", lead="요약"), 1, 8).compare_table([("가", "나")]),
             lambda: Col(_m0.page("비교", lead="요약"), 1, 8).compare_table([("가", "나", " ")]),
             lambda: Col(_m0.page("비교", lead="요약"), 1, 8).compare_table([("가", "나", "다")], highlight=(3,))):
    try:
        _bad()
        raise AssertionError("막혀야 할 비교 표 입력이 통과함")
    except ValueError:
        pass

# 규칙 목록 이름표: 왼쪽 위 남색 네모 (가운데 제목과 좌우 선은 없음)
_rs = _m0.page("목록", lead="요약")
rule_list(_rs, 1, 2, 4, "이름표", [("doc_id_green", "가")])
_lab = [sh_ for sh_ in _rs.shapes if sh_.has_text_frame and sh_.text_frame.text == "이름표"]
assert len(_lab) == 1 and abs(_lab[0].left / 914400 - 1) < 0.001 and abs(_lab[0].top / 914400 - 2) < 0.001
assert str(_lab[0].fill.fore_color.rgb) == ACCENT and _lab[0].width / 914400 < 4

# 패널 높이 맞춤 시험 (fit, fit=False)
for _fit in (True, False):
    _buf = io.StringIO()
    with contextlib.redirect_stdout(_buf):
        _fs = _m0.page("줄어듦", lead="요약")
        _fa, _fb = _fc = report_panels(_fs, ["가", "나"], widths=(6, 6), fit=_fit)
        _fa.section("제목", ["줄 하나"])
        _fb.section("제목", ["줄 하나", "줄 둘", "줄 셋"])
        for _c in _fc:
            _c.done()
    _hs = [b_.height / 914400 for b_ in _fc.bodies]
    assert abs(_hs[0] - _hs[1]) < 0.001, _hs
    assert (_hs[0] < 2.0) == _fit and (_hs[0] > 4.0) == (not _fit), (_fit, _hs)
    assert ("빔" in _buf.getvalue()) == (not _fit), _buf.getvalue()
    assert not [m_ for m_ in layout_problems(_fs) if "빔" in m_ or "45%" in m_], layout_problems(_fs)
_f = Deck(author="소속 / 점검")   # 줄어든 패널은 엄격 저장을 막지 않음
_f.cover("점검", "2026.11.16")
_f.section("점검")
_fs = _f.page("줄어듦", lead="요약")
_fa, = _fc = report_panels(_fs, ["가"], widths=(12.8,))
_fa.section("제목", ["줄 하나"])
_fa.done()
_f.save(os.path.join(_tmp, "점검_fit_v0.1.0_261116.pptx"))

# 패널 머리와 소제목이 같은 말이면 출력만 (저장은 멈추지 않음)
for _pt, _hit in (("일정 현황 (2026.05.28)", True), ("수집 현황", False)):
    _buf = io.StringIO()
    with contextlib.redirect_stdout(_buf):
        _mini("v", lambda m, t=_pt: report_panels(m.page("일정 현황 (2026.05.28)", lead="요약"), [t], widths=(12.8,)))
    assert ("소제목과 같은 말" in _buf.getvalue()) == _hit, (_pt, _buf.getvalue())

# 날짜 표기: 5종 중 2종 넘게 섞이면 주의 (표지 날짜 YYYY.MM.DD 포함)
assert check_text.date_kinds(["2026.05.29 기준"]) == {"YYYY.MM.DD": "2026.05.29"}
assert check_text.date_kinds(["2026.05 결과서", "95.7 %", "v0.1.0", "1500.5 mm"]) == {"YYYY.MM": "2026.05"}
assert set(check_text.date_kinds(["2026-06-15", "6/15", "6월 15일"])) == {"YYYY-MM-DD", "M/D", "M월 D일"}
_f, _w = check_text.check(_mini("w", lambda m: text(m.page("날짜", lead="요약"), 1, 2, 8, 0.3, "6월 15일 검수")))
assert not any("날짜 표기" in x for x in _w), _w
_f, _w = check_text.check(_mini("x", lambda m: text(m.page("날짜", lead="요약"), 1, 2, 8, 0.3, "6월 15일 및 2026-06-15")))
assert any("날짜 표기 3종" in x for x in _w), _w

# im-not-ai 어휘 사전: 필드가 빠진 항목은 건너뜀, 갱신은 형식이 맞을 때만
_good = {"id": "X-1", "headword": "표제", "tier": "change", "regex": "가나", "change_when": {"보고서": "바꿈"}}
assert check_text.lexicon_entry(_good)[0] == "X-1"
for _bad in ({**_good, "change_when": {}}, {k_: v_ for k_, v_ in _good.items() if k_ != "headword"},
             {**_good, "regex": "(["}, {**_good, "tier": "conditional"}):
    assert check_text.lexicon_entry(_bad) is None, _bad
_lx = os.path.join(_tmp, "lex.json")
open(_lx, "w", encoding="utf8").write(json.dumps({"entries": [_good, {"id": "X-2", "tier": "change", "regex": "다"}]}))
_old_lx, check_text.LEXICON = check_text.LEXICON, _lx
assert [x[0] for x in check_text.load_lexicon()] == ["X-1"], "필드 빠진 항목이 예외를 냄"
check_text.LEXICON = _old_lx
update_imnotai.valid_lexicon(json.dumps({"entries": [_good]}))
for _bad in ({"entries": [_good, {**_good, "change_when": {}}]}, {"entries": [{**_good, "id": ""}]},
             {"entries": [{**_good, "tier": "conditional"}]}):
    try:
        update_imnotai.valid_lexicon(json.dumps(_bad))
        raise AssertionError("형식이 틀린 lexicon 이 통과함")
    except ValueError:
        pass

# 판 번호: plugin.json 먼저, 없으면 package.json
_gh = gate.HERE
_vt = tempfile.mkdtemp(prefix="ixver_")
os.makedirs(os.path.join(_vt, "scripts"))
os.makedirs(os.path.join(_vt, ".claude-plugin"))
gate.HERE = os.path.join(_vt, "scripts")
assert gate._version() == "?"
json.dump({"version": "1.0.0"}, open(os.path.join(_vt, "package.json"), "w"))
assert gate._version() == "1.0.0"
json.dump({"version": "9.9.9"}, open(os.path.join(_vt, ".claude-plugin", "plugin.json"), "w"))
assert gate._version() == "9.9.9"
gate.HERE = _gh
assert gate._version() == _vs[".claude-plugin/plugin.json"]

# 장 순서는 sldIdLst 순서 (파일 번호 순서 아님)
_PNS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def _rezip(src, dst, edit):
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            zo.writestr(info, edit(info.filename, zi.read(info.filename)))


def _reverse_slides(name, data):
    if name != "ppt/presentation.xml":
        return data
    root = etree.fromstring(data)
    lst = root.find(f".//{_PNS}sldIdLst")
    ids = list(lst)
    for e_ in ids:
        lst.remove(e_)
    lst.extend(reversed(ids))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


_rv = os.path.join(_tmp, "reversed.pptx")
_rezip(_mp0, _rv, _reverse_slides)
_hf, _h0 = gate.slide_hashes(_mp0)
assert len(set(_h0)) == len(_h0) and gate.slide_hashes(_rv)[1] == _h0[::-1], "장 해시가 sldIdLst 순서를 따르지 않음"

# 틀 해시: 관계 파일, media, 테마, 렌더 스크립트 포함
with zipfile.ZipFile(_mp0) as _zz:
    _names = _zz.namelist()
for _pre in ("ppt/slideLayouts/_rels/", "ppt/media/", "ppt/presentation.xml", "ppt/theme/"):
    _first = next(n_ for n_ in _names if n_.startswith(_pre))
    _tp = os.path.join(_tmp, "touched.pptx")
    _rezip(_mp0, _tp, lambda name, data, f=_first: data + b" " if name == f else data)
    assert gate.slide_hashes(_tp)[0] != _hf, f"{_first} 가 틀 해시에 안 들어감"
_gt = tempfile.mkdtemp(prefix="ixgate_")
for _f_ in ("render.py", "render.ps1"):
    shutil.copy(os.path.join(_here, _f_), os.path.join(_gt, _f_))
gate.HERE = _gt
_h_a = gate.slide_hashes(_mp0)[0]
open(os.path.join(_gt, "render.ps1"), "ab").write(b" ")
assert gate.slide_hashes(_mp0)[0] != _h_a, "렌더 스크립트가 틀 해시에 안 들어감"
gate.HERE = _gh

# gate --deck: build 없이 검사, 통과 기록 (렌더는 가짜)
_real_run, _cmds = gate.run, []


def _fake_run(cmd, cwd=None, timeout=600):
    if os.path.basename(cmd[2]) != "render.py":
        return _real_run(cmd, cwd, timeout)
    from PIL import Image as _Img
    _cmds.append(cmd)
    sel = cmd[cmd.index("--slides") + 1] if "--slides" in cmd else "all"
    for i_ in (range(1, len(Presentation(cmd[3]).slides) + 1) if sel == "all" else
               [] if sel == "none" else map(int, sel.split(","))):
        _Img.new("RGB", (160, 90), "white").save(os.path.join(cmd[4], f"s{i_}.png"))
    return 0, ""


def _gate_deck(path):
    gate.run = _fake_run
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            rc_ = gate.check_deck(path, [])
    finally:
        gate.run = _real_run
    return rc_, buf.getvalue()


_img = os.path.join(tempfile.gettempdir(), "interx_gate", os.path.splitext(os.path.basename(_mp0))[0])
shutil.rmtree(_img, ignore_errors=True)
_rc, _o = _gate_deck(_mp0)
assert _rc == 0 and "결과: 통과" in _o and "== 1." not in _o, _o[-600:]
assert os.path.exists(gate._record(_mp0)) and gate.verify(_mp0) == 0
assert "- 3쪽 1-1. 흐름: 요약" in open(os.path.join(_img, "titles.md"), encoding="utf8").read()
_gate_deck(_mp0)   # 같은 덱: 그림 재사용, 렌더 대상 없음
assert _cmds[-1][_cmds[-1].index("--slides") + 1] == "none", _cmds[-1]
_sub = os.path.join(_tmp, "sub")
os.makedirs(_sub)
_ed = os.path.join(_sub, os.path.basename(_mp0))   # 같은 이름의 고친 덱: 바뀐 장만 렌더
_rc, _o = _py("edit_deck.py", _mp0, _ed, "--replace", "3", "시작", "개시")
assert _rc == 0 and "gate.py" in _o and "--deck" in _o, _o
_gate_deck(_ed)
assert _cmds[-1][_cmds[-1].index("--slides") + 1] == "3", _cmds[-1]
_bad_ed = os.path.join(_sub, "점검_bad_v0.1.0_261116.pptx")   # 자리표시자가 들어간 덱은 통과 기록이 없어야 함
_s3 = edit_deck.slide_parts(zipfile.ZipFile(_ed))[2]
_rezip(_ed, _bad_ed, lambda n_, d_: d_.replace("개시".encode(), "TODO".encode()) if n_ == _s3 else d_)
_rc, _o = _gate_deck(_bad_ed)
assert _rc == 1 and "결과: 실패" in _o and not os.path.exists(gate._record(_bad_ed)), "금지 표현이 통과함"
_rc, _o = _py("gate.py", "--deck", os.path.join(_tmp, "없는덱.pptx"))
assert _rc == 1 and "덱 파일이 없음" in _o, _o
assert _py("gate.py", "--deck")[0] == 2

# edit_deck: 옛 글이 빈 문자열이면 저장하지 않음
_rc, _o = _py("edit_deck.py", _mp0, os.path.join(_tmp, "edit_empty.pptx"), "--replace", "3", "", "새 글")
assert _rc != 0 and "빈 문자열" in _o and not os.path.exists(os.path.join(_tmp, "edit_empty.pptx")), _o

# Mermaid: 체인의 건너뛰기, 되돌아가기 선은 우회
_CH = ["graph TD|A[가] --> B[나]|B --> C[다]|A --> C", "graph TD|A[가] --> B[나]|B --> C[다]|C --> A",
       "graph TD|A[가] --> B[나]|B --> C[다]|C --> D[라]|D -- 재시도 --> B|A --> C",
       "graph LR|A[가] --> B[나]|B --> C[다]|A --> C", "graph LR|A[가] --> B[나]|B --> C[다]|C --> A",
       "graph LR|A[가] --> B[나]|B --> C[다]|C --> D[라]|D -- 실패 --> B|A --> C"]
for _src in _CH:
    _e = _strict(lambda m, sr=_src: mermaid(m.page("우회", lead="요약"), sr.replace("|", chr(10)), 1.0, 2.0, 11.3, 4.8, title="우회 흐름"))
    assert _e == "", (_src, _e)
    _rc, _o = _py("check.py", _mp)
    assert "장 밖으로" not in _o and "연결선" not in _o, (_src, _o)
    _zm = zipfile.ZipFile(_mp)
    _x5 = _zm.read(edit_deck.slide_parts(_zm)[2]).decode("utf8")
    assert _x5.count("<p:cxnSp>") == _x5.count("<a:stCxn") == _x5.count("<a:endCxn"), "우회선이 도형에 안 붙음"
    assert int(max(map(int, re.findall(r'name="adj1" fmla="val (\d+)"', _x5)))) > 1000000, "우회선 꺾임이 바깥이 아님"
# 한 줄 체인이 아닌 그림의 되돌아가는 선은 못 그리고 이유를 알림
try:
    mermaid(_m0.page("분기", lead="요약"), chr(10).join(["graph TD", "A[가] --> B[나]", "B --> C[다]", "B --> D[라]",
                                                       "C --> E[마]", "D --> E", "E --> B"]), 1.0, 2.0, 11.3, 4.8)
    raise AssertionError("그릴 수 없는 선이 통과함")
except ValueError as _ex:
    assert "E -> B" in str(_ex) and "피해" in str(_ex), _ex
# 노드 글: 따옴표는 빼고 <br> 과 <br/> 는 줄 바꿈
_sh = mermaid(_m0.page("글", lead="요약"), "graph TD" + chr(10) + 'A["수집<br/>시작"] --> B[읽기<br>실패 시 "재시도"]', 1.0, 2.0, 11.3, 4.8, title="글 처리")
assert _sh["A"].text_frame.text == "수집" + chr(10) + "시작" and _sh["B"].text_frame.text == "읽기" + chr(10) + "실패 시 재시도"
assert _sh["A"].height == _sh["B"].height and _sh["A"].height / 914400 >= 0.34
_sq = mermaid(_m0.page("순서", lead="요약"), chr(10).join(["sequenceDiagram", 'participant A as "설비<br/>PLC"',
                                                          "participant B as 수집기", "A->>B: 값 요청<br/>1초마다"]),
              1.0, 2.0, 11.3, 4.8, title="참가자 글")
assert _sq["A"].text_frame.text == "설비" + chr(10) + "PLC"

# 구성도 문법 주의: 선 색 셋(빨강 제외), 아이콘 세트 섞임
_dn = Deck(author="AX Development / 시험 연구원", purpose="문서")
_dn.cover("구성도 문법", "2026.11.16")
_dn.section("구성")
_ds = _dn.page("선과 아이콘", lead="요약")
_b = [box(_ds, 1 + 2.2 * i, 3.0, 1.0, 0.3, f"상자{i}") for i in range(5)]
for _i, _c in enumerate((NAVY, GRAY, TEAL)):
    connect(_ds, _b[_i], "right", _b[_i + 1], "left", color=_c)
assert not any("연결선 색" in m for m in diagram_notes(_ds)), diagram_notes(_ds)   # 셋까지는 통과
connect(_ds, _b[3], "right", _b[4], "left", color="E36C0A")
assert any("연결선 색 4종" in m for m in diagram_notes(_ds)), diagram_notes(_ds)
icon(_ds, "server_rack", 1.0, 5.0, w=0.4)
icon(_ds, "plc", 2.0, 5.0, w=0.4)
assert any("섞임" in m and "plc" in m for m in diagram_notes(_ds)), diagram_notes(_ds)
# 단위 필수 판정: 날짜, 버전, 연도, 첫 열 순번은 숫자로 세지 않음
_ws = len(_reg(_ds).get("warn", []))
table_caption(_ds, 1.0, 1.8, 6.0, "일정", rows=[["순번", "날짜", "판"], ["1", "2026.05.01", "1.2.0"], ["2", "2026", "1.3.0"]])
assert len(_reg(_ds).get("warn", [])) == _ws, _reg(_ds)["warn"][_ws:]
table_caption(_ds, 1.0, 1.8, 6.0, "수량", rows=[["항목", "값"], ["가", "12"], ["나", "3.5"]])
assert "단위 없음" in _reg(_ds)["warn"][-1], _reg(_ds)["warn"][-1]
# 방법론 부품: DAS 구분, 품질 지수 (원본 예시 98.43), 우선순위 4분면, 수준 단계
for _bad in (lambda: Col(_ds, 1, 11).das_table([("X9", "가", "Float", "-", "-", "-", "-", "PLC")]),
             lambda: Col(_ds, 1, 11).quality_table({"완전성": 97}),
             lambda: Col(_ds, 1, 11).priority_table([("가", 6, 1, 1, 1, 1, 1)]),
             lambda: Col(_ds, 1, 11).levels(4)):
    try:
        _bad()
        raise AssertionError("방법론 부품 잘못된 입력이 통과함")
    except ValueError:
        pass
_qc = Col(_ds, 1, 11, top=2.0, bottom=6.5).quality_table({"완전성": 97.38, "유일성": 100, "유효성": 100, "일관성": 100,
                                                           "정확성": 100}, title="품질")
_qc.done()
assert any("98.43" in c.text for sh in _ds.shapes if sh.has_table for r in sh.table.rows for c in r.cells)
_pr = Col._priority([("OEE", 4, 4, 5, 4, 4, 4), ("설비", 5, 4, 3, 2, 3, 2), ("품질", 3, 3, 2, 3, 3, 3),
                     ("검사", 3, 3, 3, 4, 4, 4)], 3.5)
assert [p[-1] for p in _pr] == ["QW", "SI", "LN", "NH"] and _pr[0][2] == 4.3 and _pr[0][4] == 4.0, _pr
# 견본 지표: 그룹 안 도형과 연결선까지 셈
import samples as _smp  # noqa: E402
_grp = _ds.shapes.add_group_shape([_b[0], _b[1]])
assert _smp.kind(_ds) == "구성도" and _smp.metrics(_ds, Inches(13.333))["shapes"] >= 7

# 표 및 그림 제목: 덱 전체 연번 (장 순서, 장 안은 왼쪽 칸부터), 제목과 단위 누락, 한 줄
_NUMS = [["가", "건수"], ["나", "12"], ["다", "34"]]
_LONG = "공정별 처리 건수와 불량 건수를 한 줄로 이어 보는 긴 제목 점검 문구 입니다"


def _cap(m):
    for k_, (kind, a_t, b_t) in enumerate((("table", "왼쪽 표", "오른쪽 표"), ("chart", "왼쪽 그림", "오른쪽 그림"))):
        a_, b_ = cc_ = report_panels(m.page(f"제목 {k_}", lead="요약"), ["왼쪽", "오른쪽"], arrows=False)
        for c_, t_ in ((b_, b_t), (a_, a_t)):   # 오른쪽 칸을 먼저 만들어도 왼쪽 칸이 앞 번호
            if kind == "table":
                c_.table(_NUMS, title=_LONG if t_ == a_t else t_, unit="건")
            else:
                c_.chart("col", ["가", "나"], {"값": [1, 2]}, title=t_, number_format='0"건"')
        for c_ in (b_, a_):
            c_.done()


assert _strict(_cap) == "", _strict(_cap)
_cs = Presentation(os.path.join(_tmp, "점검_strict_v0.1.0_261116.pptx")).slides
_cn = [[sh_ for sh_ in sl_.shapes if sh_.name.startswith("IXCAP_")] for sl_ in list(_cs)[2:4]]
_ct = [[sh_.text_frame.text for sh_ in c_] for c_ in _cn]
assert sorted(_ct[0]) == sorted([f"[표 1. {_LONG}]", "[표 2. 오른쪽 표]"]), _ct   # (a) 장 순서와 왼쪽 칸 먼저
assert sorted(_ct[1]) == ["[그림 1. 왼쪽 그림]", "[그림 2. 오른쪽 그림]"], _ct
assert all(sh_.text_frame.text.startswith("[") for c_ in _cn for sh_ in c_)
assert all(len(sh_.text_frame.paragraphs) == 1 and sh_._element.find(".//" + qn("a:br")) is None
           for c_ in _cn for sh_ in c_), "제목이 한 문단 한 줄이 아님"   # (d)
assert any(sh_.text_frame.text == "(단위: 건)" for sh_ in _cs[2].shapes), "단위 글이 없음"


def _tb(**kw):
    return lambda m: Col(m.page("표", lead="요약"), 1, 6).table(_NUMS, **kw).done()


_e = _strict(_tb(unit="건"))   # (b) 표 제목 없음
assert "표 제목 없음" in _e, _e
assert "단위 없음" not in _e, _e
_e = _strict(_tb(title="숫자 표"))   # (c) 단위 없음
assert "단위 없음" in _e and "표 제목 없음" not in _e, _e
assert _strict(_tb(title="숫자 표", unit="건")) == ""

print("OK", out, "cxn", x.count("<p:cxnSp>"))
