"""구성도 역할 부품(ixdiag.py)과 문장 검사 추가분 자체 시험"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))
from check_text import check  # noqa: E402
from ixdeck import *  # noqa: E402,F403

d = Deck(author="AX Development / 시험 연구원", purpose="문서")
d.cover("기능 상세 설계서", "2026.11.18")   # 제품 이름 없는 표지: 주의
d.section("AAS Server기능 상세")   # 영문 뒤 한글 붙임: 주의
for k in range(6):
    s = d.page(f"역할 부품 {k + 1}", lead="요약 문장", doc=True, figure="시험 구성도", notes=None)
    frame(s, 0.6, 2.0, 3.0, 2.0, "묶음", kind="platform")
    for role in ROLE:
        role_box(s, 0.8 + 0.7 * list(ROLE).index(role), 2.6, 0.6, 0.4, "칩", role=role, sub="(sub)")
    block_arrow(s, 4.0, 2.4, 2.0, 0.4, label="큰 흐름")
    pill(s, 4.5, 3.0, "MQTT")
    abbr_legend(s, 0.6, 4.4, 6.0, bullets=["요약 한 줄"], abbrs=[("MQTT", "Message Queuing Telemetry Transport")])
    text(s, 0.6, 5.2, 6.0, 0.25, "※ 출처: 같은 문구", size=7.26)   # 여섯 장 모두 같은 출처
try:
    role_box(s, 0, 0, 1, 1, "x", role="없는역할")
    raise AssertionError("role 검사")
except ValueError:
    pass
try:
    abbr_legend(s, 0, 0, 5)
    raise AssertionError("빈 범례 검사")
except ValueError:
    pass
import subprocess  # noqa: E402
_here = os.path.dirname(os.path.abspath(__file__))
assert subprocess.run([sys.executable, "-P", "-c", "import sys; sys.path.insert(0, %r); import ixdiag, ixdeck; assert hasattr(ixdeck, 'role_box')" % _here]).returncode == 0, "ixdiag 먼저 import"
s2 = d.page("단계 부품", lead="요약 문장", doc=True, figure="시험 단계")
cols = step_header(s2, 0.6, 2.0, 8.0, ["가", "나", "다"])
assert len(cols) == 3 and cols[1][0] > cols[0][0] + cols[0][1]
assert summary_bar(s2, cols, ["a", "b", "c"], 4.0) > 4.0
speech_bubble(s2, 1.0, 5.2, 4.0, 0.5, "말풍선 시험", tail="right")
_v = block_arrow(s2, 10.0, 3.0, 0.4, 2.0, direction="down")   # 세로: (x, y, w, h) 가 보이는 크기
assert abs(_v.rotation - 90) < 1e-6 and abs(_v.height / 914400 - 0.4) < 1e-6 and abs(_v.width / 914400 - 2.0) < 1e-6, (_v.width, _v.height)
for fn in (lambda: step_header(s2, 0, 0, 5, ["하나"]), lambda: summary_bar(s2, cols, ["a"], 4.0),
           lambda: speech_bubble(s2, 0, 0, 1, 1, "x", tail="위")):
    try:
        fn()
        raise AssertionError("단계 부품 검사")
    except ValueError:
        pass
s3 = d.page("헤드라인과 배경", lead="요약 문장", doc=True, figure="시험 헤드라인")
figure_caption(s3, "시험 헤드라인", 9.5, 6.55, 3.0)   # 같은 제목이 한 장에 두 번: 문장 검사가 주의
_y = headline(s3, 0.6, 1.95, 8.0, [[("키워드만 ", False), ("색을 입힘", True)], [("둘째 줄", False)]])
assert _y > 1.95
_cols = step_header(s3, 0.6, _y + 0.1, 8.0, ["가", "나"])
backdrop(s3, _cols, _y + 0.7, 2.0)
try:
    headline(s3, 0, 0, 5, [])
    raise AssertionError("빈 헤드라인")
except ValueError:
    pass
# 기능 색, 흐름 역할, 선 빠짐 주의
from ixdeck import _reg, _set_flow, diagram_notes  # noqa: E402
s4 = d.page("흐름 역할", lead="요약 문장", doc=True, figure="시험 흐름")
_a = db(s4, 1.0, 2.2, 1.5, 0.3, "처리 A")
_b = db(s4, 4.0, 2.2, 1.5, 0.3, "처리 B")
_c = db(s4, 7.0, 2.2, 1.5, 0.3, "도착 C")
for _x, _n, _r in ((_a, "처리 A", "process"), (_b, "처리 B", "process"), (_c, "도착 C", "sink")):
    _set_flow(s4, _x, _n, _r)
link(s4, _a, _b)
_msgs = diagram_notes(s4)
assert any("'처리 A' 에 들어오는 선이 없음" in m for m in _msgs) and any("'처리 B' 에 나가는 선이 없음" in m for m in _msgs), _msgs
assert any("'도착 C' 에 들어오는 선이 없음" in m for m in _msgs), _msgs
link(s4, _b, _c)   # 처리 B 가 이제 들어옴과 나감을 가짐
assert not any("'처리 B'" in m for m in diagram_notes(s4)), diagram_notes(s4)
assert any("'처리 A' 에 들어오는 선이 없음" in m for m in diagram_notes(s4))   # 나가는 선만 있으면 계속 주의 (방향을 셈)
for _bad_kw in ({"flow": "처리"}, ):
    try:
        _set_flow(s4, _a, "x", _bad_kw["flow"])
        raise AssertionError("flow 검사")
    except ValueError:
        pass
legend(s4, 1.0, 4.0, [("1F4E8C", "수집 기능"), ("299288", "전송 기능")], kind="square")
assert set(TONES) == {"collect", "store", "send", "manage", "report"}
g4 = DocGrid(s4, right=None, rows=3)
try:
    g4.group("platform", "서버", tone="없는색")
    raise AssertionError("tone 검사")
except ValueError:
    pass
s5 = d.page("격자 흐름 역할", lead="요약 문장", doc=True, figure="시험 격자")
g5 = DocGrid(s5, right=None, rows=4, bottom=6.0)
grp = g5.group("platform", "서버", header="프로그램", cols=1, tone="collect")
sb = grp.sub(0, 1, 3, header="묶음", tone="send")
bx1, bx2 = sb.box(2, "상자 A", flow="process"), sb.box(3, "상자 B")
link(s5, bx1, bx2)
assert _reg(s5)["flow_roles"][bx1.shape_id] == ("process", "상자 A")
try:
    grp.sub(0, 1, 3, tone="없는색")
    raise AssertionError("sub tone 검사")
except ValueError:
    pass
out = os.path.join(tempfile.mkdtemp(), "diag.pptx")
d.save(out)
_, warns = check(out)
for key in ("표지 제목", "영문 뒤 한글", "같은 출처 문구", "한 장에 두 번"):
    assert any(key in w for w in warns), (key, warns[:5])

# 목차: 두 열일 때 같은 줄의 번호가 같은 높이 (절마다 쪽 수가 달라도)
_dt = Deck(author="AX Development / 시험 연구원", purpose="문서")
_dt.cover("Sample System 설계서", "2026.11.18")
for _i, _n in enumerate((3, 2, 8, 8, 3, 3, 6), 1):
    _dt.section(f"절 {_i}")
    for _k in range(_n):
        _dt.page(f"항목 {_i}-{_k + 1}", lead="요약 문장")
_pt = os.path.join(tempfile.mkdtemp(), "toc.pptx")
_dt.save(_pt)
from pptx import Presentation  # noqa: E402
_tops = {sh.text_frame.text: sh.top for sh in Presentation(_pt).slides[1].shapes if sh.has_text_frame and sh.text_frame.text in ("01", "02", "03", "05", "06", "07")}
assert _tops["01"] == _tops["05"] and _tops["02"] == _tops["06"] and _tops["03"] == _tops["07"], _tops

# 글꼴 검사: 기준 글꼴이 아니면 check.py 가 잡음 (정규식에 백스페이스 문자가 들어 있어 한동안 못 잡았음)
import re  # noqa: E402
import zipfile  # noqa: E402
_bad_font = os.path.join(tempfile.mkdtemp(), "bad_font.pptx")
_done = False
with zipfile.ZipFile(out) as zi, zipfile.ZipFile(_bad_font, "w", zipfile.ZIP_DEFLATED) as zo:
    for _i in zi.infolist():
        _d = zi.read(_i.filename)
        if not _done and re.match(r"ppt/slides/slide\d+\.xml$", _i.filename) and b'typeface="Pretendard"' in _d:
            _d, _done = _d.replace(b'typeface="Pretendard"', b'typeface="Arial"', 1), True
        zo.writestr(_i, _d)
_run = lambda f: subprocess.run([sys.executable, "-P", os.path.join(_here, "check.py"), f], capture_output=True,
                                env={**os.environ, "PYTHONIOENCODING": "utf-8"}).stdout.decode("utf8", "replace")
assert "Pretendard 아닌 글꼴" in _run(_bad_font) and "Pretendard 아닌 글꼴" not in _run(out), "글꼴 검사"

# 새 구성도 부품: 계층 띠, 구역, 사양 표, 노드, 등각, 모호한 이름, 선 위 번호
from ixdeck import layer_bands, bus_line, zone_box, tier_spec, node, iso_plate, iso_box, abbr_legend, role_box  # noqa: E402
_dn = Deck(author="AX Development / 시험 연구원", purpose="문서")
_dn.cover("Sample System 설계서", "2026.11.18")
_dn.section("구성")
sn = _dn.page("부품", lead="요약 문장", doc=True, figure="시험 부품")
_b = layer_bands(sn, 0.75, 2.3, 11.8, [("계층 A", 0.9, ["부제"]), ("계층 B", 0.9, None)])
assert len(_b["rows"]) == 2 and _b["x1"] < 12.55
try:
    layer_bands(sn, 0.75, 2.3, 11.8, [("하나", 1.0, None)])
    raise AssertionError("계층 하나")
except ValueError:
    pass
try:
    zone_box(sn, 1, 1, 1, 1, "x", kind="없는종류")
    raise AssertionError("zone kind")
except ValueError:
    pass
try:
    iso_plate(sn, 1, 1, 1, 1, "x", tone="없는색")
    raise AssertionError("iso tone")
except ValueError:
    pass
_a = role_box(sn, 2.0, 2.45, 1.8, 0.4, "수집 계획 제공", role="base", flow="process")
_c = role_box(sn, 6.0, 2.45, 1.8, 0.4, "수집 항목 및 주기 제공", role="base", flow="sink")
link(sn, _a, _c, color=GRAY, step=(1, "전달"), cond="실패 시")   # 조건 글이 붙은 번호도 저장됨
assert any("'수집 계획 제공' 이 모호함" in m for m in diagram_notes(sn)), diagram_notes(sn)
assert not any("'수집 항목 및 주기 제공' 이 모호함" in m for m in diagram_notes(sn))
abbr_legend(sn, 0.75, 5.0, 11.8, abbrs=[("수집 계획", "수집 항목과 주기를 적은 설정")])
assert not any("모호" in m for m in diagram_notes(sn)), diagram_notes(sn)   # 풀이하면 주의에서 빠짐
print("selfcheck_diag OK")
