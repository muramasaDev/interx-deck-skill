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
out = os.path.join(tempfile.mkdtemp(), "diag.pptx")
d.save(out)
_, warns = check(out)
for key in ("표지 제목", "영문 뒤 한글", "같은 출처 문구"):
    assert any(key in w for w in warns), (key, warns[:5])
print("selfcheck_diag OK")
