"""설계 문서 부품(ixdoc.py) 자체 시험: 장이 만들어지고, 잘못된 입력은 막힘"""
import os
import re
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))
from PIL import Image  # noqa: E402
from ixdeck import *  # noqa: E402,F403

tmp = tempfile.mkdtemp()
png = os.path.join(tmp, "ui.png")
Image.new("RGB", (1600, 900), "#F4F6F8").save(png)
d = Deck(author="AX Development / 시험 연구원", purpose="문서")
d.cover("설계 부품 시험", "2026.11.18")
d.section("기능 상세")
IO = [("항목 A", "입력", "문자열", "필수", "설명")]
ERR = [("400", "조건", "이미 등록된 이름입니다")]
LIM = [("항목", "5,000", "개", "-")]
d.func_page("기능 하나 (M-001)", "요약 문장", flow=["첫 단계", "둘째 단계"], limits=LIM, io=IO, errors=ERR, refs={"기능": "F-001"}, source="시험 자료 (2026.11.18)")
d.func_page("기능 둘 (M-002)", "요약 문장", flow=["첫 단계", "둘째 단계", "셋째 단계"], io=IO)
d.section("화면 상세")
assert d.allow == [re.escape("이미 등록된 이름입니다")], d.allow   # 오류 표의 화면 문구는 문장 검사 예외로 자동 등록
d.screen_page("화면 하나 (SCR-01)", "요약 문장", image=png, marks=[(0.1, 0.1, 0.5, 0.3, 1)],
              elements=[("1", "요소", "동작", "검증")])


def bad(fn, msg):
    try:
        fn()
    except ValueError:
        return
    raise AssertionError(msg)


c = Col(Deck(author="AX Development / 시험 연구원", purpose="문서").report_slide("1. 시험", "1-9. 막힘"), 1.0, 5.0)   # 저장하지 않는 덱
bad(lambda: c.field_table([("항목", "양방향", "문자열", "필수", "설명")]), "방향 검사")
bad(lambda: c.error_table([("오류", "조건", "문구")]), "코드 꼴 검사")
bad(lambda: c.element_table([("가", "요소", "동작", "검증")]), "번호 숫자 검사")
bad(lambda: c.flow(["한 단계뿐"]), "흐름 두 단계 이상")
bad(lambda: c.limit_table([("항목", "", "개", "-")]), "빈 칸 검사")
bad(lambda: d.screen_page("번호 어긋남", "요약", image=png, marks=[(0, 0, 1, 1, 2)], elements=[("1", "요소", "동작", "검증")]),
    "캡처와 표 번호")
bad(lambda: d.screen_page("범위 거꾸로", "요약", image=png, marks=[(0.5, 0.5, 0.1, 0.1, 1)], elements=[("1", "요소", "동작", "검증")]), "marks 범위")
bad(lambda: d.screen_page("번호 겹침", "요약", image=png, marks=[(0, 0, 0.5, 0.5, 1), (0.5, 0.5, 1, 1, 1)], elements=[("1", "요소", "동작", "검증")]), "marks 번호 겹침")
bad(lambda: c.element_table([("1", "요소", None, "검증")]), "None 칸")
d.func_page("흐름만 (M-009)", "요약 문장", flow=["첫 단계", "둘째 단계"])   # 한 구역이어도 패널이 전체 폭
_txt = " ".join(sh.text_frame.text for sl in d.prs.slides for sh in sl.shapes if sh.has_text_frame)
assert "※ 출처: 시험 자료 (2026.11.18)" in _txt, "source 줄"
out = os.path.join(tmp, "doc.pptx")
d.save(out)
assert os.path.getsize(out) > 20000
# 패널 제목과 표 제목이 같으면 주의
from check_text import check  # noqa: E402
_d2 = Deck(author="AX Development / 시험 연구원", purpose="문서")
_d2.cover("제품 설계서 시험", "2026.11.18")
_d2.section("시험")
_s2 = _d2.page("중복 제목", lead="요약 문장")
(_c2,) = report_panels(_s2, ["입출력 항목"], fit=True)
_c2.field_table(IO, title="입출력 항목")
_c2.done()
_p2 = os.path.join(tmp, "dup.pptx")
_d2.save(_p2)
assert any("두 번 나옴" in w for w in check(_p2)[1]), check(_p2)[1]
print("selfcheck_doc OK")
