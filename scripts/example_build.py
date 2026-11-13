"""새 덱 출발점 (build.py 로 복사, 내용만 교체)"""
import os
import sys
import tempfile

def _engine():
    """엔진 위치: INTERX_DECK_ROOT, skills, plugins 순"""
    import glob
    cfg = os.environ.get("CLAUDE_CONFIG_DIR") or os.path.join(os.path.expanduser("~"), ".claude")
    roots = ([os.environ["INTERX_DECK_ROOT"]] if os.environ.get("INTERX_DECK_ROOT") else []) + \
        [os.path.join(cfg, "skills", "interx-deck")] + \
        sorted(glob.glob(os.path.join(cfg, "plugins", "cache", "*", "interx-deck", "*")),
               key=os.path.getmtime, reverse=True)
    for r in roots:
        if os.path.exists(os.path.join(r, "scripts", "ixdeck.py")):
            return os.path.join(r, "scripts")
    sys.exit("interx-deck 엔진을 찾지 못함: 플러그인을 설치하거나 INTERX_DECK_ROOT 를 준다")


sys.path.insert(0, _engine())
from ixdeck import *  # noqa: E402,F403

d = Deck(author=profile_author() or "소속 / 이름 직급")  # 개인 설정 값. 덱마다 사용자에게 확인
OUT = sys.argv[1] if len(sys.argv) > 1 else d.filename("INTERX덱예시", "0.1.0", tempfile.gettempdir())
d.cover("예시 과제 결과 보고", "2026.11.05")      # 덱 기준일

d.section("과제 결과")
s = d.page("핵심 수치 (예시 값)", lead="예시 수치 셋 및 구간별 건수를 한 장에 둠")
top = banner(s, "기준 안 항목 15개, 검토 결함 22건 모두 고침", tags=["1단계", "2단계", "3단계"], current=1)
a, b, c = cols = report_panels(s, ["핵심 수치", "구간별 건수", "검토 결함"], top=top,
                               footnote="* 예시 값: 형식을 보이기 위한 수")
(a.kpis([("15", "개", "기준 안 항목"), ("22", "건", "검토 결함")])   # 핵심 수치는 덱 본문의 1/5 장까지
  .num_cards([("기준 안 항목", "전체 22개 중 15개가 0.2 미만")])
  .formula(r"R = \frac{N_{기준}}{N} \times 100",   # 한 글자 기호, 기호마다 설명과 단위
           where=["R: 기준 안 비율 (%)", "N_{기준}: 거리 0.2 in 미만 항목 (개)", "N: 전체 항목 (개)"]))
(b.chart("col", ["0.1 미만", "0.1 ~ 0.2", "0.2 ~ 0.5", "0.5 이상"], {"건수": [7, 8, 3, 4]}, h=2.6,
         highlight=(0, 1), number_format='0"개"')
  .notes(["15개가 0.2 미만", "강조 막대는 기준 안 구간"]))
(c.chart("bar", ["검토 1", "검토 2", "검토 3", "검토 4"], {"결함": [6, 9, 5, 2]}, h=2.6, number_format='0"건"')
  .notes(["모두 22건, 모두 고침"]))
for col in cols:
    col.done()

s = d.page("현행 및 개선 비교 (예시)", lead="같은 항목을 두 표로 나란히 두고 바뀐 줄만 강조함")
a, b = cols = report_panels(s, ["현행", "개선"], widths=(5.9, 5.9))
rows_a = [["항목", "값"], ["양식", "개인 양식"], ["글꼴", "여러 글꼴"], ["글자 크기", "9.5 ~ 12pt"],
          ["연결선", "도형에 안 붙음"], ["번호 설명", "아래 범례 줄"]]
rows_b = [["항목", "값"], ["양식", "회사 양식"], ["글꼴", "Noto Sans KR"], ["글자 크기", "6.35 ~ 8.16pt"],
          ["연결선", "도형에 붙음"], ["번호 설명", "선 옆"]]
(a.table(rows_a, ratio=[1.2, 2], row_h=0.32)
  .rule_list("현행 문제", [("warning", "보는 PC 마다 다르게 보임"), ("doc_check", "선을 옮기면 도형과 떨어짐"),
                         ("checklist_green", "장마다 제목 꼴이 다름")], pitch=0.42)
  .notes(["예시 값"]))
(b.table(rows_b, ratio=[1.2, 2], row_h=0.32, highlight=(0, 1, 2, 3, 4))
  .icon_cards([("monitor_gear", "같은 화면"), ("gears_cycle", "자동 배치"), ("chart_distribution", "검사 통과")])
  .tab_box("바뀐 점", ["회사 양식 하나", "저장할 때 겹침 검사"])
  .notes(["강조 줄은 바뀐 항목"]))
for col in cols:
    col.done()

d.section("일정 및 방법")
s = d.page("단계별 일정 (예시)", lead="막대 색은 끝남 회색, 진행 남색, 예정 남회색이고 마름모는 마일스톤임")
(a,) = cols = report_panels(s, ["단계별 일정"], widths=(12.8,))
a.chevrons(["1단계 요구 및 설계", "2단계 구현 및 시험", "3단계 현장 적용"], current=1)
a.gantt(["1월", "2월", "3월", "4월", "5월", "6월"],
        [("요구 정리", 0, 1, "done", "끝남"), ("설계", 0.5, 2, "done", "끝남"), ("구현", 1.5, 4, "now", "진행"),
         ("시험", 3.5, 5, "plan"), ("현장 적용", 4.5, 6, "plan"), ("검수", 5.5, 0, "mile", "검수")],
        row_h=0.5, label_w=2.4, today=2.6).notes(["예시 일정"])
a.done()

s = d.page("진행 방법 (예시)", lead="단계마다 입력, 방법, 확인을 목록으로 둠")
a, b, c = cols = report_panels(s, ["입력", "방법", "확인"])
(a.section("받는 자료", ["장 구성", "장마다 할 말 한 줄", "근거 자료 경로"], row_h=0.38)
  .section("먼저 묻는 것", ["작성자 정보"], row_h=0.38, highlight=(0,))
  .rule_list("자료 기준", [("doc_check", "원천 명칭 그대로"), ("checklist_green", "확인한 사실만")])
  .notes(["예시 목록"]))
(b.section("배치", ["패널, 열, 여백", "묶음 안 상자 자리", "줄 높이"], row_h=0.32)
  .section("선과 번호", ["변 및 꺾임", "선 다발", "번호 자리"], row_h=0.32)
  .section("글꼴 및 크기", ["Noto Sans KR 하나", "기본 양식 글자 크기"], row_h=0.32)
  .notes(["모두 자동"]))
(c.section("검사", ["겹침 및 넘침", "제목 번호", "글꼴 및 크기"], row_h=0.38)
  .section("넘기기 전", ["렌더 그림 확인", "사실 검토 1건"], row_h=0.38, highlight=(1,))
  .section("고칠 때", ["build.py 를 고쳐 다시 실행", "PowerPoint 로 연 파일은 넘기지 않음"], row_h=0.38)
  .notes(["gate.py 한 줄"]))
for col in cols:
    col.done()

d.section("시스템 구성")
s = d.page("수집 구성 (예시)", doc=True)
g = DocGrid(s, right=None, rows=7, bottom=6.3)
g.source("source", [("plc", "설비 1"), ("plc", "설비 2"), ("plc", "설비 3")], title="설비 PLC")
plcs = g.icons
e = g.group("edge", "EdgePC", kind="edge", sub="설비 옆", header="수집 프로그램", cols=2, area=False)
s1 = e.sub(0, 0, 6, header="수집 및 보고")
cache, core = s1.box(1, "수집 계획 캐시"), s1.box(3, "수집 코어")
report, alarm = s1.box(5, "현황 보고"), s1.box(6, "알림 메일 보내기")
s2 = e.sub(1, 0, 4, header="저장 및 전송")
csv, send, spool = s2.box(2, "CSV 기록"), s2.box(3, "송신"), s2.box(4, "보관 큐")
p = g.group("platform", "중앙 서버", header="관리 서버", cols=2, area=False)
s3 = p.sub(0, 0, 6, header="수집 관리")
plan, status = s3.box(1, "수집 계획 제공"), s3.box(5, "수집기 현황")
srv_alarm = s3.box(6, "서버 알림")
s4 = p.sub(1, 0, 3, header="저장 및 화면")
db_, web = s4.box(1, "DB"), s4.box(3, "웹 화면")
lake = g.bar("platform", "데이터 저장소")
mail = g.bar("edge", "메일 서버", fill=GRAY)
link(s, plan, cache, color=TEAL, text="HTTPS", step=(1, "계획 받기"))
fan(s, core, plcs, color=ACCENT, step=(2, "PLC 읽기"))
link(s, core, csv, color=ACCENT, step=(3, "기록"))
link(s, core, send, color=ACCENT)
link(s, send, lake, color=ACCENT, text="HTTP", step=(4, "전송"))
link(s, send, spool, color=RED, dash="sysDot", step=(5, "실패분 보관"))
link(s, report, status, color=GRAY, step=(6, "현황 보고"))
link(s, alarm, mail, color=GRAY, step=(7, "알림 메일"))
link(s, srv_alarm, mail, color=GRAY, step=(8, "서버 알림 메일"))
link(s, plan, db_, color=GRAY, head="both")
link(s, web, db_, color=GRAY, head="both")
note_by(s, web, "설명은 설명하는 상자 옆에", ["오른쪽 여백에 설명 열을 따로 두지 않음"])

d.save(OUT)
print(OUT)
