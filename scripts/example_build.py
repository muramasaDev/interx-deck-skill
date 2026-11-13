"""새 덱 출발점 (build.py 로 복사, 내용만 교체)
- 예시 과제: 설비 데이터 수집 시스템 (수치는 형식을 보이기 위한 값)"""
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
OUT = sys.argv[1] if len(sys.argv) > 1 else d.filename("설비 수집 진행 보고", "0.1.0", tempfile.gettempdir())
d.cover("설비 데이터 수집 진행 보고", "2026.05.29")      # 덱 기준일

# 요약 줄(lead)은 그 장의 결론. 본문은 근거만
d.section("시험 결과 및 개선")
s = d.page("수집 성능 (1 ~ 4차 시험)",
           lead="4차 시험에서 검수 기준 2개를 모두 충족했고 1 ~ 4차 결함 22건은 조치 완료")
a, b, c = cols = report_panels(s, ["검수 지표 (4차 시험)", "수집 지연 분포 (4차 시험)", "시험 회차별 발견 결함"])
(a.kpis([("99.8", "%", "수집 성공률"), ("95.7", "%", "0.2초 이내 저장")])   # 핵심 수치는 덱 본문의 1/5 장까지
  .formula(r"S = \frac{N_{저장}}{N_{수집}} \times 100",   # 한 글자 기호, 기호마다 뜻과 단위
           where=["S: 수집 성공률 (%)", "N_{저장}: 데이터 저장소에 저장된 수집 회차 (86,227회)",
                  "N_{수집}: 24시간 동안의 수집 회차 (86,400회)"])
  .num_cards([("검수 기준", "수집 성공률 99.5% 이상, 0.2초 이내 저장 95% 이상"),
              ("미저장 173회", "EdgePC 재시작 중 저장되지 않은 회차 (4차 결함 2건 중 1건, 조치 완료)")])
  .notes(["수집 회차: 1초마다 설비 3대를 한 번 읽는 단위"]))
(b.hist([0.04] * 61200 + [0.15] * 21300 + [0.3] * 3500 + [0.7] * 227, [0, 0.1, 0.2, 0.5, 1.0], unit=" 초",
        name="수집 회차", count_unit="회", h=2.6, highlight=(0, 1))
  .tab_box("지연 정의", ["수집 지연: 저장 시각에서 수집 시각을 뺀 값", "0.2초 이내 저장: 82,500회 / 86,227회 = 95.7%",
                       "0.5초 초과 227회는 네트워크 재연결 구간"]))
(c.chart("bar", ["1차 시험", "2차 시험", "3차 시험", "4차 시험"], {"결함": [6, 9, 5, 2]}, h=2.6,
         number_format='0"건"')
  .tab_box("주요 결함", ["2차 시험: 재연결 뒤 같은 값 중복 저장 4건", "3차 시험: 보관 큐 재전송 순서 오류 3건"])
  .notes(["결함 수: 시험 결과서 기준"]))
for col in cols:
    col.done()

s = d.page("현행 및 개선 비교", lead="6월 검수 뒤 수기 기록을 자동 수집으로 바꾸면 기록 간격이 1일에서 1초로 단축")
a, b = cols = report_panels(s, ["현행 (수기 기록)", "개선 (자동 수집)"], widths=(5.9, 5.9), arrows=True)
rows_a = [["항목", "내용"], ["기록 방식", "작업자가 설비별 PC 파일에 직접 입력"], ["기록 간격", "1일"], ["누락 확인", "월말 대조"],
          ["데이터 위치", "설비별 PC"]]
rows_b = [["항목", "내용"], ["기록 방식", "PLC 값 자동 수집"], ["기록 간격", "1초"], ["누락 확인", "5분 이상 누락 시 메일"],
          ["데이터 위치", "중앙 서버 데이터 저장소"]]
(a.table(rows_a, ratio=[1, 2.2], row_h=0.32)
  .rule_list("현행 문제", [("warning", "기록 값을 다음 날에야 확인"), ("warning_orange", "누락이 월말에야 드러남"),
                         ("warning_navy", "설비마다 파일 위치가 다름")], pitch=0.5)
  .notes(["현행 기준: 2026년 1월 현장 확인"]))
(b.table(rows_b, ratio=[1, 2.2], row_h=0.32)
  .tab_box("6월 검수 후 바뀌는 업무", ["기록 값을 수집 1초 뒤 데이터 저장소에서 조회", "5분 이상 누락되면 담당자가 메일로 확인",
                                  "설비별 파일 대신 중앙 서버 데이터 저장소에서 조회"])
  .num_cards([("적용 범위", "설비 3대, 수집 항목 120개 (3대 합계)"), ("알림 수신", "누락: 생산기술팀 설비 담당자 2명, 현황 보고 5분 이상 끊김: 정보팀 운영 담당자 1명")]))
for col in cols:
    col.done()

d.section("일정 및 방법")
s = d.page("단계별 일정 (2026년)", lead="구현 및 시험을 마치고 현장 적용 중이며 6월 15일 현장 검수 예정")
(a,) = cols = report_panels(s, ["단계별 일정"], widths=(12.8,))
a.gantt(["1월", "2월", "3월", "4월", "5월", "6월"],
        [("요구 정리", 0, 1, "done", "완료"), ("설계", 0.5, 2, "done", "완료"), ("구현", 1.5, 3.5, "done", "완료"),
         ("시험 (1 ~ 4차)", 3, 4.6, "done", "완료"), ("현장 적용", 4.6, 5.6, "now", "진행"),
         ("현장 검수", 5.5, 0, "mile", "6/15 검수")],
        row_h=0.5, label_w=2.4, today=4.95).notes(["기준일 2026.05.29 (빨간 점선)"])
a.done()

s = d.page("진행 방법", lead="수집 항목 정의서를 받아 1초 수집을 구현하고 저장 회차 및 수집 지연으로 검증")
a, b, c = cols = report_panels(s, ["받은 자료 및 현장 확인", "구현", "검증"], arrows=True)   # 순서가 있을 때만 화살표
(a.section("현업에서 받은 자료", ["수집 항목 정의서 (2026.01.09)", "현장 네트워크 구성도 (2026.01.16)"],
           row_h=0.38)
  .section("정의서 내용", ["설비별 PLC 주소", "수집 항목 이름 및 단위 (3대 합계 120개)"], row_h=0.38)
  .tab_box("현장 확인 결과", ["PLC 통신: Ethernet, Modbus TCP", "EdgePC 와 중앙 서버 사이 방화벽 개방 (2026.02.03)"])
  .notes(["웹 화면의 항목 이름은 수집 항목 정의서 명칭과 같음"]))
(b.section("수집", ["중앙 서버가 수집 계획 (PLC 주소, 항목 목록) 전달", "PLC 값을 1초마다 읽고 수집 시각 기록"],
           row_h=0.38)
  .section("전송", ["중앙 서버로 전송 (사본은 EdgePC CSV 에 7일)", "실패분은 보관 후 다시 전송"], row_h=0.38)
  .section("현황", ["수집 프로그램 현황을 1분마다 보고"], row_h=0.38)
  .section("알림", ["5분 이상 누락되면 설비 담당자에게 메일", "현황 보고가 5분 이상 끊기면 운영 담당자에게 메일"], row_h=0.38))
(c.section("시험", ["24시간 연속 수집", "수집 회차 대비 저장 회차 대조", "수집 지연 (저장 시각 - 수집 시각) 측정"], row_h=0.38)
  .section("현장 검수 (6월 15일)", ["현업 담당자와 설비 화면 값 및 웹 화면 값 대조", "누락 알림 동작 확인"], row_h=0.38)
  .tab_box("검수 기준", ["수집 성공률 99.5% 이상", "0.2초 이내 저장 95% 이상"])
  .section("기록", ["시험 결과서 및 검수 확인서 작성"], row_h=0.38))
for col in cols:
    col.done()

d.section("시스템 구성")
s = d.page("수집 구성 (설비 3대 기준)", doc=True,
           lead="설비 3대의 PLC 값을 EdgePC 가 1초마다 읽어 중앙 서버 데이터 저장소로 전송")
g = DocGrid(s, right=None, rows=7, bottom=6.3)
g.source("source", [("plc", "설비 1"), ("plc", "설비 2"), ("plc", "설비 3")], title="설비 PLC")
plcs = g.icons
e = g.group("edge", "EdgePC", kind="edge", sub="설비 옆", header="수집 프로그램", cols=2, area=False)
s1 = e.sub(0, 0, 6, header="수집 및 보고")
cache, core = s1.box(1, "수집 계획 보관"), s1.box(3, "수집 코어")
report, alarm = s1.box(5, "현황 보고"), s1.box(6, "알림 메일 발송")
s2 = e.sub(1, 0, 4, header="저장 및 전송")
csv, send, spool = s2.box(2, "CSV 기록"), s2.box(3, "송신"), s2.box(4, "보관 큐")
p = g.group("platform", "중앙 서버", header="관리 서버", cols=2, area=False)
s3 = p.sub(0, 0, 6, header="수집 관리")
plan, status = s3.box(1, "수집 계획 제공"), s3.box(5, "수집 프로그램 현황")
srv_alarm = s3.box(6, "보고 끊김 감지")
s4 = p.sub(1, 0, 3, header="저장 및 화면")
db_, web = s4.box(1, "설정 DB"), s4.box(3, "웹 화면")
lake = g.bar("platform", "데이터 저장소")
mail = g.bar("edge", "메일 서버", fill=GRAY)
link(s, plan, cache, color=TEAL, text="HTTPS", step=(1, "계획 받기"))
fan(s, core, plcs, color=ACCENT, step=(2, "PLC 읽기"))
link(s, core, csv, color=ACCENT, step=(3, "기록"))
link(s, core, send, color=ACCENT)
link(s, send, lake, color=ACCENT, text="HTTPS", step=(4, "전송"))
link(s, send, spool, color=RED, dash="sysDot", step=(5, "실패분 보관"), head="both")   # 보관 뒤 다시 전송
link(s, report, status, color=GRAY, step=(6, "현황 보고"))
link(s, alarm, mail, color=GRAY, step=(7, "누락 알림 메일"))
link(s, srv_alarm, mail, color=GRAY, step=(8, "보고 끊김 메일"))
link(s, plan, db_, color=GRAY, head="both")
link(s, web, db_, color=GRAY, head="both")
note_by(s, web, "수집 계획 및 수집 프로그램 현황 조회")   # 설명은 가리키는 상자 옆

d.save(OUT)
print(OUT)
