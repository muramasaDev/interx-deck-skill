"""새 덱 출발점 (build.py 로 복사, 내용만 교체)
- 예시 과제: 설비 데이터 수집 시스템 (수치는 형식을 보이기 위한 값)"""
import os
import sys
import tempfile

def _engine():
    """엔진 위치: INTERX_DECK_ROOT, skills, plugins 순 (옛 이름 interx-deck 도 찾음)"""
    import glob
    cfg = os.environ.get("CLAUDE_CONFIG_DIR") or os.path.join(os.path.expanduser("~"), ".claude")
    roots = ([os.environ["INTERX_DECK_ROOT"]] if os.environ.get("INTERX_DECK_ROOT") else []) + \
        [os.path.join(cfg, "skills", n) for n in ("ixdeck", "interx-deck")] + \
        sorted((g for n in ("ixdeck", "interx-deck") for g in glob.glob(os.path.join(cfg, "plugins", "cache", "*", n, "*"))),
               key=os.path.getmtime, reverse=True)
    for r in roots:
        if os.path.exists(os.path.join(r, "scripts", "ixdeck.py")):
            return os.path.join(r, "scripts")
    sys.exit("ixdeck 엔진을 찾지 못함: 플러그인을 설치하거나 INTERX_DECK_ROOT 를 준다")


sys.path.insert(0, _engine())
from ixdeck import *  # noqa: E402,F403

# 쓰임새 사내보고: 첫 본문 장 소제목에 '요약'. 발표 15분이면 본문 10장까지
d = Deck(author=profile_author() or "소속 / 이름 직급", purpose="사내보고", minutes=15)   # 작성자는 덱마다 사용자에게 확인
OUT = sys.argv[1] if len(sys.argv) > 1 else d.filename("설비 수집 진행 보고", "0.1.0", tempfile.gettempdir())
d.cover("설비 데이터 수집 진행 보고", "2026.05.29")      # 덱 기준일. 날짜는 YYYY.MM.DD 하나로

# 요약 줄(lead)은 그 장의 결론. 본문은 근거만
d.section("요약")
s = d.page("진행 요약 (2026.05.29 기준)",
           lead="4차 시험에서 검수 지표 4개 중 3개 목표 충족, 누락 확인 1개는 현장 검수 (2026.06.15) 에서 확인",
           notes="목표 충족 3개와 남은 현장 검수 1개를 먼저 말하고, 다음 장에서 현행 및 개선으로 이어감")
a, b = cols = report_panels(s, ["진행 단계 및 시험", "검수 지표 판정"], widths=(6.3, 6.1))
(a.chevrons(["요구 정리", "설계", "구현", "시험", "현장 적용", "현장 검수"], current=4)
  .num_cards([("1 ~ 4차 시험 완료", "24시간 연속 수집 시험 4회, 설비 3대"),
              ("결함 22건 조치", "1 ~ 4차 시험 결함 전부, 재시험 완료 (2026.05.28)"),
              ("현장 적용 중", "설비 3대, 수집 항목 120개")], numbered=False)
  .tab_box("수집 회차", ["1초마다 설비 3대를 한 번 읽는 단위", "저장 회차: 수집 회차 중 데이터 저장소에 저장된 회차",
                                 "시험은 24시간 연속 수집, 재시험도 같은 조건"]))
(b.section("충족 (4차 시험 및 재시험)", ["수집 성공률: 4차 99.8%, 재시험 99.9%", "0.2초 이내 저장: 4차 95.7%, 재시험 97.3%",
                                     "기록 간격: 4차 시험 수집 회차 1초 간격"], row_h=0.4)
  .section("확인 예정 (현장 검수)", ["누락 확인"], row_h=0.4)
  .tab_box("현장 검수 (2026.06.15)", ["현업 담당자와 설비 화면 값 및 웹 화면 값 대조", "누락 알림 동작 확인"])
  .notes(["출처: 4차 시험 결과서 (2026.05.27), 재시험 결과서 (2026.05.28)"]))
for col in cols:
    col.done()

d.section("현행 및 개선")
s = d.page("현행 및 개선 비교",
           lead="자동 수집으로 기록 간격을 1일에서 1초로, 누락 확인을 월말 대조에서 누락 알림 메일로 개선")
a, b = cols = report_panels(s, ["항목별 비교", "개선 후 업무"], widths=(7.9, 4.6))
(a.compare_table([("기록 방식", "작업자가 설비별 PC 파일에 입력", "PLC 값 자동 수집"),
                  ("데이터 위치", "설비별 PC", "중앙 서버 데이터 저장소"),
                  ("현황 확인", "현장 방문", "웹 화면에서 조회"),
                  ("수집 항목 변경", "설비별 PC 에서 각각 수정", "중앙 서버 수집 계획에서 한 번에 수정"),
                  ("기록 간격", "1일", "1초"),
                  ("누락 확인", "월말 대조", "5분 이상 누락 시 메일")],
                 row_h=0.36, title="기록 방식 현행 및 개선")
  .tab_box("현행 문제", ["기록 값을 다음 날에야 확인", "누락이 월말 대조에서야 드러남", "설비마다 파일 위치가 달라 조회가 번거로움"])
  .notes(["출처: 현장 확인 (2026.01.16)"]))
(b.tab_box("바뀌는 업무", ["누락 메일을 받은 담당자가 설비와 EdgePC 상태 확인", "누락분은 EdgePC 의 CSV 사본 (7일 보관) 으로 확인",
                          "수집 계획 변경은 중앙 서버에서 하고 EdgePC 가 다시 받음"])
  .num_cards([("알림 수신", ["누락: 생산기술팀 설비 담당자 2명", "보고 끊김 (현황 보고가 끊긴 경우): 정보팀 운영 담당자 1명"])],
             numbered=False))
for col in cols:
    col.done()

d.section("검수 지표 및 결과")
s = d.page("검수 지표 및 집계 방법", lead="검수 지표 4개의 기존 수치, 목표 수치 및 계산식 확정")
a, b = cols = report_panels(s, ["검수 지표 (기존 수치 및 목표 수치)", "지표 집계 방법"], widths=(7.9, 4.6))
(a.kpi_table([("수집 성공률", "측정 없음", "99.5% 이상", "저장 회차 / 수집 회차 × 100", "4차 시험 충족"),
              ("0.2초 이내 저장", "측정 없음", "95% 이상", "0.2초 이내 저장 회차 / 저장 회차 × 100", "4차 시험 충족"),
              ("기록 간격", "1일", "1초", "다음 수집 시각 - 수집 시각 (초)", "4차 시험 충족 (수집 회차 1초 간격)"),
              ("누락 확인", "월말 대조", "100% (5분 이상 누락 구간 전부 알림)", "알림 발송 구간 / 5분 이상 누락 구간 × 100",
               "현장 검수에서 확인")], row_h=0.4, last="판정", title="검수 지표 정의 및 판정")
  .notes(["수집 지연 = 저장 시각 - 수집 시각", "출처: 4차 시험 결과서 (2026.05.27), 현장 확인 (2026.01.16)"]))
(b.collect_table([("수집", "PLC 값 120개 (3대 합계)", "자동"), ("저장", "회차별 저장 시각", "자동"),
                  ("누락 확인", "5분 이상 누락 여부", "자동"), ("현장 검수", "설비 화면 값 및 웹 화면 값 대조", "수동")],
                 first="단계", title="지표 집계 단계 및 방법")
  .tab_box("자동 및 수동 기준", ["자동: 수집 프로그램 및 중앙 서버가 집계", "수동: 현업 담당자가 두 화면의 값을 대조"])
  .notes(["출처: 수집 항목 정의서 (2026.01.09), 4차 시험 결과서 (2026.05.27)"]))
for col in cols:
    col.done()

s = d.page("4차 시험 결과 (수집 성능)",
           lead="4차 시험에서 수집 성공률과 0.2초 이내 저장 목표를 충족했고 1 ~ 4차 시험 결함 22건은 조치 완료",
           notes="결함 2건의 이름과 수치를 말하고, 조치 후 재시험 완료 (2026.05.28) 로 이어감")   # notes: 발표자 노트 (말할 내용)
a, b, c = cols = report_panels(s, ["검수 지표 (4차 시험)", "수집 지연 분포 (4차 시험)", "시험 회차별 발견 결함"])
(a.kpis([("99.8", "%", "수집 성공률"), ("95.7", "%", "0.2초 이내 저장")], h=1.2)   # 핵심 수치는 덱 본문의 1/5 장까지
  .tab_box("판정 기준", ["목표: 수집 성공률 99.5% 이상", "목표: 0.2초 이내 저장 95% 이상", "결함 기준: 수집 지연 0.5초 초과"])
  .notes(["출처: 4차 시험 결과서 (2026.05.27)"]))
(b.hist([0.04] * 61200 + [0.15] * 21300 + [0.3] * 3500 + [0.7] * 227, [0, 0.1, 0.2, 0.5, 1.0], unit=" 초",
        name="저장 회차", count_unit="회", h=2.0, highlight=(0, 1), title="수집 지연 구간별 저장 회차")
  .tab_box("집계 결과", ["0.2초 이내 저장: 82,500회 / 86,227회 = 95.7%",
                       "0.5초 초과 227회는 저장은 됐으나 지연된 회차 (결함)"]))
(c.chart("bar", ["1차 시험", "2차 시험", "3차 시험", "4차 시험"], {"결함": [6, 9, 5, 2]}, h=2.0,
         number_format='0"건"', title="시험 회차별 발견 결함")
  .tab_box("주요 결함", ["2차 시험: 재연결 뒤 같은 값 중복 저장 4건", "3차 시험: 보관 큐 재전송 순서 오류 3건",
                       "4차 시험: 미저장 173회 (EdgePC 재시작)", "4차 시험: 0.5초 초과 지연 227회 (네트워크 재연결)"])
  .notes(["출처: 4차 시험 결과서 (2026.05.27)"]))
for col in cols:
    col.done()

d.section("일정 및 방법")


def pos(month, day, days):   # 날짜 위치: 월 번호 + 그 날 한가운데까지 지난 비율
    return month - 1 + (day - 0.5) / days


s = d.page("단계별 일정 (2026년)", lead="구현 및 시험을 마치고 현장 적용 중이며 2026.06.15 현장 검수 예정")
(a,) = cols = report_panels(s, ["2026년 1 ~ 6월"], widths=(12.8,))
a.gantt(["1월", "2월", "3월", "4월", "5월", "6월"],
        [("요구 정리", 0, 1, "done", "완료"), ("설계", 0.5, 2, "done", "완료"), ("구현", 1.5, 3.5, "done", "완료"),
         ("시험 (1 ~ 4차, 재시험)", 3, pos(5, 28, 31), "done", "완료"),
         ("현장 적용", pos(5, 28, 31), pos(6, 12, 30), "now", "진행"),
         ("현장 검수", pos(6, 15, 30), 0, "mile", "06.15 검수")],
        row_h=0.5, label_w=2.4, today=pos(5, 29, 31), title="2026년 단계별 일정").notes(["기준일 2026.05.29 (빨간 점선)"])
a.done()

s = d.page("진행 방법", lead="수집 항목 정의서를 받아 1초 수집을 구현하고 저장 회차 및 수집 지연으로 검증")
a, b, c = cols = report_panels(s, ["받은 자료 및 현장 확인", "구현", "검증"], widths=(3.6, 4.2, 3.6),
                               arrows=True)   # 순서가 있을 때만 화살표
(a.section("현업에서 받은 자료", ["수집 항목 정의서 (2026.01.09)", "현장 네트워크 구성도 (2026.01.16)"], row_h=0.4)
  .section("정의서 내용", ["설비별 PLC 주소", "수집 항목 이름 및 단위"], row_h=0.4)
  .section("현장 확인 (2026.01.16)", ["PLC 통신: Ethernet, Modbus TCP", "EdgePC 와 중앙 서버 사이 방화벽 개방"], row_h=0.4))
(b.section("수집", ["중앙 서버가 수집 계획 (주소, 항목) 전달", "PLC 값을 1초마다 읽고 수집 시각 기록"], row_h=0.4)
  .section("전송 및 보관", ["중앙 서버로 전송, CSV 사본 7일 보관", "실패분은 보관 후 다시 전송"], row_h=0.4)
  .section("현황 및 알림", ["현황을 1분마다 보고", "5분 이상 누락되면 설비 담당자에게 메일"], row_h=0.4))
(c.section("시험", ["24시간 연속 수집", "수집 회차 대비 저장 회차 대조", "수집 지연 (저장 시각 - 수집 시각) 측정"], row_h=0.4)
  .section("현장 검수 (2026.06.15)", ["현업 담당자와 설비 화면 값 및 웹 화면 값 대조", "누락 알림 동작 확인"], row_h=0.4)
  .section("기록", ["시험 결과서 및 검수 확인서 작성"], row_h=0.4))
for col in cols:
    col.done()

d.section("시스템 구성")
s = d.page("수집 구성 (설비 3대 기준)", doc=True,
           lead="설비 3대의 PLC 값을 EdgePC 가 1초마다 읽어 중앙 서버 데이터 저장소로 전송",
           figure="설비 데이터 수집 구성 (설비 3대)")   # 구성도 장 그림 제목: 도식 아래 가운데
g = DocGrid(s, right=None, rows=7, bottom=6.0)
g.source("source", [("plc", "설비 1"), ("plc", "설비 2"), ("plc", "설비 3")], title="설비 PLC")
plcs = g.icons
e = g.group("edge", "EdgePC", kind="edge", sub="설비 옆", header="수집 프로그램", cols=2, area=False)
s1 = e.sub(0, 0, 6, header="수집 및 보고")
cache, core = s1.box(1, "수집 계획 보관"), s1.box(3, "수집 코어")
report, alarm = s1.box(5, "현황 보고"), s1.box(6, "알림 메일 발송")
s2 = e.sub(1, 0, 4, header="저장 및 전송")
csv, send, spool = s2.box(2, "CSV 기록"), s2.box(3, "송신"), s2.box(4, "보관 큐 (7일)")
p = g.group("platform", "중앙 서버", cols=2, area=False)
s3 = p.sub(0, 0, 6, header="수집 관리")
plan, status = s3.box(1, "수집 계획 제공"), s3.box(5, "수집 프로그램 현황")
srv_alarm = s3.box(6, "보고 끊김 감지")
s4 = p.sub(1, 0, 3, header="저장 및 화면")
db_, web = s4.box(1, "설정 DB"), s4.box(3, "웹 화면")
s5 = p.sub(1, 6, 6)
lake = s5.box(6, "데이터 저장소 (1초 값)")
mail = g.bar("edge", "메일 서버", fill=GRAY)
link(s, plan, cache, color=GRAY, dash="dash", text="HTTPS", step=(1, "계획 받기"))
fan(s, core, plcs, color=ACCENT, step=(2, "PLC 읽기"))
link(s, core, csv, color=ACCENT, step=(3, "기록"))
link(s, core, send, color=ACCENT)
link(s, send, lake, color=ACCENT, text="HTTPS", step=(4, "전송"))
link(s, send, spool, color=GRAY, dash="sysDot", step=(5, "실패분 보관"), head="both")   # 보관 뒤 다시 전송
link(s, report, status, color=GRAY, step=(6, "현황 보고"))
link(s, alarm, mail, color=GRAY, step=(7, "누락 알림 메일"))
link(s, srv_alarm, mail, color=GRAY, step=(8, "보고 끊김 메일"))
link(s, plan, db_, color=GRAY, head="both")
link(s, web, db_, color=GRAY, head="both")
link(s, lake, web, color=ACCENT)
note_by(s, web, "값, 계획, 현황 조회")   # 설명은 가리키는 상자 옆
for x_, c_, t_, d_ in ((6.71, ACCENT, "수집 값 흐름", None), (8.11, GRAY, "수집 계획 전달", "dash"),
                       (9.61, GRAY, "설정, 현황 및 알림 메일", None), (11.51, GRAY, "실패 시 보관", "sysDot")):
    legend(s, x_, 6.42, [(c_, t_, d_)], kind="line")   # 선 범례 한 줄. 색은 둘, 구분은 선 모양

stamp(s, "개념도")   # 도장: 실데이터가 아닌 장 표시. 한 장에 하나
d.save(OUT)
print(OUT)
