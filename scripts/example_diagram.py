"""구성도 출발점 (build.py 로 복사, 내용만 교체): 읽히는 구성도의 세 그림
 1. 배치도: 어디에 무엇이 있고 값이 어디서 어디로 가는가 (제품 이름 영역이 하위 기능을 감쌈, 기능 색, 선 방향, 범례)
 2. 기능과 배치 대응표: 기능 묶음마다 모듈, 위치, 받는 값(원천), 내보내는 값(도착)
 3. 처리 순서: 시퀀스 (Mermaid)
- 자리표시 데이터만 있음 (시스템, 설비, 수치는 형식을 보이기 위한 값)
사용: python -P build.py [출력.pptx]   (gate.py 로 검사)"""
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

d = Deck(author=profile_author() or "소속 / 이름 직급", purpose="문서")   # 작성자는 덱마다 사용자에게 확인
OUT = sys.argv[1] if len(sys.argv) > 1 else d.filename("구성도 예시", "0.1.0", tempfile.gettempdir())
d.cover("Sample System 구성도", "2026.11.18")

d.section("구성")
s = d.page("배치도 (설비 3대 기준)", doc=True,
           lead="EdgePC 의 수집 프로그램이 서버의 수집 항목 및 주기대로 PLC 를 읽어 저장소로 전송, 현황과 알림을 서버에 보고",
           figure="Sample System 배치도")
g = DocGrid(s, right=None, rows=7, bottom=6.0)
g.source("source", [("plc", "설비 1"), ("plc", "설비 2"), ("plc", "설비 3")], title="설비 PLC")
plcs = g.icons
# header 는 제품 이름 띠, area=True 라서 그 아래 영역이 하위 묶음(수집 및 보고, 저장 및 전송)을 감쌈
e = g.group("edge", "EdgePC", kind="edge", sub="설비 옆", header="수집 프로그램", cols=2)
s1 = e.sub(0, 0, 6, header="수집 및 보고", tone="collect")
cache = s1.box(1, "항목 및 주기 보관", flow="sink")   # 받아서 보관만 함, 수집 코어가 조회
core = s1.box(3, "수집 코어", flow="process")
report = s1.box(5, "현황 보고", flow="process")   # 원천: 수집 코어 (아래에서 선으로 보임)
alarm = s1.box(6, "알림 메일 발송", flow="process")
s2 = e.sub(1, 0, 4, header="저장 및 전송", tone="send")
csv = s2.box(2, "CSV 기록", flow="sink")
send = s2.box(3, "송신", flow="process")
spool = s2.box(4, "보관 큐 (7일)")
p = g.group("platform", "중앙 서버", cols=2)
s3 = p.sub(0, 0, 6, header="수집 관리", tone="manage")
plan = s3.box(1, "항목 및 주기 제공", flow="process")
status = s3.box(5, "수집 프로그램 현황", flow="process")
srv_alarm = s3.box(6, "보고 끊김 감지", flow="process")
s4 = p.sub(1, 0, 3, header="저장 및 화면", tone="store")
db_, web = s4.box(1, "설정 DB"), s4.box(3, "웹 화면")
s5 = p.sub(1, 6, 6, tone="store")
lake = s5.box(6, "데이터 저장소 (1초 값)", flow="sink")
mail = g.bar("edge", "메일 서버", fill=GRAY)
link(s, plan, cache, color=GRAY, dash="dash", text="HTTPS", step=(1, "항목 및 주기 받기"))
fan(s, core, plcs, color=ACCENT, step=(2, "PLC 읽기"))
link(s, core, csv, color=ACCENT, step=(3, "기록"))
link(s, core, send, color=ACCENT)
link(s, send, lake, color=ACCENT, text="HTTPS", step=(4, "전송"))
link(s, send, spool, color=GRAY, dash="sysDot", head="both")   # 짧은 선은 번호를 붙이지 않음 (범례의 점선: 실패 시 보관)
link(s, core, report, color=GRAY)   # 현황의 원천: 수집 코어의 수집 건수와 지연
link(s, report, status, color=GRAY, step=(5, "현황 보고"))
link(s, report, alarm, color=GRAY)  # 알림의 원천: 현황 보고가 찾은 누락
link(s, alarm, mail, color=GRAY, step=(6, "누락 알림 메일"))
link(s, status, srv_alarm, color=GRAY)   # 끊김 감지의 원천: 받은 현황의 시각
link(s, srv_alarm, mail, color=GRAY, step=(7, "보고 끊김 메일"))
link(s, plan, db_, color=GRAY, head="both")
link(s, web, db_, color=GRAY, head="both")
link(s, lake, web, color=ACCENT)
for x_, c_, t_, d_ in ((6.71, ACCENT, "수집 값 흐름", None), (8.11, GRAY, "수집 항목 및 주기 전달", "dash"),
                       (9.61, GRAY, "설정, 현황 및 알림", None), (11.51, GRAY, "실패 시 보관", "sysDot")):
    legend(s, x_, 6.42, [(c_, t_, d_)], kind="line")   # 선 범례: 색 둘, 구분은 선 모양
for i, (k, n) in enumerate((("collect", "수집"), ("send", "전송"), ("manage", "관리"), ("store", "저장"))):
    legend(s, 0.75 + i * 1.2, 6.42, [(TONES[k][0], n + " 기능")], kind="square")   # 기능 묶음 색 범례
stamp(s, "개념도")

s = d.page("기능과 배치 대응표", lead="기능 묶음마다 모듈이 어디서 실행되고 값을 어디서 받아 어디로 내보내는지")
(a,) = cols = report_panels(s, ["기능 묶음별 모듈"], widths=(12.8,))
a.table([["기능 묶음", "모듈", "위치", "받는 값 (원천)", "내보내는 값 (도착)"],
         ["수집", "수집 코어", "EdgePC", "PLC 값, 수집 항목 및 주기", "송신, CSV 기록, 현황 보고"],
         ["수집", "항목 및 주기 보관", "EdgePC", "항목 및 주기 제공 (서버)", "보관만 함 (수집 코어가 조회)"],
         ["전송", "송신", "EdgePC", "수집 코어", "데이터 저장소, 보관 큐"],
         ["전송", "CSV 기록", "EdgePC", "수집 코어", "로컬 파일"],
         ["현황 및 알림", "현황 보고", "EdgePC", "수집 코어 (수집 건수, 지연)", "수집 프로그램 현황 (서버)"],
         ["현황 및 알림", "알림 메일 발송", "EdgePC", "현황 보고 (누락 감지)", "메일 서버"],
         ["관리", "항목 및 주기 제공", "중앙 서버", "설정 DB", "항목 및 주기 보관 (EdgePC)"],
         ["관리", "보고 끊김 감지", "중앙 서버", "수집 프로그램 현황 (보고 시각)", "메일 서버"]],
        ratio=[1.2, 1.6, 1.0, 2.6, 2.8], row_h=0.36, title="기능 묶음별 모듈과 입출력").notes(["출처: 설계 자료 (2026.11.18)"])
a.done()

s = d.page("수집 및 전송 순서", lead="수집 항목 및 주기를 받은 뒤 PLC 값을 읽어 기록하고 전송, 실패하면 보관 후 다시 전송")   # Mermaid 는 title= 이 그림 제목
mermaid(s, "\n".join(["sequenceDiagram", "participant S as 중앙 서버", "participant C as 수집 코어", "participant P as PLC",
                      "participant L as 데이터 저장소", "participant Q as 보관 큐", "S->>C: 수집 항목 및 주기 전달", "C->>P: 값 읽기 (1초마다)", "P-->>C: 값",
                      "C->>L: 값 전송", "L-->>C: 응답", "C->>Q: 실패분 보관"]),
        1.0, 2.0, 11.3, 4.8, title="수집 및 전송 순서")
d.save(OUT)
print(OUT)
