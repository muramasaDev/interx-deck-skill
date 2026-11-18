"""구성도 종류별 출발점 (build.py 로 복사, 내용만 교체). 어떤 종류를 쓸지는 skills/design/references/구성도_종류_고르기.md
 1. 계층형 가로 띠: 위에서 아래로 계층(단말, 서버, 제어기, 현장 장치)을 띠로 쌓고 공통 통신선이 가로지름
 2. 구역 배치도: DMZ, Private 같은 구역 안에 서버를 두고 서버 아래에 소프트웨어 구성 표
 3. 토폴로지: 아이콘 노드를 선으로 잇고 선 위에 프로토콜
 4. 등각: 바닥판 위에 육면체 서버 (편집되는 도형이라 정밀한 3D 는 아님)
- 자리표시 데이터만 있음 (이름과 수치는 형식을 보이기 위한 값)
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
OUT = sys.argv[1] if len(sys.argv) > 1 else d.filename("구성도 종류 예시", "0.1.0", tempfile.gettempdir())
d.cover("Sample System 구성도 종류", "2026.11.18")
d.section("구성")

# ---- 1. 계층형 가로 띠: 띠를 먼저 그리고 상자는 띠 안(x0 ~ x1)에
s = d.page("계층 구성도", doc=True, lead="관리자 단말에서 현장 PLC 까지 네 계층, 서버와 EdgePC 사이는 TCP/IP 로 연결",
           figure="Sample System 계층 구성도")
b = layer_bands(s, 0.75, 2.3, 11.8, [("관리자 단말", 0.9, ["웹 브라우저"]), ("중앙 서버", 0.9, ["화면, 저장"]),
                                     ("EdgePC", 0.9, ["수집 프로그램"]), ("설비 PLC", 0.9, ["현장 장치"])])
ry = [(y0 + y1) / 2 - 0.2 for y0, y1 in b["rows"]]
browser = role_box(s, 3.0, ry[0], 1.8, 0.4, "웹 브라우저", role="base")
web = role_box(s, 3.0, ry[1], 1.8, 0.4, "웹 화면 제공", role="platform", flow="process")
store = role_box(s, 6.4, ry[1], 1.8, 0.4, "데이터 저장소", role="platform", flow="sink")
prog = role_box(s, 6.4, ry[2], 1.8, 0.4, "수집 프로그램", role="ai", flow="process")
plc1 = role_box(s, 5.1, ry[3], 1.4, 0.4, "PLC 1", role="ext", flow="source")
plc2 = role_box(s, 8.1, ry[3], 1.4, 0.4, "PLC 2", role="ext", flow="source")
bus_line(s, 0.75, b["rows"][1][1], b["x1"] - 0.75, "TCP/IP (HTTPS)")
link(s, browser, web, color=GRAY, head="both", text="HTTPS")
link(s, web, store, color=GRAY, head="both", text="조회")
link(s, prog, store, color=ACCENT)   # 선 이름은 가로지르는 공통 통신선(TCP/IP)이 대신함
link(s, plc1, prog, color=ACCENT)
link(s, plc2, prog, color=ACCENT)
legend(s, 0.75, 6.05, [(ACCENT, "수집 값 흐름", None), (GRAY, "화면과 저장소 조회", None)], kind="line")

# ---- 2. 구역 배치도 + 계층(티어) 사양
s = d.page("구역 배치도", doc=True, lead="사용자는 인터넷을 거쳐 DMZ 의 웹 서버로 들어오고 내부 구역의 응용과 데이터베이스는 외부에 열리지 않음",
           figure="Sample System 3계층 구역 배치도")
users = node(s, 0.8, 2.9, "users", "현장 사용자")
net = node(s, 2.3, 2.9, "globe", "인터넷")
zone_box(s, 3.7, 2.3, 2.7, 4.0, "DMZ", kind="dmz")
zone_box(s, 6.6, 2.3, 5.9, 4.0, "내부망 (Private)", kind="private")
web = server(s, 4.05, 2.95, 2.0, 0.45, "웹 서버")
was = server(s, 7.0, 2.95, 2.0, 0.45, "응용 서버 (WAS)")
dbs = db(s, 10.1, 2.95, 2.0, 0.45, "데이터베이스 서버")
for x_, rows in ((4.05, [("소프트웨어", "Nginx"), ("운영체제", "Linux"), ("가상화", "가상 사설망")]),
                 (7.0, [("소프트웨어", "Tomcat"), ("운영체제", "Linux"), ("가상화", "가상 사설망")]),
                 (10.1, [("소프트웨어", "PostgreSQL"), ("운영체제", "Linux"), ("가상화", "가상 사설망")])):
    tier_spec(s, x_, 3.7, 2.0, rows)
link(s, users, net, color=GRAY, head="both")
link(s, net, web, color=GRAY, head="both", text="HTTPS")
link(s, web, was, color=GRAY, head="both", text="HTTP")
link(s, was, dbs, color=GRAY, head="both", text="SQL")

# ---- 3. 토폴로지: 아이콘 노드와 선 위 프로토콜
s = d.page("토폴로지 구성도", doc=True, lead="현장 PLC 값이 EdgePC 를 거쳐 방화벽 뒤 저장소로 가고, 화면은 같은 서버에서 조회",
           figure="Sample System 토폴로지")
mon = node(s, 3.3, 2.6, "monitor", "관리자 화면", "웹 브라우저")
plc = node(s, 0.9, 5.45, "plc", "설비 PLC")
edge = node(s, 3.3, 5.45, "pc", "EdgePC", "수집 프로그램")
fw = node(s, 5.9, 4.0, "firewall", "방화벽")
srv = node(s, 8.4, 4.0, "server_rack", "중앙 서버", "수집 및 화면")
dbn = node(s, 10.9, 4.0, "database", "데이터 저장소", "1초 값")
link(s, plc, edge, color=ACCENT, text="MC 프로토콜")
link(s, edge, fw, color=ACCENT, text="HTTPS")
link(s, mon, fw, color=GRAY, head="both", text="HTTPS")
link(s, fw, srv, color=ACCENT, text="HTTPS")
link(s, srv, dbn, color=ACCENT, text="SQL")

# ---- 4. 등각: 바닥판 위 육면체 (정밀한 3D 가 아니라 편집되는 도형)
s = d.page("등각 구성도", doc=True, lead="EdgePC 구역과 중앙 서버 구역을 바닥판으로 나누고 판 위에 서버를 놓음",
           figure="Sample System 등각 구성도")
iso_plate(s, 0.9, 2.9, 5.0, 3.4, "EdgePC 구역", tone="collect")
iso_plate(s, 6.6, 2.5, 5.6, 3.8, "중앙 서버 구역", tone="manage")
cl2 = iso_box(s, 1.5, 4.5, 0.6, "설비 PC", "client")
cl = iso_box(s, 3.6, 3.7, 0.8, "수집 프로그램", "server")
sv = iso_box(s, 7.8, 3.3, 0.8, "수집 서버", "server")
dv = iso_box(s, 10.2, 4.3, 0.7, "데이터 저장소", "db")
link(s, cl2, cl, color=ACCENT, text="값")
link(s, cl, sv, color=ACCENT, text="HTTPS")
link(s, sv, dv, color=ACCENT, text="SQL")
d.save(OUT)
print(OUT)
