"""설계 문서 덱 출발점 (build.py 로 복사, 내용만 교체): 기능 상세, 화면 상세, 단계 장
- 자리표시 데이터만 있음 (시스템, 항목, 수치 모두 형식을 보이기 위한 값)
사용: python -P build.py [출력.pptx]   (gate.py 로 검사)"""
import os
import sys
import tempfile

from PIL import Image, ImageDraw


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
OUT = sys.argv[1] if len(sys.argv) > 1 else d.filename("설계서 예시", "0.1.0", tempfile.gettempdir())
d.cover("Sample System 설계서", "2026.11.18")   # 표지 제목에 제품이나 과제 이름을 넣음

d.section("기능 상세")
d.func_page("항목 편집 (M-001)", "항목을 단건 편집하고 최대 5,000개를 일괄 변경, 변경 이력을 기록", source="코드 및 시험 결과 (2026.11.18)",
            flow=["분류 선택 및 항목 목록 조회", "입력값 검증", "쓰기 잠금 확보", "항목 저장", "변경 이력 기록"],
            limits=[("수집 주기", "0.05 ~ 3600", "초", "빈 값은 기본 설정"), ("일괄 변경 상한", "5,000", "개", "한 번에 저장"),
                    ("쓰기 잠금", "60", "초", "-")],
            io=[("분류 코드", "입력", "문자열", "필수", "수정할 분류"), ("항목 이름", "입력", "문자열", "필수", "분류 안에서 유일"),
                ("변경 이력", "출력", "목록", "-", "변경 전, 변경 후")],
            errors=[("400", "입력 검증 실패", "항목별 안내 문구를 표시"), ("409", "항목 이름 중복", "이미 등록된 이름입니다 표시")],
            refs={"기능": "F-001 ~ 003", "화면": "SCR-01"})
d.func_page("상태 조회 (M-002)", "연결 상태를 60초 주기로 보고하고 180초 무응답이면 끊김으로 표시",
            flow=["상태 수집", "마지막 보고 시각 비교", "끊김 판정"],
            io=[("연결 이름", "입력", "문자열", "필수", "조회 대상"), ("상태", "출력", "문자열", "-", "정상, 주의, 끊김")])

d.section("화면 상세")
img = os.path.join(tempfile.gettempdir(), "ixdoc_example_ui.png")   # 실제 캡처 경로로 교체
im = Image.new("RGB", (1600, 900), "#F4F6F8")
dr = ImageDraw.Draw(im)
dr.rectangle([270, 100, 1560, 160], fill="#FFFFFF", outline="#D0D5DB")
dr.rectangle([270, 190, 1560, 800], fill="#FFFFFF", outline="#E1E5EA")
im.save(img)
d.screen_page("목록 화면 (SCR-01)", "항목 목록을 검색하고 편집 및 삭제하는 화면", image=img,
              marks=[(0.17, 0.11, 0.97, 0.18, 1), (0.17, 0.21, 0.97, 0.89, 2)],
              elements=[("1", "검색", "이름으로 검색", "빈 값이면 전체 표시"), ("2", "항목 표", "행 선택 후 수정 또는 삭제", "한 번에 최대 5,000개")],
              rules=[("표시 행", "50", "개", "쪽마다")], refs={"기능": "F-001"})
d.save(OUT)
print(OUT)
