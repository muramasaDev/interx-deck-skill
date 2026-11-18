"""개인 설정과 통과 기록의 위치 (플러그인 밖, 납품 폴더에는 아무것도 남기지 않음)"""
import hashlib
import os

BASE = os.environ.get("CLAUDE_CONFIG_DIR") or os.path.join(os.path.expanduser("~"), ".claude")
PROFILE_DIR = os.path.join(BASE, "ixdeck", "profile")
_OLD = os.path.join(BASE, "interx-deck")   # 1.7.x 까지의 이름


def _merge(src, dst):
    """src 에서 dst 에 없는 것만 옮김 (같은 이름의 폴더는 안으로 들어가 합침). 다 옮기면 src 를 지움"""
    os.makedirs(dst, exist_ok=True)
    for n in os.listdir(src):
        a, b = os.path.join(src, n), os.path.join(dst, n)
        if os.path.isdir(a) and os.path.isdir(b):
            _merge(a, b)
        elif not os.path.exists(b):
            os.rename(a, b)
    try:
        os.rmdir(src)
    except OSError:
        pass


def _migrate():
    """옛 개인 설정 폴더를 새 폴더로 합침. 새 폴더가 먼저 생겨 있어도(시작 훅 등) 빠진 것만 옮기고, 잠긴 파일은 다음 실행 때"""
    if os.path.isdir(_OLD):
        try:
            _merge(_OLD, os.path.dirname(PROFILE_DIR))
        except OSError:
            pass


_migrate()
GATE_DIR = os.path.join(PROFILE_DIR, "gate")


def gate_record(deck):
    """덱 경로마다 하나인 통과 기록 파일 (이름이 같은 덱이 여러 폴더에 있어도 섞이지 않게 경로 해시를 붙임)"""
    p = os.path.abspath(deck)
    tag = hashlib.sha1(p.lower().encode("utf8")).hexdigest()[:10]
    return os.path.join(GATE_DIR, f"{os.path.splitext(os.path.basename(p))[0]}_{tag}.json")


def legacy_gate_record(deck):
    """1.7.1 까지 덱 옆에 두던 위치"""
    return os.path.splitext(os.path.abspath(deck))[0] + ".gate.json"
