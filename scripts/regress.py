"""다른 과제 덱 빌더를 새 엔진으로 다시 빌드해 검사 (결과물은 임시 폴더, 원본 안 덮음)
사용: regress.py [build.py ...], 없으면 regress.txt 목록"""
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from ixdeck import PROFILE_DIR  # noqa: E402


def run(cmd, cwd=None, env=None):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, timeout=900,
                       env={**os.environ, "PYTHONIOENCODING": "utf-8", **(env or {})})
    return r.returncode, (r.stdout + r.stderr).decode("utf-8", "replace")


def one(build):
    if not os.path.isfile(build):
        return "빌더 없음", ["경로가 바뀌었으면 regress.txt 를 고침"]
    out = tempfile.mkdtemp(prefix="ixregress_")
    try:
        rc, log = run([sys.executable, "-P", build], cwd=os.path.dirname(build), env={"INTERX_OUT_DIR": out})
    except subprocess.TimeoutExpired:
        return "빌드 시간 초과", ["15분 안에 끝나지 않음"]
    decks = [os.path.join(out, f) for f in os.listdir(out) if f.endswith(".pptx")]
    if rc or not decks:
        return "빌드 실패", log.strip().splitlines()[-3:]
    res, notes = [], []
    for name, script in (("check", "check.py"), ("금지", "check_ban.py"), ("문장", "check_text.py")):
        for deck in decks:   # check.py 는 덱 하나씩
            r, o = run([sys.executable, "-P", os.path.join(HERE, script), deck])
            if not r:
                continue
            res.append(name)
            notes += [ln for ln in o.splitlines() if ln and not ln.startswith(("주의", "OK", "금지 표현 0", "문장 실패"))][:3]
    return ("통과" if not res else "실패 " + ", ".join(res)), notes


def main():
    builds = sys.argv[1:]
    lst = os.path.join(PROFILE_DIR, "regress.txt")
    if not builds and os.path.exists(lst):
        builds = [ln.strip() for ln in open(lst, encoding="utf8") if ln.strip() and not ln.startswith("#")]
    if not builds:
        sys.exit(f"빌더 목록 없음: 인자로 주거나 {lst} 에 한 줄에 하나")
    bad = 0
    for b in builds:
        verdict, notes = one(os.path.abspath(b))
        bad += verdict != "통과"
        print(f"{verdict:12} {b}")
        for n in notes:
            print("             " + n)
    print(f"회귀: {len(builds) - bad}/{len(builds)} 통과")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    sys.stdout.reconfigure(errors="replace")
    main()
