"""사람 판별 시험: 사람이 만든 덱과 엔진 덱의 본문 장을 섞어 검토자가 사람 것을 맞히게 함
사용: blind.py make 폴더 --human 사람.pptx [...] --engine 엔진.pptx [...] [--n 8]
      blind.py score 폴더   (폴더의 answers.json 을 채점, 결과를 프로필 blind_history.csv 에 쌓음)
검토자는 폴더의 prompt.md 를 읽고 answers.json 을 씀. 엔진 장을 사람 것으로 본 비율이 높을수록 좋음"""
import argparse
import csv
import json
import os
import random
import shutil
import subprocess
import sys
import tempfile
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from ixdeck import PROFILE_DIR  # noqa: E402

PROMPT = """# 사람 판별 시험
이 폴더의 그림 {n}장은 PowerPoint 장입니다. 일부는 사람이 PowerPoint 에서 손으로 만들었고, 일부는 프로그램이 만들었습니다.
회사 양식(머리 띠, 로고, 바닥글)은 모두 같으니 판정 근거로 쓰지 않습니다. 주제와 분야 차이도 근거로 쓰지 않습니다.
장마다 배치, 여백, 글 밀도, 정렬, 선과 아이콘 쓰임, 문장 꼴을 보고 판정합니다.

같은 폴더에 answers.json 을 씁니다:
{{"01": {{"who": "사람" 또는 "프로그램", "why": "근거 한 줄"}}, "02": ...}}
"""


def _keyfile(d):
    """답안은 검토자가 보는 폴더와 먼 프로필 폴더에"""
    k = os.path.join(PROFILE_DIR, "blind_keys")
    os.makedirs(k, exist_ok=True)
    return os.path.join(k, os.path.basename(os.path.normpath(d)) + ".key.json")


def _render(deck, out):
    r = subprocess.run([sys.executable, "-P", os.path.join(HERE, "render.py"), deck, out], capture_output=True)
    if r.returncode:
        sys.exit(f"렌더 실패: {deck}\n{r.stdout.decode('utf8', 'replace')[-400:]}")
    return sorted((os.path.join(out, f) for f in os.listdir(out) if f.endswith(".png")),
                  key=lambda f: int(os.path.basename(f)[1:-4]))


def _body(pngs):
    """엔진 덱의 표지, 목차, 끝 장을 뺀 본문 장"""
    return pngs[2:-1] if len(pngs) > 4 else pngs


def make(a):
    os.makedirs(a.dir, exist_ok=True)
    for f in os.listdir(a.dir):   # 지난 시험의 그림과 답이 섞이지 않게
        if f.endswith(".png") or f in ("answers.json", "prompt.md"):
            os.remove(os.path.join(a.dir, f))
    rnd = random.Random(a.seed)
    pool = []
    for who, decks in (("사람", a.human), ("프로그램", a.engine)):
        cand = []
        for d in decks:
            pngs = _render(d, tempfile.mkdtemp(prefix="ixblind_"))
            cand += [(who, d, p) for p in (_body(pngs) if who == "프로그램" else pngs[1:])]   # 표지, 목차, 끝 장 뺌
        pool += rnd.sample(cand, min(a.n, len(cand)))
    rnd.shuffle(pool)
    key = {}
    for i, (who, d, p) in enumerate(pool, 1):
        k = f"{i:02d}"
        shutil.copy(p, os.path.join(a.dir, k + ".png"))
        key[k] = {"who": who, "deck": os.path.basename(d), "slide": os.path.basename(p)[1:-4]}
    json.dump(key, open(_keyfile(a.dir), "w", encoding="utf8"), ensure_ascii=False, indent=1)
    open(os.path.join(a.dir, "prompt.md"), "w", encoding="utf8").write(PROMPT.format(n=len(pool)))
    print(f"{len(pool)}장 준비: {a.dir} (검토자에게는 이 폴더 경로만 줌)")


def score(a):
    key = json.load(open(_keyfile(a.dir), encoding="utf8"))
    ans = json.load(open(os.path.join(a.dir, "answers.json"), encoding="utf8"))
    miss = sorted(set(key) - set(ans))
    if miss:
        sys.exit(f"답 없는 장: {miss}")
    bad = [k for k in key if not isinstance(ans[k], dict) or ans[k].get("who") not in ("사람", "프로그램")]
    if bad:
        sys.exit(f"who 가 '사람' 또는 '프로그램' 이 아닌 장: {bad}")
    eng = [k for k in key if key[k]["who"] == "프로그램"]
    hum = [k for k in key if key[k]["who"] == "사람"]
    fooled = [k for k in eng if ans[k]["who"] == "사람"]
    right = sum(ans[k]["who"] == key[k]["who"] for k in key)
    print(f"맞힘 {right}/{len(key)} ({right / len(key):.0%}), 엔진 장을 사람으로 봄 {len(fooled)}/{len(eng)}, "
          f"사람 장을 사람으로 봄 {sum(ans[k]['who'] == '사람' for k in hum)}/{len(hum)}")
    for k in eng:
        if k not in fooled:
            print(f"- 들킨 엔진 장 {key[k]['deck']} {key[k]['slide']}쪽: {ans[k].get('why', '')}")
    sys.path.insert(0, HERE)
    from gate import _version
    os.makedirs(PROFILE_DIR, exist_ok=True)
    path = os.path.join(PROFILE_DIR, "blind_history.csv")
    new = not os.path.exists(path)
    with open(path, "a", encoding="utf8", newline="") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["날짜", "엔진", "장 수", "맞힘", "엔진을 사람으로"])
        w.writerow([date.today().isoformat(), _version(), len(key), right, f"{len(fooled)}/{len(eng)}"])
    print(f"기록: {path}")


if __name__ == "__main__":
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["make", "score"])
    ap.add_argument("dir")
    ap.add_argument("--human", nargs="*", default=[])
    ap.add_argument("--engine", nargs="*", default=[])
    ap.add_argument("--n", type=int, default=8)
    ap.add_argument("--seed", type=int, default=None)
    a = ap.parse_args()
    if a.cmd == "make" and not (a.human and a.engine):
        sys.exit("--human 과 --engine 덱이 모두 필요")
    make(a) if a.cmd == "make" else score(a)
