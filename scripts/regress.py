"""다른 과제 덱 빌더를 새 엔진으로 다시 빌드해 검사하고 그림을 기준과 비교 (결과물은 임시 폴더, 원본 안 덮음)
사용: regress.py [--baseline] [build.py ...], 없으면 regress.txt 목록
--baseline: 지금 그림을 기준으로 저장 (엔진을 바꾸기 전에 한 번)"""
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.realpath(__file__))
sys.path.insert(0, HERE)
from ixdeck import PROFILE_DIR  # noqa: E402

BASE = os.path.join(PROFILE_DIR, "regress_base")
DIFF_PX = 0.0003   # 바뀐 화소 비율이 이보다 크면 바뀐 장 (머리 띠 하나 정도)


def run(cmd, cwd=None, env=None):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, timeout=900,
                       env={**os.environ, "PYTHONIOENCODING": "utf-8", **(env or {})})
    return r.returncode, (r.stdout + r.stderr).decode("utf-8", "replace")


def _key(build):
    return os.path.basename(os.path.dirname(build)) + "_" + hashlib.sha1(build.lower().encode()).hexdigest()[:8]


def _pngs(d):
    return sorted((f for f in os.listdir(d) if f.endswith(".png")), key=lambda f: int(f[1:-4]))


def compare(build, png_dir, baseline):
    """기준 그림과 장마다 비교. 바뀐 장은 나란히 붙인 그림을 남김"""
    from PIL import Image, ImageChops
    ref = os.path.join(BASE, _key(build))
    if baseline:
        shutil.rmtree(ref, ignore_errors=True)
        shutil.copytree(png_dir, ref)
        return [f"기준 저장 {len(_pngs(ref))}장: {ref}"]
    if not os.path.isdir(ref):
        return ["기준 그림 없음: regress.py --baseline 먼저"]
    now, old = _pngs(png_dir), _pngs(ref)
    out = [f"장 수 {len(old)} → {len(now)}"] if len(now) != len(old) else []
    for f in sorted(set(now) & set(old), key=lambda f: int(f[1:-4])):
        a, b = Image.open(os.path.join(ref, f)).convert("L"), Image.open(os.path.join(png_dir, f)).convert("L")
        if a.size != b.size:
            b = b.resize(a.size)
        diff = ImageChops.difference(a, b).point(lambda v: 255 if v > 40 else 0)
        ratio = sum(diff.histogram()[255:]) / (a.width * a.height)
        if ratio > DIFF_PX:
            pair = Image.new("RGB", (a.width * 2, a.height), "white")
            pair.paste(Image.open(os.path.join(ref, f)).convert("RGB").resize(a.size), (0, 0))
            pair.paste(Image.open(os.path.join(png_dir, f)).convert("RGB").resize(a.size), (a.width, 0))
            p = os.path.join(png_dir, "diff_" + f)
            pair.save(p)
            out.append(f"{f[1:-4]}쪽 바뀜 {ratio:.1%}: {p}")
    return out


def one(build, baseline):
    if not os.path.isfile(build):
        return "빌더 없음", ["경로가 바뀌었으면 regress.txt 를 고침"]
    out = tempfile.mkdtemp(prefix="ixregress_")
    try:
        rc, log = run([sys.executable, "-P", build], cwd=os.path.dirname(build), env={"INTERX_OUT_DIR": out, "INTERX_REGRESS": "1"})
    except subprocess.TimeoutExpired:
        return "빌드 시간 초과", ["15분 안에 끝나지 않음"]
    decks = [os.path.join(out, f) for f in os.listdir(out) if f.endswith(".pptx")]
    if rc or not decks:
        return "빌드 실패", log.strip().splitlines()[-3:]
    res = ["엄격"] if "엄격 실패: " in log else []
    notes = [ln[:160] for ln in log.splitlines() if ln.startswith("엄격 실패: ")][:3]
    for name, script in (("check", "check.py"), ("금지", "check_ban.py"), ("문장", "check_text.py")):
        for deck in decks:   # check.py 는 덱 하나씩
            r, o = run([sys.executable, "-P", os.path.join(HERE, script), deck])
            if not r:
                continue
            res.append(name)
            notes += [ln for ln in o.splitlines() if ln and not ln.startswith(("주의", "OK", "금지 표현 0", "문장 실패"))][:3]
    png = os.path.join(out, "png")
    r, o = run([sys.executable, "-P", os.path.join(HERE, "render.py"), decks[0], png])
    if r:
        res.append("렌더")
        notes += o.strip().splitlines()[-2:]
    else:
        changed = compare(build, png, baseline)
        if changed and not baseline and not changed[0].startswith("기준 그림 없음"):
            res.append("그림 바뀜")
        notes += changed
    return ("통과" if not res else "실패 " + ", ".join(res)), notes


def main():
    args = sys.argv[1:]
    baseline = "--baseline" in args
    builds = [a for a in args if a != "--baseline"]
    lst = os.path.join(PROFILE_DIR, "regress.txt")
    if not builds and os.path.exists(lst):
        builds = [ln.strip() for ln in open(lst, encoding="utf8") if ln.strip() and not ln.startswith("#")]
    if not builds:
        sys.exit(f"빌더 목록 없음: 인자로 주거나 {lst} 에 한 줄에 하나")
    bad = 0
    for b in builds:
        verdict, notes = one(os.path.abspath(b), baseline)
        bad += verdict != "통과"
        print(f"{verdict:12} {b}")
        for n in notes:
            print("             " + n)
    print(f"회귀: {len(builds) - bad}/{len(builds)} 통과")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    sys.stdout.reconfigure(errors="replace")
    main()
