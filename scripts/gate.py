"""덱 품질 관문: 빌드, 검사, 금지 표현, 렌더
사용: gate.py build.py [장 수] / gate.py --deck 덱.pptx [장 수] (build 없이 기존 덱) / gate.py --verify 덱.pptx"""
import glob
import hashlib
import json
import os
import posixpath
import re
import subprocess
import sys
import tempfile
import time
import zipfile
from datetime import datetime

sys.stdout.reconfigure(errors="replace")
HERE = os.path.dirname(os.path.realpath(__file__))
sys.path.insert(0, HERE)   # python -P 는 스크립트 폴더를 넣지 않음
from ixpaths import gate_record, legacy_gate_record  # noqa: E402
BAN = os.path.join(HERE, "check_ban.py")


def _sha(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def _record(pptx):
    """통과 기록 위치: 개인 설정 폴더 (납품 폴더에는 남기지 않음). 옛 위치(덱 옆)에 있으면 그것을 읽음"""
    new = gate_record(pptx)
    old = legacy_gate_record(pptx)
    return old if not os.path.exists(new) and os.path.exists(old) else new


def _version():
    """플러그인 판 번호 (plugin.json 먼저)"""
    for rel in (".claude-plugin/plugin.json", "package.json"):
        try:
            return json.load(open(os.path.join(HERE, "..", rel), encoding="utf8"))["version"]
        except (OSError, KeyError, ValueError):
            pass
    return "?"


def verify(pptx):
    """통과 기록과 파일 해시 대조. 같으면 0"""
    rec = _record(pptx)
    if not os.path.exists(pptx) or not os.path.exists(rec):
        print(f"통과 기록 없음: {rec} (gate.py 를 통과해야 생김)")
        return 1
    r = json.load(open(rec, encoding="utf8"))
    same = r.get("sha256") == _sha(pptx) and r.get("result") == "통과"
    print(f"통과 기록과 {'같음' if same else '다름'}: {pptx} (기록 {r.get('time')}, 판 {r.get('version')})")
    return 0 if same else 1


def slide_hashes(pptx):
    """(틀 해시, 장별 해시). 장은 presentation.xml 순서
    - 틀: 레이아웃, 마스터, 테마, media, 렌더 스크립트 등"""
    from edit_deck import slide_parts
    z = zipfile.ZipFile(pptx)

    def h(names, extra=()):
        m = hashlib.sha256()
        for n in names:
            m.update(n.encode() + z.read(n))
        for b in extra:
            m.update(b)
        return m.hexdigest()
    frame = h(sorted(n for n in z.namelist() if re.match(r"ppt/(slideLayouts|slideMasters|theme|media)/", n)
                     or n in ("ppt/presentation.xml", "ppt/_rels/presentation.xml.rels")),
              [open(os.path.join(HERE, f), "rb").read() for f in ("render.py", "render.ps1")])
    out = []
    for n in slide_parts(z):
        rel = n.replace("slides/", "slides/_rels/") + ".rels"
        names = [n, rel]
        if rel in z.namelist():
            for t in re.findall(r'Target="([^"]+)"', z.read(rel).decode("utf8")):
                p = posixpath.normpath(posixpath.join("ppt/slides", t))
                if p in z.namelist() and not re.match(r"ppt/(slideLayouts|notesSlides)/", p):
                    names.append(p)
        out.append(h(names))
    return frame, out


def run(cmd, cwd=None, timeout=600):
    try:
        r = subprocess.run(cmd, cwd=cwd, capture_output=True, timeout=timeout,
                           env={**os.environ, "PYTHONIOENCODING": "utf-8"})   # 자식 출력은 UTF-8
    except subprocess.TimeoutExpired:
        return 124, f"{timeout}초 안에 끝나지 않음: {cmd[-3:]}"
    out = (r.stdout + r.stderr).decode("utf-8", "replace")
    return r.returncode, out


def check_deck(pptx, want):
    """만든 덱에 검사, 금지 표현, 문장, 렌더, 글 뽑기. 통과 기록. 종료 코드 반환"""
    fails = []
    t0 = time.time()
    rc, out = run([sys.executable, "-P", os.path.join(HERE, "check.py"), pptx] + want)
    print("== 2. check\n" + out.strip())
    if rc:
        fails.append("check")
    if not os.path.exists(BAN):
        print(f"== 3. 금지 표현 검사 파일 없음: {BAN} (ixdeck/scripts 에 있어야 함)")
        fails.append("금지 표현 검사 없음")
    else:
        rc, out = run([sys.executable, "-P", BAN, pptx])
        print("== 3. 금지 표현\n" + out.strip())
        if rc:
            fails.append("금지 표현")
    rc, out = run([sys.executable, "-P", os.path.join(HERE, "check_text.py"), pptx])
    print("== 3-2. 문장 (주의는 읽고 판단)\n" + out.strip())
    if rc:
        fails.append("문장")
    rc, out = run([sys.executable, "-P", os.path.join(HERE, "samples.py"), "check", pptx])
    print("== 3-3. 사람 견본과 배치 수치 비교 (주의만)\n" + ("견본 비교 실행 실패 (관문 판정에는 넣지 않음)\n" if rc else "")
          + out.strip()[-1500:])
    name = os.path.splitext(os.path.basename(pptx))[0]
    img = os.path.join(tempfile.gettempdir(), "interx_gate", name)
    os.makedirs(img, exist_ok=True)
    frame, hs = slide_hashes(pptx)   # 바뀐 장만 렌더 (hashes.json 과 대조)
    hfile = os.path.join(img, "hashes.json")
    try:
        old = json.load(open(hfile, encoding="utf8"))
    except (OSError, ValueError):
        old = {}
    n = len(hs)
    reuse = (old.get("frame") == frame and len(old.get("slides", [])) == n
             and all(os.path.exists(os.path.join(img, f"s{i}.png")) for i in range(1, n + 1)))
    changed = [i for i in range(1, n + 1) if not reuse or old["slides"][i - 1] != hs[i - 1]]
    for f in glob.glob(os.path.join(img, "*.md")) + ([] if reuse else glob.glob(os.path.join(img, "*.png"))):
        os.remove(f)   # 예전 것 섞임 방지
    if os.path.exists(hfile):
        os.remove(hfile)   # 렌더가 끝나야 다시 씀
    pdf = os.path.splitext(pptx)[0] + ".pdf"   # 결과물: pptx 옆에 같은 이름
    sel = ["--slides", ",".join(map(str, changed)) if changed else "none"] if reuse else []
    rc, out = run([sys.executable, "-P", os.path.join(HERE, "render.py"), pptx, img, "--pdf", pdf] + sel)
    pngs = sorted((f for f in os.listdir(img) if f[0] == "s" and f[1:-4].isdigit()), key=lambda f: int(f[1:-4]))
    if rc or not pngs:
        fails.append("렌더")
        print("== 4. 렌더 실패\n" + out.strip()[-500:])
    else:
        from PIL import Image
        W, H = 800, 450
        cols = 2
        rows = (len(pngs) + cols - 1) // cols
        sheet = Image.new("RGB", (W * cols, H * rows), "white")
        for k, f in enumerate(pngs):
            with Image.open(os.path.join(img, f)) as im:
                sheet.paste(im.resize((W, H)), ((k % cols) * W, (k // cols) * H))
        sheet.save(os.path.join(img, "sheet.png"))
        json.dump({"frame": frame, "slides": hs}, open(hfile, "w", encoding="utf8"))
        print(f"== 4. 렌더 {len(pngs)}장: {img} (모아 보기 sheet.png). 장마다 SKILL.md '장 검토 목록' 으로 본다")
        print(f"   바뀐 장 {len(changed)}장 렌더" + (f" {changed}" if 0 < len(changed) < n else "")
              + f", 나머지 {n - len(changed)}장은 이전 그림 재사용. 걸린 시간 {time.time() - t0:.0f}초")
        try:   # 지난 통과 때와 엔진 판이 다르면 바뀐 장을 눈으로 확인
            prev = json.load(open(_record(pptx), encoding="utf8")).get("version")
        except (OSError, ValueError):
            prev = None
        if prev and prev != _version():
            print(f"   엔진 판 {prev} → {_version()}: 그림이 바뀐 장 {changed} 를 눈으로 확인 (같은 글인데 바뀌면 엔진 영향)")
        rc, out = run([sys.executable, "-P", os.path.join(HERE, "deck_text.py"), pptx, os.path.join(img, "text.md"),
                      os.path.join(img, "titles.md")])
        if rc:
            fails.append("글 뽑기")
            print("== 5. 글 뽑기 실패\n" + out.strip()[-500:])
        print(f"== 5. 글 검토: {os.path.join(img, 'text.md')}, 요약 줄 목록 {os.path.join(img, 'titles.md')}, sheet.png 로 "
              "using SKILL '사람이 쓴 덱 검토'")
        print(f"== 파일: {pptx}")
        print(f"         {pdf}")
    new, old = gate_record(pptx), legacy_gate_record(pptx)
    try:
        if os.path.exists(old):
            os.remove(old)   # 덱 옆의 옛 기록은 치움 (납품 폴더에 부산물을 남기지 않음)
    except OSError:
        pass   # 지울 권한이 없어도 관문 결과는 그대로
    if fails:
        if os.path.exists(new):
            os.remove(new)   # 실패한 덱에 예전 통과 기록을 남기지 않음
    else:   # 넘기기 전 gate.py --verify 가 이 기록과 대조
        os.makedirs(os.path.dirname(new), exist_ok=True)
        for f in glob.glob(os.path.join(os.path.dirname(new), "*.json")):   # 덱이 없고 30일 넘게 안 쓴 기록만 치움 (USB, 동기화 폴더 보호)
            try:
                with open(f, encoding="utf8") as fh:
                    gone = not os.path.exists(json.load(fh).get("path", f))
                if gone and time.time() - os.path.getmtime(f) > 30 * 86400:
                    os.remove(f)
            except (OSError, ValueError):
                pass
        json.dump({"pptx": os.path.basename(pptx), "path": os.path.abspath(pptx), "sha256": _sha(pptx), "result": "통과", "version": _version(),
                   "time": datetime.now().isoformat(timespec="seconds")},
                  open(new, "w", encoding="utf8"), ensure_ascii=False, indent=1)
    print("결과: " + ("실패 " + ", ".join(fails) if fails else "통과"))
    return 1 if fails else 0


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    if sys.argv[1] == "--verify":
        sys.exit(verify(os.path.abspath(sys.argv[2])) if len(sys.argv) == 3 else 2)
    if sys.argv[1] == "--deck":   # build 없이 기존 덱 (edit_deck.py 결과 등)
        if len(sys.argv) not in (3, 4) or (len(sys.argv) == 4 and not sys.argv[3].isdigit()):
            print(__doc__)
            sys.exit(2)
        pptx = os.path.abspath(sys.argv[2])
        if not os.path.isfile(pptx):
            print(f"== 1. 덱 파일이 없음: {pptx}")
            sys.exit(1)
        print(f"== 1. 기존 덱 (build 없음): {pptx}")
        sys.exit(check_deck(pptx, sys.argv[3:]))
    build = os.path.abspath(sys.argv[1])
    want = sys.argv[2:]
    if len(want) > 1 or (want and not want[0].isdigit()):
        print("둘째 인자는 장 수 하나만 (--soft 같은 옵션은 관문에서 쓰지 않음)")
        sys.exit(2)
    src = open(build, encoding="utf8").read()
    if "strict=False" in src or "--soft" in src:
        print("== 1. build 실패: build.py 가 strict=False 나 --soft 로 검사를 끔. 빼고 걸린 장을 고침")
        sys.exit(1)
    t0 = time.time()
    rc, out = run([sys.executable, "-P", build], cwd=os.path.dirname(build))
    lines = [l.strip() for l in out.splitlines() if l.strip()]
    pptx = next((l for l in reversed(lines) if l.lower().endswith(".pptx")), None)
    if pptx:
        pptx = os.path.join(os.path.dirname(build), pptx)  # 상대 경로는 build.py 폴더 기준
    if rc != 0 or not pptx or not os.path.exists(pptx) or os.path.getmtime(pptx) < t0 - 1:
        print("== 1. build 실패 (오류, 경로 없음, 또는 이번에 저장하지 않은 예전 파일)\n" + "\n".join(lines[-25:]))
        sys.exit(1)
    print(f"== 1. build OK: {pptx}")
    sys.exit(check_deck(pptx, want))


if __name__ == "__main__":
    main()
