"""덱 품질 관문: 빌드, 검사, 금지 표현, 렌더"""
import glob
import os
import subprocess
import sys
import tempfile
import time

sys.stdout.reconfigure(errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
BAN = os.path.join(HERE, "check_ban.py")


def run(cmd, cwd=None, timeout=600):
    try:
        r = subprocess.run(cmd, cwd=cwd, capture_output=True, timeout=timeout,
                           env={**os.environ, "PYTHONIOENCODING": "utf-8"})   # 자식 출력은 UTF-8
    except subprocess.TimeoutExpired:
        return 124, f"{timeout}초 안에 끝나지 않음: {cmd[-3:]}"
    out = (r.stdout + r.stderr).decode("utf-8", "replace")
    return r.returncode, out


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    build = os.path.abspath(sys.argv[1])
    want = sys.argv[2:]
    if len(want) > 1 or (want and not want[0].isdigit()):
        print("둘째 인자는 장 수 하나만 (--soft 같은 옵션은 관문에서 쓰지 않음)")
        sys.exit(2)
    src = open(build, encoding="utf8").read()
    if "strict=False" in src or "--soft" in src:
        print("== 1. build 실패: build.py 가 strict=False 나 --soft 로 검사를 끔. 빼고 걸린 장을 고침")
        sys.exit(1)
    fails = []
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
    rc, out = run([sys.executable, "-P", os.path.join(HERE, "check.py"), pptx] + want)
    print("== 2. check\n" + out.strip())
    if rc:
        fails.append("check")
    if not os.path.exists(BAN):
        print(f"== 3. 금지 표현 검사 파일 없음: {BAN} (interx-deck/scripts 에 있어야 함)")
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
    name = os.path.splitext(os.path.basename(pptx))[0]
    img = os.path.join(tempfile.gettempdir(), "interx_gate", name)
    os.makedirs(img, exist_ok=True)
    for old in glob.glob(os.path.join(img, "*.png")) + glob.glob(os.path.join(img, "text.md")):  # 예전 것 섞임 방지
        os.remove(old)
    pdf = os.path.splitext(pptx)[0] + ".pdf"   # 결과물: pptx 옆에 같은 이름
    rc, out = run([sys.executable, "-P", os.path.join(HERE, "render.py"), pptx, img, "--pdf", pdf])
    pngs = sorted((f for f in os.listdir(img) if f[0] == "s" and f[1:-4].isdigit()), key=lambda f: int(f[1:-4]))
    if rc or not pngs:
        fails.append("렌더")
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
        print(f"== 4. 렌더 {len(pngs)}장: {img} (모아 보기 sheet.png). 장마다 SKILL.md '장 검토 목록' 으로 본다")
        rc, out = run([sys.executable, "-P", os.path.join(HERE, "deck_text.py"), pptx, os.path.join(img, "text.md")])
        if rc:
            fails.append("글 뽑기")
            print("== 5. 글 뽑기 실패\n" + out.strip()[-500:])
        print(f"== 5. 글 검토: {os.path.join(img, 'text.md')} 와 sheet.png 로 using SKILL '사람이 쓴 덱 검토'")
        print(f"== 파일: {pptx}")
        print(f"         {pdf}")
    print("결과: " + ("실패 " + ", ".join(fails) if fails else "통과"))
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
