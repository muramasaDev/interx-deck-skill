"""채운 명세서 HTML 을 실제 브라우저(Chromium)에서 열어 양식 자신의 저장으로 되돌려 보고, 우리가 넣은 값과 비교한다.
양식은 불러올 때 맞지 않는 값을 조용히 지우므로(없는 D#, 쉼표 있는 숫자, 허용 밖 이미지 등) 이 비교가 '제출 때 값이 빠지지 않음'의 증거.
사용: python -P spec_verify.py 파일.html [--shots 폴더]
종료 코드: 0 통과, 1 값이 빠지거나 화면 오류, 2 브라우저를 못 띄움"""
import json
import os
import re
import sys

BLOCK = re.compile(r'<script type="application/json" id="form-data">([\s\S]*?)</script>')
SKIP = {"savedAt", "modifiedAt", "formVersion"}   # 저장 때 양식이 바꾸는 값


def _data(text):
    m = BLOCK.search(text)
    if not m:
        raise ValueError("form-data 블록이 없음")
    return json.loads(m.group(1) or "{}")


def _empty(v):
    return v is None or (isinstance(v, (str, list, dict)) and not v)   # 0 과 False 는 값


def diff(ours, saved, path=""):
    """우리가 넣은 값 중 저장본에서 사라지거나 바뀐 것. 저장본에 양식이 더한 기본 키는 무시"""
    out = []
    if isinstance(ours, dict):
        if not isinstance(saved, dict):
            return [f"{path}: 구조가 바뀜"] if not _empty(ours) else []
        for k, v in ours.items():
            if k in SKIP:
                continue
            out += diff(v, saved.get(k), f"{path}.{k}" if path else k)
    elif isinstance(ours, list):
        if not isinstance(saved, list):
            return [f"{path}: 목록이 사라짐"] if ours else []
        if len(saved) != len(ours):
            out.append(f"{path}: 항목 수 {len(ours)} -> {len(saved)} (양식이 지움)")
        for i, (a, b) in enumerate(zip(ours, saved)):
            out += diff(a, b, f"{path}[{i}]")
    elif not _empty(ours) and str(ours) != str(saved):
        out.append(f"{path}: '{str(ours)[:40]}' -> '{str(saved)[:40]}'")
    return out


def verify(path, shots=None):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("playwright 가 없음: pip install playwright 후 playwright install chromium")
        return 2
    ours = _data(open(path, encoding="utf8").read())
    errs, dialogs = [], []
    with sync_playwright() as p:
        try:
            b = p.chromium.launch(headless=True)
        except Exception as ex:   # 설치된 Chrome 이나 Edge 로 다시 시도
            for ch in ("chrome", "msedge"):
                try:
                    b = p.chromium.launch(headless=True, channel=ch)
                    break
                except Exception as ex2:
                    ex = ex2
            else:
                print("브라우저를 못 띄움:", str(ex)[:160])
                return 2
        page = b.new_context(accept_downloads=True, viewport={"width": 1400, "height": 1000}).new_page()
        page.on("console", lambda m: errs.append(m.text) if m.type == "error" else None)
        page.on("pageerror", lambda e: errs.append(str(e)))
        page.on("dialog", lambda d: (dialogs.append(d.message), d.dismiss()))
        page.add_init_script("window.showSaveFilePicker = undefined;")   # 파일 선택 창 대신 내려받기 방식으로 저장하게
        page.goto("file:///" + os.path.abspath(path).replace("\\", "/"))
        page.wait_for_timeout(1200)
        if shots:
            os.makedirs(shots, exist_ok=True)
            page.screenshot(path=os.path.join(shots, "home.png"), full_page=True)
        with page.expect_download(timeout=15000) as dl:
            page.click("#bSave")
        saved = _data(open(dl.value.path(), encoding="utf8").read())
        if shots:   # 첫 모델의 작성 화면도 한 장
            btn = page.get_by_role("button", name="작성").first
            if btn.count():
                btn.click()
                page.wait_for_timeout(500)
                page.screenshot(path=os.path.join(shots, "model1.png"), full_page=True)
        b.close()
    lost = diff(ours, saved)
    for e in errs:
        print("화면 오류:", e[:160])
    for d in dialogs:
        print("알림창:", d[:160])
    for x in lost[:60]:
        print("값이 바뀜 또는 사라짐:", x)
    if len(lost) > 60:
        print(f"... 외 {len(lost) - 60}건")
    n = len(ours.get("models", []))
    print(f"모델 {n}개, 양식 저장본과 비교: {'통과' if not (lost or errs or dialogs) else '실패'}")
    return 0 if not (lost or errs or dialogs) else 1


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    shots = sys.argv[sys.argv.index("--shots") + 1] if "--shots" in sys.argv else None
    if shots and shots in args:
        args.remove(shots)
    if not args:
        sys.exit("사용: spec_verify.py 파일.html [--shots 폴더]")
    sys.exit(verify(args[0], shots))
