"""장 유형 견본 라이브러리: 사람이 손으로 만든 장의 배치 수치를 모으고, 엔진 덱 장을 같은 유형 견본 범위와 비교
사용: samples.py add 사람.pptx [장 번호 ...] [--assets]   (번호 없으면 모든 장, 유형은 자동 분류)
      samples.py check 덱.pptx               (견본 범위를 크게 벗어난 장을 주의로, 실패 아님)
      samples.py list
견본은 수치만(글과 그림은 넣지 않음): 플러그인 기본 assets/samples.json + 개인 설정 samples.json
(add 는 개인 설정에, --assets 면 플러그인에)"""
import json
import os
import sys

from pptx import Presentation
from pptx.util import Emu

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from ixdeck import PROFILE_DIR  # noqa: E402

BASE_DB = os.path.join(HERE, "..", "assets", "samples.json")
USER_DB = os.path.join(PROFILE_DIR, "samples.json")
KEYS = {"shapes": "도형 수", "chars": "글자 수", "fills": "채움 색 수", "sizes": "글자 크기 종류",
        "pics": "그림 수", "lines": "연결선 수", "cover": "내용이 덮은 면적 비율"}
SLACK = 0.3   # 견본 최소 및 최대에서 이만큼 더 벗어나면 주의
LEAF = ".//p:sp | .//p:pic | .//p:cxnSp | .//p:graphicFrame"


def kind(sl):
    """그룹 안까지 봄. 연결선 셋 이상이면 표가 섞여도 구성도"""
    x = sl.shapes._spTree
    if len(x.xpath(".//p:cxnSp")) >= 3:
        return "구성도"
    if x.xpath(".//c:chart"):
        return "차트"
    if x.xpath(".//a:tbl"):
        return "표"
    return "보고"


def metrics(sl, sw):
    shapes = [s for s in sl.shapes if not s.is_placeholder]
    leaf = [e for s in shapes for e in
            (s._element.xpath(LEAF) if s._element.tag.endswith("grpSp") else [s._element])]
    fills, sizes, chars = set(), set(), 0
    for x in leaf:
        fills.update(x.xpath("./p:spPr/a:solidFill/a:srgbClr/@val"))
        sizes.update(x.xpath(".//a:rPr/@sz"))
        chars += sum(len(t.replace(" ", "")) for t in x.xpath(".//a:t/text()"))
    # 내용 면적: 0.1 in 격자로 덮인 칸 비율 (머리 띠 아래, 바닥글 위). 그룹은 그룹 테두리로
    top, bottom = 1.15, 7.2
    gw, gh = int(Emu(sw).inches * 10), int((bottom - top) * 10)
    grid = set()
    for s in shapes:
        if s.width is None or s.top is None:
            continue
        x0, y0 = Emu(s.left).inches, Emu(s.top).inches
        x1, y1 = x0 + Emu(s.width).inches, y0 + Emu(s.height).inches
        for gx in range(max(0, int(x0 * 10)), min(gw, int(x1 * 10) + 1)):
            for gy in range(max(0, int((y0 - top) * 10)), min(gh, int((y1 - top) * 10) + 1)):
                grid.add((gx, gy))
    return {"shapes": len(leaf), "chars": chars, "fills": len(fills), "sizes": len(sizes),
            "pics": sum(x.tag.endswith("}pic") for x in leaf),
            "lines": sum(x.tag.endswith("}cxnSp") for x in leaf),
            "cover": round(len(grid) / max(1, gw * gh), 3)}


def _read(p):
    return json.load(open(p, encoding="utf8")) if os.path.exists(p) else {}


def _load():
    db = _read(BASE_DB)
    for k, rows in _read(USER_DB).items():
        db[k] = db.get(k, []) + rows
    return db


def add(path, nums, to_assets=False):
    prs = Presentation(path)
    target = BASE_DB if to_assets else USER_DB
    db = _read(target)
    src = os.path.basename(path)
    for i, sl in enumerate(prs.slides, 1):
        if nums and i not in nums:
            continue
        k = kind(sl)
        rows = [r for r in db.get(k, []) if not (r["src"] == src and r["slide"] == i)]
        rows.append({"src": src, "slide": i, **metrics(sl, prs.slide_width)})
        db[k] = rows
        print(f"{i}쪽 → {k}")
    os.makedirs(os.path.dirname(target), exist_ok=True)
    json.dump(db, open(target, "w", encoding="utf8"), ensure_ascii=False, indent=1)
    print(f"저장: {target}")


def check(path):
    db = _load()
    prs = Presentation(path)
    out = []
    for i, sl in enumerate(prs.slides, 1):
        if sl.slide_layout.name != "본문_소제목 O":
            continue
        k = kind(sl)
        ref = db.get(k, [])
        if len(ref) < 2:
            continue
        m = metrics(sl, prs.slide_width)
        for key, name in KEYS.items():
            lo, hi = min(r[key] for r in ref), max(r[key] for r in ref)
            if m[key] < lo * (1 - SLACK) or m[key] > hi * (1 + SLACK) + (1 if key != "cover" else 0):
                out.append(f"주의 {i}쪽 ({k}): {name} {m[key]} (사람 견본 {len(ref)}장 {lo} ~ {hi})")
    return out


if __name__ == "__main__":
    sys.stdout.reconfigure(errors="replace")
    if len(sys.argv) < 2 or sys.argv[1] not in ("add", "check", "list"):
        sys.exit(__doc__)
    if sys.argv[1] == "add":
        args = [a for a in sys.argv[2:] if a != "--assets"]
        add(args[0], {int(n) for n in args[1:]}, "--assets" in sys.argv)
    elif sys.argv[1] == "check":
        res = check(sys.argv[2])
        print("\n".join(res) if res else "견본 범위 안 (또는 같은 유형 견본 2장 미만)")
    else:
        for k, rows in _load().items():
            print(f"{k}: {len(rows)}장 ({', '.join(sorted({r['src'] for r in rows}))})")
