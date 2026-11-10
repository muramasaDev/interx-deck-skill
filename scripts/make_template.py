"""기본 양식 PPT 로 assets/interx_template.pptx 생성"""
import posixpath
import re
import sys
import zipfile

import fitz  # SVG 를 투명 PNG 로

THEME_COLORS = {  # 견본 색에서 고름
    "dk2": "0E2841",      # 짙은 남색
    "lt2": "E8E8E8",      # 원본 유지: 바꾸면 표 머리칸이 파래짐
    "accent1": "0035D5",  # AAS 파랑 (기본 연결선 색이 이 값을 따름)
    "accent2": "EA580B",  # 주황 강조, 번호 원
    "accent3": "13C690",  # 초록 500 (ISH, EAC)
    "accent4": "546680",  # 남회색 600 (DB 상자)
    "accent5": "28303A",  # 남회색 900 (패널 머리)
    "accent6": "FF0000",  # 빨강 설명 문구
}
FONT = "Noto Sans KR"
TYPOS = {
    "Denstiy": "Density",
    "Schma ": "Schema ",
    "PostrgreSQL": "PostgreSQL",
    "idShot": "idShort",
}


def fix_theme(x: str) -> str:
    for k, v in THEME_COLORS.items():
        x, n = re.subn(rf"<a:{k}>.*?</a:{k}>", f'<a:{k}><a:srgbClr val="{v}"/></a:{k}>', x, count=1, flags=re.S)
        assert n == 1, k
    x = re.sub(r'(<a:(?:major|minor)Font>\s*<a:latin typeface=")[^"]*"', rf'\g<1>{FONT}"', x)
    x = re.sub(r'(<a:(?:major|minor)Font>\s*<a:latin typeface="[^"]*"[^>]*/>\s*<a:ea typeface=")[^"]*"', rf'\g<1>{FONT}"', x)
    x = re.sub(r'<a:font script="Hang" typeface="[^"]*"/>', f'<a:font script="Hang" typeface="{FONT}"/>', x)
    return x


def fix_part(x: str, counts: dict) -> str:
    x = re.sub(r'(<a:(?:latin|ea|cs|sym) typeface=")(?:Malgun Gothic|맑은 고딕|Calibri)"', rf'\g<1>{FONT}"', x)
    for a, b in TYPOS.items():
        counts[a] = counts.get(a, 0) + x.count(a)
        x = x.replace(a, b)
    return x


EMU_PER_IN = 914400
DPI = 220


def svg_to_png(svg: bytes, px_w: int) -> bytes:
    doc = fitz.open(stream=svg, filetype="svg")
    page = doc[0]
    zoom = max(px_w / page.rect.width, 0.1)
    return page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=True).tobytes("png")


def add_png_fallbacks(files: dict) -> int:
    """svgBlip 만 있는 그림에 PNG 대체를 단다"""
    made = 0
    seq = 0
    for n in sorted(files):
        if not (re.match(r"ppt/(slides|slideLayouts|slideMasters)/[^/]+\.xml$", n)):
            continue
        x = files[n].decode("utf8")
        rels_n = posixpath.join(posixpath.dirname(n), "_rels", posixpath.basename(n) + ".rels")
        if rels_n not in files:
            continue
        rels = files[rels_n].decode("utf8")
        ids = {m.group(1): m.group(2) for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)}
        ids.update({m.group(2): m.group(1) for m in re.finditer(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)})
        next_id = max([int(i[3:]) for i in ids if i.startswith("rId")] + [0]) + 1
        out, pos = [], 0
        for m in re.finditer(r'<p:pic>.*?</p:pic>', x, re.S):
            pic = m.group(0)
            b = re.search(r'<a:blip(?![^>]*r:embed)([^>]*)>', pic)
            sv = re.search(r'svgBlip[^>]*r:embed="(rId\d+)"', pic)
            if not (b and sv):
                continue
            target = ids[sv.group(1)]
            svg_path = posixpath.normpath(posixpath.join(posixpath.dirname(n), target))
            cx = re.search(r'<a:ext cx="(\d+)"', pic)
            px_w = int(int(cx.group(1)) / EMU_PER_IN * DPI) if cx else 1600
            png = svg_to_png(files[svg_path], max(px_w, 64))
            seq += 1
            png_name = f"ppt/media/svgfallback{seq}.png"
            files[png_name] = png
            rid = f"rId{next_id}"
            next_id += 1
            rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                                f'officeDocument/2006/relationships/image" Target="../media/svgfallback{seq}.png"/>'
                                "</Relationships>")
            new_pic = pic.replace(b.group(0), f'<a:blip r:embed="{rid}"{b.group(1)}>', 1)
            out.append(x[pos:m.start()] + new_pic)
            pos = m.end()
            made += 1
        if out:
            files[n] = ("".join(out) + x[pos:]).encode("utf8")
            files[rels_n] = rels.encode("utf8")
    ct = files["[Content_Types].xml"].decode("utf8")
    if 'Extension="png"' not in ct:
        ct = ct.replace("<Types ", "<Types ", 1).replace(
            "<Default ", '<Default Extension="png" ContentType="image/png"/><Default ', 1)
        files["[Content_Types].xml"] = ct.encode("utf8")
    return made


def fit_layout_pictures(path: str) -> int:
    """레이아웃 그림을 장 안에 맞춘다"""
    from pptx import Presentation
    prs = Presentation(path)
    W, H = prs.slide_width, prs.slide_height
    n = 0
    holders = [prs.slide_masters[0]] + list(prs.slide_masters[0].slide_layouts)
    for h in holders:
        for sh in h.shapes:
            if sh.shape_type != 13 or sh.width is None:
                continue
            x0, y0, x1, y1 = sh.left, sh.top, sh.left + sh.width, sh.top + sh.height
            if x0 >= 0 and y0 >= 0 and x1 <= W and y1 <= H:
                continue
            k = min(1.0, W / sh.width, H / sh.height)
            nw, nh = int(sh.width * k), int(sh.height * k)
            cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
            nx = int(min(max(cx - nw / 2, 0), W - nw))
            ny = int(min(max(cy - nh / 2, 0), H - nh))
            sh.left, sh.top, sh.width, sh.height = nx, ny, nw, nh
            n += 1
    prs.save(path)
    return n


def main(src: str, dst: str) -> None:
    counts: dict = {}
    with zipfile.ZipFile(src) as zi:
        infos = zi.infolist()
        files = {i.filename: zi.read(i.filename) for i in infos}
    for n in list(files):
        if n == "ppt/theme/theme1.xml":
            files[n] = fix_theme(files[n].decode("utf8")).encode("utf8")
        elif n.startswith(("ppt/slides/slide", "ppt/slideLayouts/", "ppt/slideMasters/")) and n.endswith(".xml"):
            files[n] = fix_part(files[n].decode("utf8"), counts).encode("utf8")
    made = add_png_fallbacks(files)
    order = [i.filename for i in infos] + sorted(set(files) - {i.filename for i in infos})
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in order:
            zo.writestr(n, files[n])
    print("오탈자 고친 수:", counts)
    print("SVG 에 PNG 대체 그림을 단 수:", made)
    print("장 밖으로 나간 레이아웃 그림을 맞춘 수:", fit_layout_pictures(dst))


def fix_deck(src: str, dst: str) -> None:
    """만든 덱에 PNG 대체 그림만 단다"""
    with zipfile.ZipFile(src) as zi:
        infos = zi.infolist()
        files = {i.filename: zi.read(i.filename) for i in infos}
    made = add_png_fallbacks(files)
    order = [i.filename for i in infos] + sorted(set(files) - {i.filename for i in infos})
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in order:
            zo.writestr(n, files[n])
    print("SVG 에 PNG 대체 그림을 단 수:", made)


if __name__ == "__main__":
    # make_template.py <원본> <템플릿>
    # make_template.py --png-fallback <덱> <출력>
    if sys.argv[1] == "--png-fallback":
        fix_deck(sys.argv[2], sys.argv[3])
    else:
        main(sys.argv[1], sys.argv[2])
