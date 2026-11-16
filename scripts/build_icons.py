"""구성도용 두 색 아이콘(assets/icons, Phosphor duotone 및 Tabler, MIT)과
기술 로고(assets/logos, Simple Icons, CC0, 브랜드 색)를 PNG 로 만듦
사용: build_icons.py [받을 폴더]   (npm pack 으로 패키지를 받아 그림, resvg-py 필요)
남색 선과 청록 옅은 바탕. 이미 있는 이름은 덮지 않음. 끝에 _catalog.png 다시 그림"""
import os
import re
import subprocess
import sys
import tarfile

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
ICONS = os.path.join(HERE, "..", "assets", "icons")
LOGOS = os.path.join(HERE, "..", "assets", "logos")
PH, TB, SI = "@phosphor-icons/core@2.1.1", "@tabler/icons@3.52.0", "simple-icons@16.38.0"
BRANDS = ("docker kubernetes python grafana apachekafka redis nginx linux ubuntu redhat mqtt eclipsemosquitto "
          "nodered rabbitmq keycloak nodedotjs react fastapi pytorch tensorflow mariadb mysql prometheus "
          "elasticsearch gitlab github jenkins apacheairflow mlflow nvidia onnx opencv siemens apachespark").split()
LINE, TINT = "#002060", "#299288"
PX = 256

# 키: (패키지, 원본 이름). 패키지 p=Phosphor duotone, t=Tabler outline
SET = {
    "server_rack": ("p", "hard-drives"), "server_disk": ("p", "hard-drive"), "pc_tower": ("p", "desktop-tower"),
    "pc": ("p", "desktop"), "laptop": ("p", "laptop"), "monitor": ("p", "monitor"),
    "mobile": ("p", "device-mobile"), "tablet": ("p", "device-tablet"), "user": ("p", "user"),
    "users": ("p", "users"), "users_three": ("p", "users-three"), "user_admin": ("p", "user-gear"),
    "user_switch": ("p", "user-switch"), "id_badge": ("p", "identification-badge"), "cloud": ("p", "cloud"),
    "cloud_upload": ("p", "cloud-arrow-up"), "cloud_download": ("p", "cloud-arrow-down"),
    "cloud_check": ("p", "cloud-check"), "db": ("p", "database"), "cpu": ("p", "cpu"),
    "circuit": ("p", "circuitry"), "router": ("t", "router"), "wifi": ("p", "wifi-high"),
    "broadcast": ("p", "broadcast"), "lan": ("p", "network"), "plugs_on": ("p", "plugs-connected"),
    "plugs": ("p", "plugs"), "plug": ("p", "plug"), "usb": ("p", "usb"), "cell_signal": ("p", "cell-signal-full"),
    "firewall": ("p", "wall"), "shield_ok": ("p", "shield-check"), "shield_warn": ("p", "shield-warning"),
    "lock_key": ("p", "lock-key"), "key": ("p", "key"), "password": ("p", "password"),
    "fingerprint": ("p", "fingerprint"), "certificate": ("p", "certificate"), "seal_ok": ("p", "seal-check"),
    "globe": ("p", "globe"), "browser": ("p", "browser"), "code": ("p", "code"), "terminal": ("p", "terminal-window"),
    "api": ("t", "api"), "webhook": ("p", "webhooks-logo"), "git_branch": ("p", "git-branch"), "bug": ("p", "bug"),
    "gear": ("p", "gear-six"), "wrench": ("p", "wrench"), "toolbox": ("p", "toolbox"), "factory": ("p", "factory"),
    "package": ("p", "package"), "truck": ("p", "truck"), "robot": ("p", "robot"), "cube": ("p", "cube"),
    "stack": ("p", "stack"), "tree": ("p", "tree-structure"), "flow": ("p", "flow-arrow"), "graph": ("p", "graph"),
    "chart_up": ("p", "chart-line-up"), "chart_bar": ("p", "chart-bar"), "chart_pie": ("p", "chart-pie-slice"),
    "gauge": ("p", "gauge"), "thermo": ("p", "thermometer"), "power": ("p", "power"), "bolt": ("p", "lightning"),
    "bell": ("p", "bell-ringing"), "envelope": ("p", "envelope"), "file_text": ("p", "file-text"),
    "files": ("p", "files"), "file_csv": ("p", "file-csv"), "file_xls": ("p", "file-xls"),
    "folder_open": ("p", "folder-open"), "archive": ("p", "archive"), "clipboard": ("p", "clipboard-text"),
    "ok_circle": ("p", "check-circle"), "warn": ("p", "warning"), "ng_circle": ("p", "x-circle"),
    "calendar": ("p", "calendar"), "clock": ("p", "clock"), "timer": ("p", "timer"), "hourglass": ("p", "hourglass"),
    "search": ("p", "magnifying-glass"), "eye": ("p", "eye"), "brain": ("p", "brain"),
    "sync": ("p", "arrows-clockwise"), "swap": ("p", "swap"), "exchange": ("p", "arrows-left-right"),
    "chain": ("p", "link"), "qr": ("p", "qr-code"), "barcode": ("p", "barcode"), "scan": ("p", "scan"),
    "cctv": ("p", "video-camera"), "printer": ("p", "printer"), "sliders": ("p", "sliders-horizontal"),
    "list_check": ("p", "list-checks"), "kanban": ("p", "kanban"), "presentation": ("p", "presentation-chart"),
    "handshake": ("p", "handshake"), "buildings": ("p", "buildings"), "map_pin": ("p", "map-pin"),
    "goal": ("p", "target"), "trophy": ("p", "trophy"), "rocket": ("p", "rocket-launch"),
    "bulb": ("p", "lightbulb"), "chats": ("p", "chats"), "headset": ("p", "headset"), "ruler": ("p", "ruler"),
    "microscope": ("p", "microscope"), "test_tube": ("p", "test-tube"), "function": ("p", "function"),
    "table": ("p", "table"), "funnel": ("p", "funnel"), "download": ("p", "download-simple"),
    "upload": ("p", "upload-simple"), "export": ("p", "export"), "share": ("p", "share-network"),
    "teacher": ("p", "chalkboard-teacher"), "pointer": ("p", "hand-pointing"),
    "plc_unit": ("t", "cpu-2"), "sensor": ("t", "device-analytics"), "edge_box": ("t", "server-2"),
    "gateway": ("t", "affiliate"),
}


def fetch(dl):
    os.makedirs(dl, exist_ok=True)
    roots = {}
    for k, spec in (("p", PH), ("t", TB), ("s", SI)):
        name = spec.lstrip("@").replace("/", "-").replace("@", "-") + ".tgz"
        tgz = os.path.join(dl, name)
        if not os.path.exists(tgz):
            subprocess.run(["npm", "pack", spec, "-q"], cwd=dl, check=True, shell=os.name == "nt")
        out = tgz[:-4]
        if not os.path.isdir(out):
            with tarfile.open(tgz) as t:
                t.extractall(out, filter="data")
        roots[k] = os.path.join(out, "package")
    return roots


def svg_of(roots, pk, name):
    if pk == "p":
        s = open(os.path.join(roots["p"], "assets", "duotone", f"{name}-duotone.svg"), encoding="utf8").read()
        s = s.replace('fill="currentColor"', f'fill="{LINE}"')
        return re.sub(r'opacity="0\.2"', f'fill="{TINT}" opacity="0.35"', s)
    s = open(os.path.join(roots["t"], "icons", "outline", f"{name}.svg"), encoding="utf8").read()
    return s.replace('stroke="currentColor"', f'stroke="{LINE}"').replace('stroke-width="2"', 'stroke-width="1.6"')


def tint_inside(path):
    """선만 있는 아이콘의 닫힌 안쪽을 청록 옅은 색으로 (Phosphor duotone 과 맞춤)"""
    from PIL import ImageDraw as D
    im = Image.open(path).convert("RGBA")
    mask = Image.new("L", im.size, 0)
    mask.paste(255, mask=im.split()[3].point(lambda a: 255 if a > 40 else 0))
    D.floodfill(mask, (0, 0), 128)              # 바깥 = 128, 선 = 255, 닫힌 안쪽 = 0
    inside = mask.point(lambda v: 90 if v == 0 else 0)
    tint = Image.new("RGBA", im.size, TINT)
    tint.putalpha(inside)
    Image.alpha_composite(tint, im).save(path)


def brand_svg(roots, slug):
    import json
    hexes = {d["slug"]: d["hex"] for d in json.load(open(os.path.join(roots["s"], "data", "simple-icons.json"),
                                                          encoding="utf8"))}
    s = open(os.path.join(roots["s"], "icons", f"{slug}.svg"), encoding="utf8").read()
    return s.replace("<svg ", f'<svg fill="#{hexes[slug]}" ', 1)


def catalog(d=ICONS):
    files = sorted(f for f in os.listdir(d) if f[0] != "_" and f.lower().endswith((".png", ".jpg")))
    cols, cw, ch = 12, 120, 138
    im = Image.new("RGB", (cols * cw, -(-len(files) // cols) * ch), "white")
    dr = ImageDraw.Draw(im)
    try:
        font = ImageFont.truetype("arial.ttf", 12)
    except OSError:
        font = ImageFont.load_default()
    for i, f in enumerate(files):
        x, y = i % cols * cw, i // cols * ch
        ic = Image.open(os.path.join(d, f)).convert("RGBA")
        ic.thumbnail((100, 100))
        im.paste(ic, (x + (cw - ic.width) // 2, y + 4), ic)
        dr.text((x + 4, y + 112), os.path.splitext(f)[0], fill="black", font=font)
    im.save(os.path.join(d, "_catalog.png"))
    return len(files)


if __name__ == "__main__":
    import resvg_py
    roots = fetch(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", ".icon_dl"))
    made = 0
    for key, (pk, name) in SET.items():
        out = os.path.join(ICONS, key + ".png")
        if os.path.exists(out):
            continue
        png = bytes(resvg_py.svg_to_bytes(svg_string=svg_of(roots, pk, name), width=PX, height=PX))
        open(out, "wb").write(png)
        if pk == "t":
            tint_inside(out)
        made += 1
    for slug in BRANDS:
        out = os.path.join(LOGOS, slug + ".png")
        if not os.path.exists(out):
            open(out, "wb").write(bytes(resvg_py.svg_to_bytes(svg_string=brand_svg(roots, slug), width=PX, height=PX)))
            made += 1
    print(f"새 그림 {made}개, 아이콘 {catalog()}개, 로고 {catalog(LOGOS)}개")
