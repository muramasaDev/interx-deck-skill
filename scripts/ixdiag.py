"""구성도 역할 부품 (skills/design/references/구성도_레퍼런스_문법.md): 역할 색 상자, 그라데이션 블록 화살표,
프로토콜 알약, 하단 약어 풀이 영역. ixdeck.py 맨 아래에서 불러 와 같은 이름공간에서 씀"""
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from ixdeck import (BLACK, GRAY_L, GREEN, NAVY, RED, SLATE0, SLATE3, SLATE6, SLATE9, TONES, WHITE, _adj, _fill, _font, _lh, _line,
                    _no_style_effects, _obst, _reg, _set_flow, _set_text, _text_w, _wall, big_arrow, box, icon, label, text)

__all__ = ["AMBER", "ROLE", "role_box", "block_arrow", "pill", "abbr_legend", "step_header", "summary_bar", "speech_bubble", "headline",
           "backdrop", "ZONES", "layer_bands", "bus_line", "zone_box", "tier_spec", "node", "iso_plate", "iso_box"]

AMBER = "FFC000"
ROLE = {"base": (SLATE6, WHITE), "ai": (GREEN, WHITE), "ext": (AMBER, BLACK), "platform": (NAVY, WHITE)}   # 칩 색의 뜻 (바탕, 글자)


def role_box(s, x, y, w, h, t, role="base", sub=None, size=7.26, flow=None):
    """역할 색 칩. role: base 기존 모듈(회색), ai AI 및 에이전트(초록), ext 확장 항목(노랑), platform 핵심 서버(남색). sub: 작은 부제"""
    if role not in ROLE:
        raise ValueError(f"role_box role '{role}' 은 {', '.join(ROLE)} 중 하나")
    fill, fg = ROLE[role]
    sh = box(s, x, y, w, h, [(t, fg)] + ([(sub, fg)] if sub else []), fill=fill, size=size, bold=True, color=fg)
    _set_flow(s, sh, t, flow)   # flow: source, sink, process 이면 선이 빠졌는지 구성도 주의가 확인
    _reg(s).setdefault("box_names", []).append((sh.shape_id, t))   # 이름이 모호한지 구성도 주의가 확인
    if sub:   # 부제는 한 단계 작게
        for r in sh.text_frame.paragraphs[1].runs:
            r.font.size = Pt(5.44)
    return _obst(s, sh)


def block_arrow(s, x, y, w, h, label=None, direction="right"):
    """큰 흐름용 그라데이션 블록 화살표 (꼬리는 옅고 머리는 짙음). label: 가로면 위, 세로면 오른쪽에 굵은 글"""
    ow, oh = (h, w) if direction in ("down", "up") else (w, h)   # 세로는 돌리기 전 틀을 맞바꿔 (x, y, w, h) 가 보이는 크기가 되게
    sh = big_arrow(s, x + (w - ow) / 2, y + (h - oh) / 2, ow, oh, direction, colors=(SLATE3, SLATE6), adj=(0.5, 0.6))
    if label:
        lines = label if isinstance(label, list) else [label]
        th = len(lines) * _lh(8) + 0.04
        if direction in ("right", "left"):
            _obst(s, text(s, x, y - th - 0.02, w, th, lines, size=8, bold=True, align="c", anchor="b"))
        else:
            _obst(s, text(s, x + w + 0.08, y + (h - th) / 2, 2.2, th, lines, size=8, bold=True, align="l"))
    return _obst(s, sh)


def pill(s, x, y, t, color=GREEN, size=5.44):
    """연결선 위 프로토콜 알약 (작은 약어 태그). (x, y) 는 왼쪽 위"""
    return box(s, x, y, _text_w(t, size) + 0.14, 0.13, t, fill=color, size=size, bold=True, color=WHITE, geom="round",
               radius=0.5, inset=(0, 0, 0, 0))


def abbr_legend(s, x, y, w, bullets=(), abbrs=()):
    """구성도 아래 영역: 왼쪽 핵심 요약 줄(앞에 '- '), 오른쪽 약어 풀이('MCP : ...'). 아래 y 를 돌려줌"""
    _reg(s).setdefault("terms", set()).update(k for k, _ in abbrs)   # 풀이한 용어는 모호한 이름 주의에서 뺌
    n = max(len(bullets), len(abbrs))
    if not n:
        raise ValueError("abbr_legend 가 비었음 (bullets 나 abbrs 를 줌)")
    h = n * _lh(8) + 0.04
    lw = w * 0.62
    if bullets:
        _obst(s, text(s, x, y, lw, h, ["- " + b for b in bullets], size=8, color=BLACK, align="l", anchor="t"))
    if abbrs:
        _obst(s, text(s, x + lw, y, w - lw, h, [f"{k} : {v}" for k, v in abbrs], size=8, bold=True, color=BLACK,
                      align="r", anchor="t"))
    return y + h


def step_header(s, x, y, w, items, h=0.42):
    """단계형 장의 머리띠: 남색 쉐브론이 가로로 이어지고 단계마다 하나. items=[글].
    아래 열을 같은 폭에 맞추도록 [(열 x, 열 너비)] 를 돌려줌"""
    n, gap, tip = len(items), 0.06, 0.18
    if n < 2:
        raise ValueError("step_header 는 두 단계 이상")
    wc = (w - tip - gap * (n - 1)) / n
    cols = [(x + i * (wc + gap), wc) for i in range(n)]
    for i, (t, (cx, _)) in enumerate(zip(items, cols)):
        box(s, cx, y, wc + tip, h, t, fill=NAVY, size=10.5, bold=True, color=WHITE,
            geom="homePlate" if i == 0 else "chevron", inset=(0.1 if i == 0 else 0.22, 0, 0.14, 0))
    _reg(s)["obst"].append(((x, y, x + w, y + h), 0))   # 쉐브론끼리 일부러 겹치므로 띠 하나로 등록
    return cols


def summary_bar(s, cols, lines, y, h=0.7):
    """단계 머리띠 아래 둥근 요약바: 열마다 한 줄씩(앞에 '• '), 열 폭에 맞춤. cols 는 step_header 의 결과. 아래 y 를 돌려줌"""
    if len(lines) != len(cols):
        raise ValueError(f"summary_bar 줄 {len(lines)}개와 열 {len(cols)}개가 다름")
    x0, x1 = cols[0][0], cols[-1][0] + cols[-1][1]
    _obst(s, box(s, x0, y, x1 - x0 + 0.18, h, fill=SLATE0, line=GRAY_L, lw=0.75, geom="round", radius=0.12))
    for (cx, cw), t in zip(cols, lines):
        text(s, cx + 0.1, y, cw - 0.1, h, "• " + t, size=9, bold=True, color=BLACK, align="l")
    return y + h


def speech_bubble(s, x, y, w, h, t, tail="left"):
    """말풍선: 현업 말투 한두 문장 (이모지는 쓰지 않음). tail: left, right, down"""
    out = 0.32   # 꼬리 끝이 상자 밖으로 나오는 길이 (in)
    pos = {"left": (-0.5 - out / w, 0.25), "right": (0.5 + out / w, 0.25), "down": (-0.25, 0.5 + out / h)}
    if tail not in pos:
        raise ValueError(f"speech_bubble tail '{tail}' 은 {', '.join(pos)} 중 하나")
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT, Inches(x), Inches(y), Inches(w), Inches(h))
    _no_style_effects(sh)
    _adj(sh, *pos[tail], 0.16667)
    _fill(sh, WHITE)
    _line(sh, SLATE6, 1.0)
    _set_text(sh, t, 9, False, BLACK, "l", "m", (0.14, 0.06, 0.14, 0.06))
    return _obst(s, sh)


def headline(s, x, y, w, lines, size=14, accent=GREEN):
    """가운데 맞춤 헤드라인: 줄마다 [(글, 강조 여부)] 조각, 강조 조각만 accent 색 (초록, 파랑 등). 아래 y 를 돌려줌"""
    if not lines or any(not parts for parts in lines):
        raise ValueError("headline 줄이 비었음: [[(글, False), (키워드, True)], ...]")
    h = len(lines) * _lh(size) + 0.06
    sh = box(s, x, y, w, h, "", fill=None, size=size, color=BLACK, inset=(0, 0, 0, 0))
    for i, parts in enumerate(lines):
        p = sh.text_frame.paragraphs[0] if i == 0 else sh.text_frame.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        for r in list(p.runs):
            p._p.remove(r._r)
        for t, on in parts:
            r = p.add_run()
            r.text = t
            _font(r, size, False, accent if on else BLACK)
    _obst(s, sh)
    return y + h


def backdrop(s, cols, y, h):
    """캡처 뒤 연한 쉐브론 배경 (단계마다 하나). cols 는 step_header 의 결과, 캡처보다 먼저 그림"""
    for i, (cx, cw) in enumerate(cols):
        box(s, cx, y, cw + 0.18, h, fill="EEF1F5", geom="homePlate" if i == 0 else "chevron", inset=(0, 0, 0, 0))



# --- 계층, 구역, 토폴로지, 등각 (skills/design/references/구성도_종류_고르기.md)
ZONES = {"public": ("2E7D4F", "EAF4EC", "dash"), "private": ("2B6CB0", "E8F0FA", "dash"), "dmz": ("C2570C", "FDF0E6", "dash"),
         "cloud": ("546680", "F6F7F9", None)}   # 구역 종류: (이름표와 테두리 색, 옅은 바탕, 선 모양)


def _behind_check(s, region, what):
    """띠와 구역은 상자와 선보다 먼저 그려야 함 (나중에 그리면 위에 덮여 상자가 가려짐)"""
    x0, y0, x1, y1 = region
    if any(x0 <= r[0] and r[2] <= x1 and y0 <= r[1] and r[3] <= y1 and sid != -1 for r, sid in _reg(s)["obst"]):
        raise ValueError(f"{what} 영역 안에 이미 상자가 있음: {what} 를 상자와 선보다 먼저 부른다 (나중에 그리면 상자를 덮음)")


def layer_bands(s, x, y, w, layers, label_w=1.45):
    """계층 가로 띠: 위에서 아래로 계층을 띠로 쌓고 계층 이름은 오른쪽 열에 (예: 클라이언트, 서버, 제어기, 현장 장치).
    layers=[(이름, 높이, 부제 목록 또는 None)]. 상자를 놓기 전에 먼저 그림. 돌려주는 값: {"rows": [(y0, y1)], "x0": x, "x1": 안쪽 오른쪽 끝}"""
    if len(layers) < 2 or w <= label_w + 1 or any(h <= 0.2 for _, h, _ in layers):
        raise ValueError("layer_bands 는 두 계층 이상, 폭은 이름 열 + 1 in 이상, 계층 높이는 0.2 in 이상")
    _behind_check(s, (x, y, x + w, y + sum(h for _, h, _ in layers)), "layer_bands")
    rows, yy = [], y
    for i, (name, h, subs) in enumerate(layers):
        box(s, x, yy, w, h, fill=SLATE0 if i % 2 == 0 else WHITE, line=SLATE3, lw=0.5)
        lines = [(name, BLACK)] + [(t, SLATE9) for t in (subs or [])]
        th = _lh(9) + len(subs or []) * _lh(7.26) + 0.04
        sh = box(s, x + w - label_w + 0.1, yy + (h - th) / 2, label_w - 0.15, th, lines, fill=None, size=7.26, bold=False, align="l",
                 inset=(0, 0, 0, 0))
        for r in sh.text_frame.paragraphs[0].runs:
            _font(r, 9, True, BLACK)
        rows.append((yy, yy + h))
        yy += h
    return {"rows": rows, "x0": x, "x1": x + w - label_w}


def bus_line(s, x, y, w, name, color=RED):
    """계층을 가로지르는 공통 통신선 (예: TCP/IP). 왼쪽 위에 이름"""
    if w <= 0:
        raise ValueError(f"bus_line 폭 {w} 은 0 보다 커야 함")
    box(s, x, y - 0.02, w, 0.04, fill=color)
    _reg(s)["segs"].append(((x, y), (x + w, y)))   # 번호와 선 이름이 버스 위에 얹히지 않게 (선이 가로지르는 것은 허용)
    return _obst(s, text(s, x + 0.05, y - 0.22, _text_w(name, 7.26) + 0.2, 0.18, name, size=7.26, bold=True, color=color))


def zone_box(s, x, y, w, h, title, kind="private"):
    """구역 상자: 옅은 바탕에 점선 테두리, 왼쪽 위에 이름표 (DMZ, Private, 서브넷, 가용 영역). kind: public, private, dmz, cloud"""
    if kind not in ZONES:
        raise ValueError(f"zone_box kind '{kind}' 은 {', '.join(ZONES)} 중 하나")
    if w <= 0.5 or h <= 0.4:
        raise ValueError(f"zone_box 크기({w} x {h})가 너무 작음")
    _behind_check(s, (x, y, x + w, y + h), "zone_box")
    c, f, dash = ZONES[kind]
    z = box(s, x, y, w, h, fill=f, line=c, lw=1.0, dash=dash)
    _wall(s, z)
    tab = box(s, x, y, _text_w(title, 7.26) + 0.24, 0.2, title, fill=c, size=7.26, bold=True, color=WHITE, inset=(0.04, 0, 0.04, 0))
    _reg(s)["obst"].append(((x, y, x + _text_w(title, 7.26) + 0.24, y + 0.2), 0))
    return z


def tier_spec(s, x, y, w, rows, label_w=0.55):
    """서버 아래 사양 표 (3계층의 WEB, WAS, DB 아래에 제품과 운영체제를 보임). rows=[(항목, 내용)]. 아래 y 를 돌려줌"""
    if not rows or w <= label_w + 0.3:
        raise ValueError("tier_spec 은 rows 가 있어야 하고 폭은 이름 열 + 0.3 in 이상")
    rh = 0.22
    for i, (k, v) in enumerate(rows):
        box(s, x, y + i * (rh + 0.02), label_w, rh, k, fill=SLATE3, size=6.35, bold=True)
        box(s, x + label_w + 0.02, y + i * (rh + 0.02), w - label_w - 0.02, rh, v, fill=SLATE0, size=6.35, align="l", inset=(0.06, 0, 0.04, 0))
    _reg(s)["obst"].append(((x, y, x + w, y + len(rows) * (rh + 0.02)), 0))
    return y + len(rows) * (rh + 0.02)


def node(s, x, y, key, name, sub=None, size=0.55, w=1.1):
    """토폴로지 노드: 아이콘 (x, y 는 왼쪽 위) + 아래 굵은 이름 + 작은 부제. 아이콘 그림을 돌려줌 (link 로 이을 수 있음)"""
    pic = icon(s, key, x, y, w=size)
    cx = x + size / 2
    label(s, cx - w / 2, y + size + 0.03, w, name, size=7.26, color=BLACK, bold=True)
    if sub:
        label(s, cx - w / 2, y + size + 0.24, w, sub, size=6.35, color=SLATE9)
    return pic


def iso_plate(s, x, y, w, h, title, tone="collect"):
    """등각 바닥판 (평행사변형). 위에 iso_box 를 놓고 왼쪽 아래에 이름. tone: 기능 색 (TONES)"""
    if tone not in TONES:
        raise ValueError(f"tone '{tone}' 은 {', '.join(TONES)} 중 하나")
    if w <= 1 or h <= 0.5:
        raise ValueError(f"iso_plate 크기({w} x {h})가 너무 작음")
    p = s.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(x), Inches(y), Inches(w), Inches(h))
    _no_style_effects(p)
    _adj(p, 0.35)
    _fill(p, TONES[tone][1])
    _line(p, TONES[tone][0], 1.0, "dash")
    _wall(s, p)
    text(s, x + 0.35, y + h - 0.25, _text_w(title, 8) + 0.3, 0.2, title, size=8, bold=True, color=TONES[tone][0], align="l")
    return p


def iso_box(s, x, y, w, name, kind="server"):
    """등각 서버(육면체), 저장소(원통), 단말(작은 육면체) + 아래 이름. kind: server, db, client. 도형을 돌려줌 (link 로 이을 수 있음)"""
    if kind not in ("server", "db", "client") or w <= 0.2:
        raise ValueError("iso_box kind 는 server, db, client 중 하나이고 폭은 0.2 in 보다 커야 함")
    geo, fill, h = {"server": (MSO_SHAPE.CUBE, SLATE6, w * 0.9), "db": (MSO_SHAPE.CAN, "3E6A8A", w * 0.85),
                    "client": (MSO_SHAPE.CUBE, SLATE3, w * 0.7)}[kind]
    sh = s.shapes.add_shape(geo, Inches(x), Inches(y), Inches(w), Inches(h))
    _no_style_effects(sh)
    _fill(sh, fill)
    _line(sh, WHITE, 0.75)
    label(s, x + w / 2 - 0.6, y + h + 0.03, 1.2, name, size=7.26, color=BLACK, bold=True)
    return _obst(s, sh)


import ixdeck as _ixdeck  # noqa: E402  ixdiag 를 먼저 불러도 ixdeck 이름공간에 들어가게 (맨 아래에서 한 번 더)
for _n in __all__:
    setattr(_ixdeck, _n, globals()[_n])
