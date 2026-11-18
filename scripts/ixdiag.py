"""구성도 역할 부품 (skills/design/references/구성도_레퍼런스_문법.md): 역할 색 상자, 그라데이션 블록 화살표,
프로토콜 알약, 하단 약어 풀이 영역. ixdeck.py 맨 아래에서 불러 와 같은 이름공간에서 씀"""
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from ixdeck import (BLACK, GRAY_L, GREEN, NAVY, SLATE0, SLATE3, SLATE6, WHITE, _adj, _fill, _font, _lh, _line,
                    _no_style_effects, _obst, _reg, _set_text, _text_w, big_arrow, box, text)

__all__ = ["AMBER", "ROLE", "role_box", "block_arrow", "pill", "abbr_legend", "step_header", "summary_bar", "speech_bubble", "headline",
           "backdrop"]

AMBER = "FFC000"
ROLE = {"base": (SLATE6, WHITE), "ai": (GREEN, WHITE), "ext": (AMBER, BLACK), "platform": (NAVY, WHITE)}   # 칩 색의 뜻 (바탕, 글자)


def role_box(s, x, y, w, h, t, role="base", sub=None, size=7.26):
    """역할 색 칩. role: base 기존 모듈(회색), ai AI 및 에이전트(초록), ext 확장 항목(노랑), platform 핵심 서버(남색). sub: 작은 부제"""
    if role not in ROLE:
        raise ValueError(f"role_box role '{role}' 은 {', '.join(ROLE)} 중 하나")
    fill, fg = ROLE[role]
    sh = box(s, x, y, w, h, [(t, fg)] + ([(sub, fg)] if sub else []), fill=fill, size=size, bold=True, color=fg)
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


import ixdeck as _ixdeck  # noqa: E402  ixdiag 를 먼저 불러도 ixdeck 이름공간에 들어가게 (맨 아래에서 한 번 더)
for _n in __all__:
    setattr(_ixdeck, _n, globals()[_n])
