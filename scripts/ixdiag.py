"""구성도 역할 부품 (skills/design/references/구성도_레퍼런스_문법.md): 역할 색 상자, 그라데이션 블록 화살표,
프로토콜 알약, 하단 약어 풀이 영역. ixdeck.py 맨 아래에서 불러 와 같은 이름공간에서 씀"""
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from ixdeck import (_A, BLACK, GREEN, NAVY, SLATE3, SLATE6, WHITE, _adj, _lh, _line, _no_style_effects, _obst, _text_w,
                    box, etree, qn, text)

__all__ = ["AMBER", "ROLE", "role_box", "block_arrow", "pill", "abbr_legend"]

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
    rot = {"right": 0, "down": 90, "left": 180, "up": 270}[direction]
    ow, oh = (h, w) if rot in (90, 270) else (w, h)   # 세로는 돌리기 전 틀을 맞바꿔 (x, y, w, h) 가 보이는 크기가 되게
    sh = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + (w - ow) / 2), Inches(y + (h - oh) / 2), Inches(ow),
                            Inches(oh))
    _no_style_effects(sh)
    _adj(sh, 0.5, 0.6)
    sh.rotation = rot
    sppr = sh._element.spPr
    for el in sppr.findall(qn("a:solidFill")):
        sppr.remove(el)
    sppr.find(qn("a:prstGeom")).addnext(etree.fromstring(
        f'<a:gradFill xmlns:a="{_A}" flip="none" rotWithShape="1"><a:gsLst>'
        f'<a:gs pos="0"><a:srgbClr val="{SLATE3}"/></a:gs><a:gs pos="100000"><a:srgbClr val="{SLATE6}"/></a:gs>'
        '</a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>'))
    _line(sh, None)
    if label:
        lines = label if isinstance(label, list) else [label]
        th = len(lines) * _lh(8) + 0.04
        if direction in ("right", "left"):
            text(s, x, y - th - 0.02, w, th, lines, size=8, bold=True, align="c", anchor="b")
        else:
            text(s, x + w + 0.08, y + (h - th) / 2, 2.2, th, lines, size=8, bold=True, align="l")
    return sh


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
        text(s, x, y, lw, h, ["- " + b for b in bullets], size=8, color=BLACK, align="l", anchor="t")
    if abbrs:
        text(s, x + lw, y, w - lw, h, [f"{k} : {v}" for k, v in abbrs], size=8, bold=True, color=BLACK, align="r",
             anchor="t")
    return y + h


import ixdeck as _ixdeck  # noqa: E402  ixdiag 를 먼저 불러도 ixdeck 이름공간에 들어가게 (맨 아래에서 한 번 더)
for _n in __all__:
    setattr(_ixdeck, _n, globals()[_n])
