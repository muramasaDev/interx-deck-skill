"""문장 검사: AI 투 표현, 서술체, 읽기 어려운 문장
사용: check_text.py 덱.pptx (종료 코드 = 실패 건수, 주의는 출력만)"""
import os
import re
import sys

from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from check_ban import allowed, slide_texts  # noqa: E402

# 실패: 보고서에 쓰지 않는 AI 투 수식어, 감탄, 서술체
FAIL = [
    (r"혁신적|획기적|극대화|시너지|원활한|다양한|효과적으로|체계적으로|차별화된|손쉽게|놀라운"
     r"|원스톱|게임 체인저|한 단계 도약|새로운 차원|핵심적인 역할|중요한 역할", "AI 투 수식어: 구체 사실이나 수치로"),
    (r"[가-힣]니다|(한다|된다|이다|했다|하였다)(\.|$)", "서술체: 개조식 명사형 (~함, ~임, 명사)"),
    (r"[가-힣)]\s*!|[\U0001F300-\U0001FAFF]", "느낌표, 그림 문자"),
]
# 주의: 사람이 읽고 판단
WARN = [
    (r"최적의|완벽하|강력한|한눈에", "꾸밈말: 수치나 사실로 바꿀 수 있는지"),
    (r"(을|를) 통해", "'~을 통해' 남용: '~로', '~에서'"),
    (r"살펴보|알아보|짚어보", "안내 말투: 내용만"),
    (r"(이|가) (말해|알려|보여) ?줌|가리킴|삼킴|고름\b|묻는다|답한다", "사물 의인화: 사람이 한 일로"),
    (r"→", "글 안 화살표: 연결어나 그림으로"),
]
# 조사, 어미로 끝나는 낱말 (명사 잇기 판단용)
ENDS = re.compile(r"(은|는|이|가|을|를|의|에|에서|로|으로|와|과|도|만|까지|부터|보다|처럼|하여|하고|해서|하면|이면|고|며|면|서|지|게"
                  r"|음|함|임|됨|듦|및|등|중|후|전|별|당|용)$")


def noun_chain(t):
    """한글 낱말 넷 이상이 조사 없이 이어진 곳"""
    run = []
    for w in re.findall(r"\S+", t):
        if re.fullmatch(r"[가-힣]+", w) and not ENDS.search(w):
            run.append(w)
            if len(run) >= 4:
                return " ".join(run)
        else:
            run = []
    return None


def check(path):
    fails, warns = [], []
    prs = Presentation(path)
    ok = allowed(prs)
    for i, s in enumerate(prs.slides, 1):
        lines = [ln.strip(" •※-") for t in slide_texts(s) for ln in (ok(t) if ok else t).split("\n")
                 if ln.strip()]
        for ln in lines:
            for pat, why in FAIL:
                m = re.search(pat, ln)
                if m:
                    fails.append(f"{i}쪽 '{m.group()}' ({why}): {ln[:60]}")
            for pat, why in WARN:
                m = re.search(pat, ln)
                if m:
                    warns.append(f"{i}쪽 '{m.group()}' ({why}): {ln[:60]}")
            c = noun_chain(ln) if len(ln.split()) >= 4 else None   # 상자 이름은 뺌
            if c:
                warns.append(f"{i}쪽 명사 잇기 '{c}': 조사를 넣어 풀어 씀")
            if len(ln) > 70:
                warns.append(f"{i}쪽 {len(ln)}자 문장: 두 줄로 나누거나 줄임: {ln[:40]}...")
        heads = [re.match(r"\S+", ln).group() for ln in lines if len(ln) > 8]
        for h in set(heads):
            if heads.count(h) >= 3 and re.fullmatch(r"[가-힣]+", h):
                warns.append(f"{i}쪽 '{h}' 로 시작하는 줄 {heads.count(h)}개: 같은 틀 되풀이")
    return fails, warns


if __name__ == "__main__":
    sys.stdout.reconfigure(errors="replace")
    n = 0
    for f in sys.argv[1:]:
        fails, warns = check(f)
        for x in fails:
            print("실패 " + x)
        for x in warns:
            print("주의 " + x)
        n += len(fails)
    print(f"문장 실패 {n}")
    sys.exit(min(n, 255))
