"""문장 검사: AI 투 표현, 서술체, 읽기 어려운 문장
사용: check_text.py 덱.pptx (종료 코드 = 실패 건수, 주의는 출력만)"""
import json
import os
import re
import sys
from collections import Counter

from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))
from check_ban import allowed, notes_of, slide_texts  # noqa: E402

LEXICON = os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))), "third_party", "im-not-ai",
                       "lexicon.json")

# 실패: 보고서에 쓰지 않는 AI 투 수식어, 감탄, 서술체
FAIL = [
    (r"혁신적|획기적|극대화|시너지|원활한|다양한|효과적으로|체계적으로|차별화된|손쉽게|놀라운"
     r"|원스톱|게임 체인저|한 단계 도약|새로운 차원|핵심적인 역할|중요한 역할", "AI 투 수식어: 구체 사실이나 수치로"),
    (r"[가-힣]니다|(한다|된다|이다|했다|하였다)(\.|$)", "서술체: 개조식 명사형 (~함, ~임, 명사)"),
    (r"[가-힣)]\s*!|[\U0001F300-\U0001FAFF]", "느낌표, 그림 문자"),
]
NOTE_FAIL = FAIL[:1] + FAIL[2:]   # 노트는 말투라 서술체 제외
# 주의: 사람이 읽고 판단
WARN = [
    (r"최적의|완벽하|강력한|한눈에", "꾸밈말: 수치나 사실로 바꿀 수 있는지"),
    (r"(을|를) 통해", "'~을 통해' 남용: '~로', '~에서'"),
    (r"살펴보|알아보|짚어보", "안내 말투: 내용만"),
    (r"(이|가) (말해|알려|보여) ?줌|가리킴|삼킴|고름\b|묻는다|답한다", "사물 의인화: 사람이 한 일로"),
    (r"→", "글 안 화살표: 연결어나 그림으로"),
    # 회사 방법론 용어 (skills/design/references/A1_방법론.md)
    (r"목표 ?변수|타[깃겟] ?변수|종속 ?변수", "방법론 용어: 결과(목적)변수 (Yn)"),
    (r"(데이터 )?품질 ?점수", "방법론 용어: 품질 지수"),
    (r"성숙도 ?(레벨|단계|수준)", "방법론 용어: Level 1 ~ 5 (모니터링, 제어, 최적화, 자율운영)"),
    (r"설정 ?변수|세팅 ?변수", "방법론 용어: 셋팅변수 (X3)"),
]
JOSA = ("에서", "으로", "에게", "까지", "부터", "처럼", "보다", "이며", "이고", "이나", "에는", "에도", "로서", "로써")
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


def lexicon_entry(e):
    """바꿀 항목이면 튜플, 필드 빠짐이나 정규식 오류면 None"""
    try:
        if e.get("tier") != "change" or not e.get("regex"):
            return None
        why = e["change_when"]["보고서"]
        return (e["id"], e["headword"], re.compile(e["regex"]),
                re.compile(e["regex_exclude"]) if e.get("regex_exclude") else None, why)
    except (KeyError, TypeError, AttributeError, re.error):
        return None


def load_lexicon():
    """im-not-ai 어휘 중 바꿀 항목 (필드가 빠진 항목은 건너뜀)"""
    return [x for x in map(lexicon_entry, json.load(open(LEXICON, encoding="utf8"))["entries"]) if x]


# 날짜 표기 5종 (YYYY.MM.DD 가 기준 꼴)
DATE_KINDS = {"YYYY.MM.DD": re.compile(r"(?<![\d.])(?:19|20)\d{2}\.\d{1,2}\.\d{1,2}(?![\d])"),
              "YYYY-MM-DD": re.compile(r"(?<![\d-])(?:19|20)\d{2}-\d{1,2}-\d{1,2}(?![\d])"),
              "M/D": re.compile(r"(?<![\d/.])(?:1[0-2]|0?[1-9])/(?:3[01]|[12]\d|0?[1-9])(?![\d/])"),
              "M월 D일": re.compile(r"\d{1,2}월\s?\d{1,2}일"),
              "YYYY.MM": re.compile(r"(?<![\d.])(?:19|20)\d{2}\.(?:0?[1-9]|1[0-2])(?!\.?\d)")}


def date_kinds(lines):
    """글에 나온 날짜 표기 종류 (이름: 첫 예)"""
    out = {}
    for ln in lines:
        for k, rx in DATE_KINDS.items():
            m = rx.search(ln)
            if m and k not in out:
                out[k] = m.group()
    return out


def lexicon_hits(where, ln, lex, seen):
    """한 줄에서 맞은 어휘 (장마다 항목 하나는 한 번만)"""
    for id_, head, rx, ex, why in lex:
        if (where, id_) not in seen and rx.search(ln) and not (ex and ex.search(ln)):
            seen.add((where, id_))
            yield f"{where} 어휘 {id_} {head}: {why} | {ln[:40]}"


def check(path):
    fails, warns = [], []
    prs = Presentation(path)
    ok = allowed(prs)
    long_ln = 110 if (prs.core_properties.category or "") == "ix-purpose:문서" else 70   # 문서 덱은 긴 줄 허용
    lex, seen = (load_lexicon(), set()) if os.path.exists(LEXICON) else ([], set())
    if not lex:
        warns.append(f"어휘 사전을 못 읽음: {LEXICON}")
    all_lines, src_by = [], []
    for i, s in enumerate(prs.slides, 1):
        lines = [ln.strip(" •※-") for t in slide_texts(s) for ln in (ok(t) if ok else t).split("\n")
                 if ln.strip()]
        all_lines += lines
        src_by.append({ln for ln in lines if ln.startswith("출처")})
        if i == 1:
            for ln in lines:
                if re.fullmatch(r"(?:[가-힣]+ ){0,2}(설계서|정의서|보고서|계획서|제안서|명세서)", ln):
                    warns.append(f"표지 제목 '{ln}' 에 제품이나 과제 이름이 없음: 받는 사람이 어느 문서인지 표지에서 알게")
        for ln in lines:
            cap = re.fullmatch(r"\[(?:표|그림) (?:#|\d+)\. (.+)\]", ln)
            if cap and cap.group(1) in lines:
                warns.append(f"{i}쪽 '{cap.group(1)}' 가 패널 제목과 표 제목에 두 번 나옴: 한쪽을 다른 말로")
            for m in re.finditer(r"[A-Za-z]{3,}[가-힣]{2,}", ln) if len(ln) <= 30 else ():
                if not re.sub(r"^[A-Za-z]+", "", m.group()).startswith(JOSA):   # 에서, 에서의, 로서의 같은 조사로 시작하면 정상
                    warns.append(f"{i}쪽 '{m.group()}' 영문 뒤 한글을 붙여 씀: 띄어쓰기 ('{ln[:30]}')")
        for ln in lines:
            for pat, why in FAIL:
                m = re.search(pat, ln)
                if m:
                    fails.append(f"{i}쪽 '{m.group()}' ({why}): {ln[:60]}")
            for pat, why in WARN:
                m = re.search(pat, ln)
                if m:
                    warns.append(f"{i}쪽 '{m.group()}' ({why}): {ln[:60]}")
            warns.extend(lexicon_hits(f"{i}쪽", ln, lex, seen))
            c = noun_chain(ln) if len(ln.split()) >= 4 else None   # 상자 이름은 뺌
            if c:
                warns.append(f"{i}쪽 명사 잇기 '{c}': 조사를 넣어 풀어 씀")
            if len(ln) > long_ln:
                warns.append(f"{i}쪽 {len(ln)}자 문장: 두 줄로 나누거나 줄임: {ln[:40]}...")
        note = notes_of(s)
        for ln in (x.strip(" •※-") for x in (ok(note) if ok else note).split("\n") if x.strip()):
            for pats, out in ((NOTE_FAIL, fails), (WARN, warns)):
                for pat, why in pats:
                    m = re.search(pat, ln)
                    if m:
                        out.append(f"{i}쪽 노트 '{m.group()}' ({why}): {ln[:60]}")
            warns.extend(lexicon_hits(f"{i}쪽 노트", ln, lex, seen))
        heads = [re.match(r"\S+", ln).group() for ln in lines if len(ln) > 8]
        for h in set(heads):
            if heads.count(h) >= 3 and re.fullmatch(r"[가-힣]+", h):
                warns.append(f"{i}쪽 '{h}' 로 시작하는 줄 {heads.count(h)}개: 같은 틀 되풀이")
    rep = Counter(x for sset in src_by for x in sset)
    for ln, n in rep.items():
        if (prs.core_properties.category or "") == "ix-purpose:문서" and n >= 5 and n >= 0.4 * len(prs.slides):
            warns.append(f"같은 출처 문구 {n}장 반복 ('{ln[:30]}'): 문서 정보 장에 한 번만 적고 장마다는 그 장의 근거만")
    kinds = date_kinds(all_lines)
    if len(kinds) > 2:
        warns.append(f"날짜 표기 {len(kinds)}종 혼용 ({', '.join(f'{k} {v}' for k, v in kinds.items())}): "
                     "YYYY.MM.DD 하나로")
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
