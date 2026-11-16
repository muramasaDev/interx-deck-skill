"""기존 덱의 지정 장 글만 바꿔 사본 저장 (나머지 파트는 바이트 그대로)
사용: edit_deck.py 원본.pptx 출력.pptx --replace 장번호 "옛 글" "새 글" [장번호 "옛 글" "새 글" ...]
- 원본은 읽기만 (python-pptx 로 열지 않음). 출력 사본은 gate.py --deck 으로 검사 및 렌더"""
import argparse
import copy
import hashlib
import os
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ixdeck import _NUM_UNIT  # noqa: E402

sys.stdout.reconfigure(errors="replace")
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def slide_parts(z):
    """장 순서대로 파트 이름"""
    rels = {r.get("Id"): r.get("Target") for r in etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    return [t[1:] if t.startswith("/") else "ppt/" + t for t in (rels[e.get(RID)] for e in pres.iter(P + "sldId"))]


def replace_in(root, old, new):
    """문단 안 old 를 new 로 (글 조각과 줄바꿈 걸침 가능). 바꾼 횟수"""
    old = old.replace(" ", " ")
    new = _NUM_UNIT.sub(" ", new)   # 숫자 및 단위 줄갈림 방지 규칙 유지
    count = 0
    for p in root.iter(A + "p"):
        segs, pos = [], 0   # (요소, 시작, 끝). 줄바꿈은 공백 한 글자
        for e in p:
            if e.tag == A + "r":
                t = e.find(A + "t")
                n = len(t.text or "")
                segs.append((t, pos, pos + n))
                pos += n
            elif e.tag == A + "br":
                segs.append((e, pos, pos + 1))
                pos += 1
        full = "".join((s[0].text or "").replace(" ", " ") if s[0].tag == A + "t" else " " for s in segs)
        hits, i = [], full.find(old)
        while i >= 0:
            hits.append(i)
            i = full.find(old, i + len(old))
        for s0 in reversed(hits):   # 뒤에서부터 바꿔야 앞 위치가 안 밀림
            e0 = s0 + len(old)
            hit = [k for k, (_, a, b) in enumerate(segs) if a < e0 and b > s0]
            if segs[hit[0]][0].tag != A + "t":
                raise ValueError(f"'{old}' 이 줄바꿈에서 시작함: 앞뒤 글을 더해 지정")
            for k in hit:
                el, a, b = segs[k]
                if el.tag == A + "br":
                    p.remove(el)
                    continue
                keep = (el.text[:max(s0, a) - a], el.text[min(e0, b) - a:])
                el.text = keep[0] + (new if k == hit[0] else "") + keep[1]
                if not el.text:
                    p.remove(el.getparent())
            count += 1
    return count


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--replace", nargs="+", action="extend", required=True, metavar="장번호 옛글 새글")
    a = ap.parse_args()
    if len(a.replace) % 3:
        sys.exit("--replace 는 '장번호 옛 글 새 글' 세 개씩")
    if os.path.exists(a.out) and os.path.samefile(a.src, a.out):
        sys.exit("출력이 원본과 같음: 원본은 덮지 않음")
    jobs = [(a.replace[i], a.replace[i + 1], a.replace[i + 2]) for i in range(0, len(a.replace), 3)]
    with zipfile.ZipFile(a.src) as zin:
        parts = slide_parts(zin)
        roots, errs = {}, []
        for no, old, new in jobs:
            if not old:
                errs.append(f"{no}쪽: 옛 글이 빈 문자열 (바꿀 글을 지정)")
                continue
            if not no.isdigit() or not 1 <= int(no) <= len(parts):
                errs.append(f"장번호 '{no}' 이 1 ~ {len(parts)} 밖")
                continue
            name = parts[int(no) - 1]
            roots.setdefault(name, etree.fromstring(zin.read(name)))
            try:
                if not replace_in(roots[name], old, new):
                    errs.append(f"{no}쪽에 '{old}' 이 없음 (글 조각을 걸친 글은 문단 안 글 그대로 지정)")
            except ValueError as e:
                errs.append(f"{no}쪽: {e}")
        if errs:
            sys.exit("저장하지 않음:\n" + "\n".join(errs))
        tmp = a.out + ".tmp"
        with zipfile.ZipFile(tmp, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename in roots:
                    data = etree.tostring(roots[info.filename], xml_declaration=True, encoding="UTF-8", standalone=True)
                zout.writestr(copy.copy(info), data)   # 원본 ZipInfo 는 쓰기가 바꾸므로 복사
        os.replace(tmp, a.out)
        # 해시 대조: 지정하지 않은 파트는 모두 같아야 함
        with zipfile.ZipFile(a.out) as zo:
            sig = lambda z, n: hashlib.sha256(z.read(n)).hexdigest()[:8]
            changed = [i for i, n in enumerate(parts, 1) if sig(zin, n) != sig(zo, n)]
            other = [n for n in zin.namelist() if n not in parts and sig(zin, n) != sig(zo, n)]
            report = ", ".join(f"{i}쪽 ({sig(zin, parts[i - 1])} -> {sig(zo, parts[i - 1])})" for i in changed)
            n_other = len(zin.namelist()) - len(parts)
    print(f"출력: {a.out}")
    print(f"바뀐 장: {report or '없음'}")
    print(f"안 바뀐 장: {len(parts) - len(changed)}장 (장 XML 해시 같음), 그 밖 파트 {n_other}개 중 다른 것 {len(other)}개")
    print(f'다음: python -P {os.path.join(os.path.dirname(os.path.abspath(__file__)), "gate.py")} --deck "{a.out}"  '
          "(검사, 렌더, 통과 기록. 출력 이름도 제목_작성자_v버전_YYMMDD 규칙)")
    sys.exit(1 if other or any(parts[i - 1] not in roots for i in changed) else 0)


if __name__ == "__main__":
    main()
