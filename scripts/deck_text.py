"""덱 글을 장마다 읽는 순서로 뽑음 (검토용)
사용: deck_text.py 덱.pptx [출력.md]"""
import os
import sys

from pptx import Presentation
from pptx.util import Emu

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from check_ban import slide_texts, texts  # noqa: E402


def dump(path):
    out = []
    for i, s in enumerate(Presentation(path).slides, 1):
        out.append(f"## {i}쪽")
        shapes = sorted(s.shapes, key=lambda sh: (round(Emu(sh.top or 0).inches * 4), sh.left or 0))
        seen = set()
        for sh in shapes:
            for t in texts(sh):
                t = t.strip()
                if t and t not in seen:
                    seen.add(t)
                    out.append("- " + t.replace("\n", " / "))
        for t in slide_texts(s):   # 수식 상자 글
            t = t.strip()
            if t and t not in seen:
                seen.add(t)
                out.append("- (수식) " + t.replace("\n", " / "))
        out.append("")
    return "\n".join(out)


if __name__ == "__main__":
    md = dump(sys.argv[1])
    if len(sys.argv) > 2:
        open(sys.argv[2], "w", encoding="utf8").write(md)
    else:
        sys.stdout.reconfigure(errors="replace")
        print(md)
