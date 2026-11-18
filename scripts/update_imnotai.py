"""im-not-ai 규칙 파일 갱신 (MIT, epoko77-ai/im-not-ai)
사용: update_imnotai.py [대상 폴더]
네트워크 실패 시 기존 파일 유지, 종료 코드 1"""
import json
import os
import sys
import urllib.request

sys.stdout.reconfigure(errors="replace")
REPO = "epoko77-ai/im-not-ai"
RAW = f"https://raw.githubusercontent.com/{REPO}/main/"
FILES = {"LICENSE": "LICENSE",
         "lexicon.json": "skills/humanize-korean/references/lexicon.json",
         "quick-rules.md": "skills/humanize-korean/references/quick-rules.md"}
DEST = os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))), "third_party", "im-not-ai")


def valid_lexicon(raw):
    """바꿀 항목 필드가 모두 있는지 확인. 아니면 ValueError"""
    es = json.loads(raw)["entries"]
    bad = [e.get("id", "?") for e in es if e.get("tier") == "change" and not (
        e.get("id") and e.get("headword") and e.get("regex") and isinstance(e.get("change_when"), dict)
        and e["change_when"].get("보고서"))]
    if bad or not any(e.get("tier") == "change" for e in es):
        raise ValueError(f"lexicon.json 형식이 다름: 필드 빠진 바꿀 항목 {bad[:5] or '또는 바꿀 항목 없음'}")


def get(url):
    with urllib.request.urlopen(url, timeout=30) as r:
        return r.read()


def main(dest=DEST):
    try:   # 다 받은 뒤에만 덮음
        got = {name: get(RAW + src) for name, src in FILES.items()}
        valid_lexicon(got["lexicon.json"])
        ver = json.loads(get(RAW + "plugin.json"))["version"]
        c = json.loads(get(f"https://api.github.com/repos/{REPO}/commits/main"))
    except Exception as e:   # 네트워크, 형식 오류 모두 기존 유지
        print(f"갱신 실패, 기존 파일 유지: {e}")
        return 1
    os.makedirs(dest, exist_ok=True)
    for name, data in got.items():
        open(os.path.join(dest, name), "wb").write(data)
    day = c["commit"]["committer"]["date"][:10]
    open(os.path.join(dest, "VERSION"), "w", encoding="utf8").write(f"{REPO} v{ver}\ncommit {c['sha'][:7]}\n{day}\n")
    print(f"갱신: {REPO} v{ver} commit {c['sha'][:7]} ({day})")
    return 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:2]))
