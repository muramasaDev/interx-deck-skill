#!/usr/bin/env python3
"""OKR 덱 찾기, 읽기, 근거 수집 (설정 불필요)"""
import datetime as dt
import glob
import json
import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import okr_deck as od  # noqa: E402
from pptx import Presentation  # noqa: E402

SKIP = {"node_modules", ".venv", "venv", ".git", "__pycache__", ".pytest_cache", "dist", "build", ".idea", ".vscode",
        "target", ".gradle", ".next"}


# --- find
def deck_period(path):
    try:
        prs = Presentation(path)
    except Exception:
        return None
    ps = [od.period_of(s) for s in prs.slides if od.board_table(s) is not None and od.period_of(s)]
    return max(ps) if ps else None


def cmd_find(folders):
    hits = []
    for root in folders:
        for p in glob.glob(os.path.join(os.path.expanduser(root), "**", "*.pptx"), recursive=True):
            name = os.path.basename(p)
            if name.startswith("~$"):
                continue
            if not re.search(r"OKR|Individual|okr", name):
                continue
            per = deck_period(p)
            if per:
                hits.append((per, os.path.getmtime(p), p))
    hits.sort(reverse=True)
    for (y, m), mt, p in hits:
        print(f"{y}-{m:02d}  {dt.datetime.fromtimestamp(mt):%Y-%m-%d}  {p}")
    if not hits:
        print("OKR 덱을 찾지 못함 → 첫 달이면 번들 템플릿(OKR_template.pptx)으로 시작")


# --- read
def _lines(cell):
    return [p.text.strip() for p in cell.text_frame.paragraphs if p.text.strip()]


def _rows(slide):
    t = od.board_table(slide)
    rows, cur_item = [], None
    for r in range(1, len(t.rows)):
        c = [t.cell(r, k).text.strip() for k in range(len(t.columns))]
        if not c[2]:
            continue
        row = {"kr": c[2], "w": c[3:7], "deliverable": c[7], "ach": c[8]}
        if c[1] and c[1] != cur_item:
            row["item"] = c[1]
            cur_item = c[1]
        rows.append(row)
    return rows


def _author(prs):
    for s in list(prs.slides)[:1]:
        for sh in s.shapes:
            if sh.has_text_frame:
                for ln in sh.text_frame.text.splitlines():
                    m = re.match(r"^\s*(.+?)\s*/\s*(\S+)\s*(.*)$", ln)
                    if m and "【입력】" not in ln and "Industrial" not in ln:
                        return {"team": m.group(1), "name": m.group(2), "title": m.group(3).strip()}
    return None


def cmd_read(path, out=None):
    prs = Presentation(path)
    boards = od.classify(prs)
    pre = {"_source": path, "author": _author(prs), "current": {}, "next": {}, "appendix": []}
    for kind, (old, newer) in boards.items():
        y, m = od.period_of(newer)
        q = (m - 1) // 3 + 1
        pre["current"]["period"] = f"{y}년 Q{q} / {m}월"
        ny, nm = (y + (m == 12), m % 12 + 1)
        pre["next"]["period"] = f"{ny}년 Q{(nm - 1) // 3 + 1} / {nm}월"
        plan = _lines(od.table_by_head(newer, "Monthly Plan").cell(1, 0))
        rows = _rows(newer)
        pre["current"][kind] = {
            "_plan_carried": plan,          # 지난달 덱의 익월 계획 = 이번 달 계획
            "review": ["【확인 필요】 계획 항목 순서대로 실적"],
            "rows": [dict(r, ach="【확인 필요】") for r in rows]}
        pre["next"][kind] = {
            "plan": ["【확인 필요】 익월 계획 (못 한 항목은 '(N월 이월)' 표시로 다시 올림)"],
            "rows": [{k: v for k, v in r.items() if k != "ach"} for r in rows]}
        pre["_prev_" + kind] = {"period": f"{od.period_of(old)[0]}-{od.period_of(old)[1]:02d}",
                               "review": _lines(od.table_by_head(old, "Monthly Review").cell(1, 0)),
                               "ach": [r["ach"] for r in _rows(old)]}
    txt = json.dumps(pre, ensure_ascii=False, indent=1)
    if out:
        open(out, "w", encoding="utf-8").write(txt)
        print("spec 초안 →", out)
    print(txt)


# --- evidence
def git(repo, *args):
    r = subprocess.run(["git", "-C", repo, "-c", "core.quotepath=off", *args], capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return r.stdout.rstrip()


def find_repos(root, depth=4):
    out = []
    root = os.path.expanduser(root)
    for dirpath, dirnames, _ in os.walk(root):
        if dirpath[len(root):].count(os.sep) >= depth:
            dirnames[:] = []
            continue
        if ".git" in dirnames or os.path.isfile(os.path.join(dirpath, ".git")):
            out.append(dirpath)          # 안쪽 저장소가 더 있을 수 있어 계속 내려감
        dirnames[:] = [d for d in dirnames if d not in SKIP and not d.startswith(".")]
    return out


def month_range(ym):
    y, m = map(int, ym.split("-"))
    return dt.date(y, m, 1), dt.date(y + (m == 12), m % 12 + 1, 1)


def cmd_evidence(ym, folders, author=None, out=None):
    start, end = month_range(ym)
    lines = [f"# {ym} 근거 (커밋·파일 시각은 단서, 수치는 문서에서 재확인)\n"]
    seen = set()
    for root in folders:
        repos = find_repos(root)
        for repo in repos:
            seen.add(os.path.abspath(repo))
            who = author
            if not who:
                parts = [git(repo, "config", "user.email"), git(repo, "config", "user.name")]
                who = "|".join(re.escape(p) for p in parts if p) or "."
            log = git(repo, "log", "--all", f"--since={start}", f"--until={end}", "-i", "-E", f"--author={who}",
                      "--format=%h %ad %s", "--date=short").splitlines()
            merges = git(repo, "log", f"--since={start}", f"--until={end}", "--merges", "--format=%h").splitlines()
            last = git(repo, "log", "-1", "--format=%ad", "--date=short")
            remote = bool(git(repo, "remote"))
            dirty = len([s for s in git(repo, "status", "--porcelain").splitlines() if s.strip()])
            lines.append(f"## {os.path.relpath(repo, os.path.expanduser(root))}  (git)")
            lines.append(f"- 내 커밋 {len(log)}건 · 머지 {len(merges)}건 · 마지막 커밋 {last or '-'} · 원격 {'있음' if remote else '없음'}"
                         f"{' · 커밋 안 된 변경 ' + str(dirty) + '건' if dirty else ''}")
            for l_ in log[:60]:
                lines.append(f"  - {l_}")
            lines.append("")
        # 그 달 수정된 파일 (커밋 안 하는 작업 포함)
        files = []
        for dirpath, dirnames, filenames in os.walk(os.path.expanduser(root)):
            dirnames[:] = [d for d in dirnames if d not in SKIP and not d.startswith(".")]
            for f in filenames:
                p = os.path.join(dirpath, f)
                try:
                    t = dt.date.fromtimestamp(os.path.getmtime(p))
                except OSError:
                    continue
                if start <= t < end and not f.startswith("~$"):
                    files.append((t, p))
        if files:
            files.sort(reverse=True)
            base = os.path.expanduser(root)
            groups = {}
            for t, p in files:
                parts = os.path.relpath(p, base).split(os.sep)
                key = os.sep.join(parts[:2]) if len(parts) > 2 else parts[0]
                g = groups.setdefault(key, [0, t, t, []])
                g[0] += 1
                g[1], g[2] = min(g[1], t), max(g[2], t)
                if len(g[3]) < 6:
                    g[3].append(parts[-1])
            lines.append(f"## {root}  (그 달 수정 파일 {len(files)}건 · 폴더별)")
            for key, (n, t0, t1, names) in sorted(groups.items(), key=lambda kv: -kv[1][0]):
                lines.append(f"- {key}  {n}건  ({t0:%m/%d}~{t1:%m/%d})  예: {', '.join(names)}")
            lines.append("")
    txt = "\n".join(lines)
    if out:
        open(out, "w", encoding="utf-8").write(txt)
    print(txt)


if __name__ == "__main__":
    a = sys.argv[1:]
    if not a:
        print(__doc__)
        sys.exit(1)

    def opt(k):
        if k in a:
            i = a.index(k)
            v = a[i + 1]
            del a[i:i + 2]
            return v
        return None

    out = opt("--out")
    author = opt("--author")
    if a[0] == "find":
        cmd_find(a[1:])
    elif a[0] == "read":
        cmd_read(a[1], out)
    elif a[0] == "evidence":
        cmd_evidence(a[1], a[2:], author, out)
    else:
        print(__doc__)
        sys.exit(1)
