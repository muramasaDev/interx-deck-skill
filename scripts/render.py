"""장마다 PNG, --pdf 면 PDF 추가 (사본만 열기)
사용: render.py pptx 폴더 [--slides 1,3 또는 none] [--pdf 파일]
Windows PowerPoint, Mac PowerPoint 또는 LibreOffice"""
import argparse
import glob
import locale
import os
import shutil
import subprocess
import sys
import tempfile
import uuid

HERE = os.path.dirname(os.path.realpath(__file__))
PNG_W = 1600   # render.ps1 과 같은 폭


def _font_dirs():
    if os.name == "nt":
        return [os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
                os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")]
    return [os.path.expanduser("~/Library/Fonts"), "/Library/Fonts"]


def _registry_fonts():
    """Windows 글꼴 등록 값 이름 (HKLM 및 HKCU)"""
    if os.name != "nt":
        return []
    import winreg
    out = []
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as k:
                i = 0
                while True:
                    try:
                        out.append(winreg.EnumValue(k, i)[0])
                    except OSError:
                        break
                    i += 1
        except OSError:
            pass
    return out


def _fc_families():
    """fc-list 글꼴 이름 (없으면 빈 글)"""
    fc = shutil.which("fc-list") if os.name != "nt" else None
    if not fc:
        return ""
    try:
        return subprocess.run([fc, ":", "family"], capture_output=True, timeout=30).stdout.decode("utf8", "replace")
    except (OSError, subprocess.TimeoutExpired):
        return ""


def font_installed():
    """Noto Sans KR 설치 여부 (없으면 렌더가 다른 글꼴로 줄 바꿈이 달라짐)
    - 글꼴 폴더 파일, Windows 등록 값, fc-list 순"""
    if any(glob.glob(os.path.join(d, "NotoSansKR*")) for d in _font_dirs()):
        return True
    return any("noto sans kr" in n.lower() for n in _registry_fonts()) or "noto sans kr" in _fc_families().lower()


def _windows(a):
    cmd = ["powershell", "-NoProfile", "-File", os.path.join(HERE, "render.ps1"), "-Pptx", a.pptx, "-Out", a.out]
    if a.slides:
        cmd += ["-Slides", a.slides]
    if a.pdf:
        cmd += ["-Pdf", a.pdf]
    r = subprocess.run(cmd, capture_output=True)
    sys.stdout.write((r.stdout + r.stderr).decode(locale.getpreferredencoding(False), "replace"))
    return r.returncode


def _mac_pdf(src, pdf):
    """사본을 PDF 로 저장
    - PowerPoint 샌드박스: 자기 폴더 안 파일만 묻지 않고 열기"""
    box = os.path.expanduser("~/Library/Containers/com.microsoft.Powerpoint/Data")
    if os.path.isdir("/Applications/Microsoft PowerPoint.app") and os.path.isdir(box):
        tid = uuid.uuid4().hex
        tmp, tpdf = os.path.join(box, f"ixrender_{tid}.pptx"), os.path.join(box, f"ixrender_{tid}.pdf")
        script = f'''tell application "Microsoft PowerPoint"
    open POSIX file "{tmp}"
    set p to active presentation
    save p in POSIX file "{tpdf}" as save as PDF
    close p saving no
    if (count of presentations) is 0 then quit
end tell'''
        try:
            shutil.copyfile(src, tmp)
            r = subprocess.run(["osascript", "-e", script], capture_output=True, timeout=300)
            if r.returncode == 0 and os.path.exists(tpdf):
                shutil.copyfile(tpdf, pdf)
                return "PowerPoint"
            sys.stderr.write("PowerPoint 렌더 실패, LibreOffice 로: " + r.stderr.decode("utf8", "replace"))
        except (OSError, subprocess.TimeoutExpired) as e:   # 컨테이너 접근 거부 등
            sys.stderr.write(f"PowerPoint 렌더 실패, LibreOffice 로: {e}\n")
        finally:
            for f in (tmp, tpdf):
                if os.path.exists(f):
                    os.remove(f)
    soffice = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not os.path.exists(soffice):
        sys.exit("렌더 실패: PowerPoint 로 못 열었고 LibreOffice 도 없음 (설치 후 다시)")
    tmpdir = tempfile.mkdtemp(prefix="ixrender_")
    try:
        tmp = os.path.join(tmpdir, "deck.pptx")
        shutil.copyfile(src, tmp)
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, tmp],
                       capture_output=True, timeout=300, check=True)
        shutil.copyfile(os.path.join(tmpdir, "deck.pdf"), pdf)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    return "LibreOffice (PowerPoint 와 줄 바꿈이 다를 수 있음)"


def _mac(a):
    import pypdfium2 as pdfium   # pdfplumber 와 함께 설치됨
    os.makedirs(a.out, exist_ok=True)
    pdf = a.pdf or os.path.join(tempfile.gettempdir(), f"ixrender_{uuid.uuid4().hex}.pdf")
    try:
        how = _mac_pdf(os.path.abspath(a.pptx), pdf)
        doc = pdfium.PdfDocument(pdf)
        pages = ([] if a.slides == "none" else [int(x) for x in a.slides.split(",")]) if a.slides else range(1, len(doc) + 1)
        for i in pages:
            page = doc[i - 1]
            png = os.path.join(a.out, f"s{i}.png")
            page.render(scale=PNG_W / page.get_width()).to_pil().save(png)
            print(png)
        doc.close()
        print(f"렌더: {how}", file=sys.stderr)
    finally:
        if not a.pdf and os.path.exists(pdf):
            os.remove(pdf)
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("out")
    ap.add_argument("--slides")
    ap.add_argument("--pdf")
    a = ap.parse_args()
    if not font_installed():
        sys.stderr.write("렌더 중단: Noto Sans KR 글꼴이 설치되어 있지 않음 (" + ", ".join(_font_dirs()) +
                         "). 렌더가 다른 글꼴로 줄 바꿈이 달라지므로 assets/fonts/NotoSansKR-VF.ttf 를 설치한 뒤 다시\n")
        sys.exit(3)
    if a.pdf:
        a.pdf = os.path.abspath(a.pdf)
    sys.exit(_windows(a) if os.name == "nt" else _mac(a))


if __name__ == "__main__":
    main()
