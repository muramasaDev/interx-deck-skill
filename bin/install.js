#!/usr/bin/env node
// npx 설치 명령: install | uninstall | --version
// 개인 설정(~/.claude/interx-deck/profile)은 건드리지 않음
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execSync } = require("child_process");

const SRC = path.resolve(__dirname, "..");
const PKG = require(path.join(SRC, "package.json"));
const CLAUDE = process.env.CLAUDE_CONFIG_DIR || path.join(os.homedir(), ".claude");
const DEST = path.join(CLAUDE, "skills", "interx-deck");
const MARK = ".installed-by-npm";
const ITEMS = [".claude-plugin", "skills", "scripts", "assets", "README.md", "CHANGELOG.md"];

const COLOR = (process.stdout.isTTY || !!process.env.FORCE_COLOR) && !process.env.NO_COLOR;
const paint = (code) => (s) => (COLOR ? `\x1b[${code}m${s}\x1b[0m` : s);
const rgb = (c, extra = "") => paint(`${extra}38;2;${c.join(";")}`);
const mix = (a, b, t) => a.map((v, i) => Math.round(v + (b[i] - v) * t));
const dim = paint("90");
const bold = paint("1;97");
const ORANGE = [234, 88, 11];
const sleep = (ms) => COLOR && Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);

// 글자 폭: 한글, 전각은 2칸
const HANGUL = /[ᄀ-ᇿ㄰-㆏가-힣一-鿿＀-｠]/;
const width = (s) => [...s.replace(/\x1b\[[0-9;]*m/g, "")].reduce((n, ch) => n + (HANGUL.test(ch) ? 2 : 1), 0);
const pad = (s, n) => s + " ".repeat(Math.max(0, n - width(s)));
const grad = (s, a, b) => [...s].map((ch, i, all) => rgb(mix(a, b, i / Math.max(1, all.length - 1)), "1;")(ch)).join("");

// 村正: 도는 카타나 (반 칸 블록, 정사각 픽셀)
const R = 21;                  // 회전 반지름 (픽셀)
const N = 2 * R + 1;           // 그림 한 변: N 칸, N/2 줄
const W = 46;                  // 정보 카드 폭
const STEEL = [150, 162, 178];
const EDGE = [232, 238, 246];
const GOLD = [212, 172, 60];
const ITO = [28, 28, 38];      // 손잡이 끈
const SAME = [196, 186, 160];  // 손잡이 상어 가죽
const RED = [200, 30, 40];

// 칼 좌표 (u 칼 방향, v 수직)의 색, 없으면 null
// - 칼끝 u=15, 손잡이 끝 u=-11 (날 약 2.6배)
function part(u, v, glint) {
  if (u >= -2.5 && u <= 15) {                   // 칼날, 휨(소리)과 칼끝 좁힘
    const vb = v - 1.5 * ((u + 2.5) / 17.5) ** 2;
    const half = u > 12.5 ? 1.15 * (15 - u) / 2.5 : 1.15;
    if (Math.abs(vb) > half) return null;
    if (Math.abs(u - glint) < 1.2) return [255, 255, 255];
    return vb < 0 ? EDGE : STEEL;
  }
  if (u >= -3.3 && u < -2.5) return Math.abs(v) <= 1.4 ? GOLD : null;        // 하바키
  if (u >= -4.3 && u < -3.3) return Math.abs(v) <= 3 ? [92, 74, 36] : null;  // 츠바
  if (u >= -11 && u < -4.3) {                   // 손잡이, 마름모 무늬
    if (Math.abs(v) > 1.35) return null;
    if (u < -10.3) return GOLD;                 // 카시라
    return Math.abs(((u * 1.2) % 2 + 2) % 2 - 1) > Math.abs(v) / 1.6 ? ITO : SAME;
  }
  return null;
}

// 각도 a 의 한 장면: 줄 목록
function frame(a, glint) {
  const c = Math.cos(a), s = Math.sin(a);
  const k = 15.5 / R;            // 칼 좌표 단위로 맞춤
  const px = (x, y) => part((x * c + y * s) * k, (-x * s + y * c) * k, glint);
  const lines = [];
  for (let r = 0; r < N; r += 2) {
    let line = "  ";
    for (let k = 0; k < N; k++) {
      const x = k - R, top = px(x, R - r), bot = r + 1 < N ? px(x, R - r - 1) : null;
      if (!COLOR) line += top && bot ? "█" : top ? "▀" : bot ? "▄" : " ";
      else if (top && bot) line += `\x1b[38;2;${top.join(";")};48;2;${bot.join(";")}m▀\x1b[0m`;
      else if (top) line += `\x1b[38;2;${top.join(";")}m▀\x1b[0m`;
      else if (bot) line += `\x1b[38;2;${bot.join(";")}m▄\x1b[0m`;
      else line += " ";
    }
    lines.push(line);
  }
  return lines;
}

function banner() {
  const end = Math.PI / 9;                       // 멈춘 자세: 칼끝 위로 20도
  const steps = COLOR ? 40 : 1;
  console.log();
  if (COLOR) process.stdout.write("\x1b[?25l"); // 커서 숨김
  for (let i = 1; i <= steps; i++) {
    const t = i / steps, ease = 1 - (1 - t) ** 3;
    const lines = frame(end - 4 * Math.PI * (1 - ease), -3 + 21 * t);
    if (i < steps) process.stdout.write(lines.join("\n") + `\n\x1b[${lines.length}A`);
    else process.stdout.write((COLOR ? "\x1b[J" : "") + lines.filter((l) => l.trim()).join("\n") + "\n\n");  // 빈 줄 제거
    sleep(28);
  }
  if (COLOR) process.stdout.write("\x1b[?25h");
  console.log(`  ${grad("村正", RED, [120, 10, 20])}  ${bold("M U R A M A S A")}  ${dim("伊勢国 桑名")}`);
  const head = `${dim("─ ")}${bold("interx-deck")} `;
  const ver = ` ${rgb(ORANGE, "1;")("v" + PKG.version)}${dim(" ─")}`;
  const row = (s) => console.log(`  ${dim("│")} ${pad(s, W - 4)} ${dim("│")}`);
  console.log();
  console.log(`  ${dim("╭")}${head}${dim("─".repeat(Math.max(0, W - 2 - width(head) - width(ver))))}${ver}${dim("╮")}`);
  row(`${dim("PPT engine for")} ${bold("Claude Code")}`);
  row(`${dim("crafted by")}  ${grad("Changhun Jeong", [255, 190, 90], [236, 72, 153])}`);
  console.log(`  ${dim("╰" + "─".repeat(W - 2) + "╯")}`);
  console.log();
}

// 단계 한 줄: ◆ 이름 ······ 결과
function step(label, good, detail = "") {
  const mark = good ? rgb([34, 197, 94], "1;")("✔") : rgb(ORANGE, "1;")("!");
  const dots = dim("·".repeat(Math.max(2, 20 - width(label))));
  console.log(`  ${rgb(ORANGE)("◆")} ${label} ${dots} ${mark} ${good ? detail : rgb(ORANGE)(detail)}`);
}
const busy = (s) => console.log(`  ${dim("◌ " + s)}`);

function fail(msg) {
  console.error(`  ${rgb([239, 68, 68], "1;")("✖")} ${msg}`);
  process.exit(1);
}

// .git 이 있는 원본 작업 폴더, 손으로 둔 폴더는 덮지 않음
function ownedByUs(dir) {
  if (!fs.existsSync(dir)) return true;
  if (fs.existsSync(path.join(dir, ".git"))) return false;
  return fs.existsSync(path.join(dir, MARK));
}

function install() {
  if (!ownedByUs(DEST)) fail(`${DEST} 가 이미 있음 (npx 로 설치한 폴더 아님). 옮긴 뒤 다시 실행`);
  const markFile = path.join(DEST, MARK);
  const old = fs.existsSync(markFile) ? fs.readFileSync(markFile, "utf8").trim() : null;
  const tmp = DEST + ".new";
  fs.rmSync(tmp, { recursive: true, force: true });
  for (const it of ITEMS) {
    const s = path.join(SRC, it);
    if (fs.existsSync(s)) fs.cpSync(s, path.join(tmp, it), { recursive: true });
  }
  fs.writeFileSync(path.join(tmp, MARK), PKG.version + "\n");
  fs.rmSync(DEST, { recursive: true, force: true });
  fs.renameSync(tmp, DEST);
  step("플러그인", true, old ? `업데이트 ${old} → ${PKG.version}` : `설치 ${PKG.version}`);
  console.log(`    ${dim("└ " + DEST)}`);
  checks();
  console.log();
  console.log(`  ${grad("❯❯", [255, 190, 90], ORANGE)} Claude Code 를 다시 시작하고 ${bold("/interx-deck:using")}`);
  console.log();
}

function uninstall() {
  if (!fs.existsSync(DEST)) return step("플러그인", true, "설치된 것이 없음");
  if (!ownedByUs(DEST)) fail(`${DEST} 는 npx 로 설치한 폴더가 아니라 지우지 않음`);
  fs.rmSync(DEST, { recursive: true, force: true });
  step("플러그인", true, "제거 (개인 설정은 남김)");
}

const PY_PKGS = "python-pptx lxml pillow pdfplumber";
const run = (c) => {
  try {
    execSync(c, { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
};
const WIN = process.platform === "win32";
const MAC = process.platform === "darwin";
const findPython = () => (WIN ? ["python", "py -3"] : ["python3", "python3.13", "python3.12", "python3.11", "python"])
  .find((py) => run(`${py} -c "import sys; assert sys.version_info >= (3, 11)"`));

// Python, 패키지가 없으면 설치 (사용자 범위, 관리자 권한 불필요)
function ensurePython() {
  let py = findPython();
  if (!py && WIN) {
    busy("Python 설치 중 (winget)");
    run("winget install -e --id Python.Python.3.12 --scope user --silent " +
        "--accept-package-agreements --accept-source-agreements");
    const exe = path.join(os.homedir(), "AppData", "Local", "Programs", "Python", "Python312", "python.exe");
    py = findPython() || (fs.existsSync(exe) ? `"${exe}"` : null);
  }
  if (!py && MAC && run("brew --version")) {
    busy("Python 설치 중 (Homebrew)");
    run("brew install python@3.12");
    py = findPython();
  }
  if (!py) return step("Python", false, "설치 실패: python.org 에서 3.11 이상 설치 후 다시 실행");
  const has = () => run(`${py} -c "import pptx, lxml, PIL, pdfplumber"`);
  if (!has()) {
    busy("Python 패키지 설치 중");
    // 쓰기 불가면 pip 이 사용자 범위로 설치
    run(`${py} -m pip install -q ${PY_PKGS}`) ||
      run(`${py} -m pip install -q --break-system-packages ${PY_PKGS}`);   // Homebrew Python (PEP 668)
  }
  if (has()) step("Python", true, "패키지 준비됨");
  else step("Python", false, `패키지 설치 실패: ${py} -m pip install ${PY_PKGS}`);
}

// Windows 는 파일이 아니라 등록 여부로 판단 (등록 안 된 파일은 안 보임)
const MAC_FONTS = [path.join(os.homedir(), "Library", "Fonts"), "/Library/Fonts"];
const hasFont = () => (WIN
  ? ["HKLM", "HKCU"].some((root) =>
    run(`reg query "${root}\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts" /f "Noto Sans KR"`))
  : MAC_FONTS.some((d) => fs.existsSync(d) && fs.readdirSync(d).some((f) => /^NotoSansKR/i.test(f))));

// 관리자 권한 없이 사용자 글꼴로 설치 (OFL)
function installFont() {
  const file = "NotoSansKR-VF.ttf";
  const dir = WIN ? path.join(os.homedir(), "AppData", "Local", "Microsoft", "Windows", "Fonts") : MAC_FONTS[0];
  try {
    fs.mkdirSync(dir, { recursive: true });
    fs.copyFileSync(path.join(SRC, "assets", "fonts", file), path.join(dir, file));
    if (WIN) execSync(`reg add "HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts" /v "Noto Sans KR (TrueType)" ` +
             `/t REG_SZ /d "${path.join(dir, file)}" /f`, { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

// 렌더 프로그램 확인
function renderer() {
  if (WIN) return step("렌더", true, "PowerPoint");
  if (MAC && fs.existsSync("/Applications/Microsoft PowerPoint.app")) return step("렌더", true, "PowerPoint for Mac");
  if (run("soffice --version") || fs.existsSync("/Applications/LibreOffice.app")) {
    return step("렌더", true, "LibreOffice (PowerPoint 와 줄 바꿈이 다를 수 있음)");
  }
  step("렌더", false, "PowerPoint 또는 LibreOffice 필요 (렌더 확인과 PDF)");
}

function checks() {
  ensurePython();
  if (!WIN && !MAC) return step("OS", false, "Windows, Mac 만 지원");
  if (hasFont()) step("글꼴", true, "Noto Sans KR");
  else if (installFont()) step("글꼴", true, "Noto Sans KR 설치 (안 보이면 다시 로그인)");
  else step("글꼴", false, "Noto Sans KR 설치 실패");
  renderer();
}

const cmd = process.argv[2] || "install";
if (cmd === "--version" || cmd === "-v") {
  console.log(PKG.version);
} else if (cmd === "install" || cmd === "uninstall") {
  banner();
  cmd === "install" ? install() : uninstall();
} else {
  fail("사용: npx --allow-git=all muramasaDev/interx-deck-skill [install|uninstall|--version]");
}
