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

const COLOR = process.stdout.isTTY && !process.env.NO_COLOR;
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

// 5x5 픽셀 글꼴
const GLYPH = {
  I: ["███", " █ ", " █ ", " █ ", "███"],
  N: ["█   █", "██  █", "█ █ █", "█  ██", "█   █"],
  T: ["█████", "  █  ", "  █  ", "  █  ", "  █  "],
  E: ["████", "█   ", "███ ", "█   ", "████"],
  R: ["████ ", "█   █", "████ ", "█  █ ", "█   █"],
  X: ["█   █", " █ █ ", "  █  ", " █ █ ", "█   █"],
  D: ["████ ", "█   █", "█   █", "█   █", "████ "],
  C: [" ████", "█    ", "█    ", "█    ", " ████"],
  K: ["█   █", "█  █ ", "███  ", "█  █ ", "█   █"],
};
const word = (w, row) => [...w].map((ch) => GLYPH[ch][row]).join(" ");
const LEFT = word("INTERX", 0).length + 3;
const ROWS = [0, 1, 2, 3, 4].map((r) => [...word("INTERX", r) + "   " + word("DECK", r)]);
const W = ROWS[0].length + 1;

// 가로 그라데이션 면, 아래로 갈수록 어둡게
function face(r, c) {
  const t = c < LEFT ? c / LEFT : (c - LEFT) / (W - LEFT);
  const base = c >= LEFT ? mix([255, 255, 255], [125, 211, 252], t)
    : t < 0.5 ? mix([255, 190, 90], ORANGE, t * 2) : mix(ORANGE, [200, 50, 30], (t - 0.5) * 2);
  return mix(base, [0, 0, 0], r * 0.07);
}

function banner() {
  const on = (r, c) => r >= 0 && c >= 0 && ROWS[r] && ROWS[r][c] === "█";
  console.log();
  for (let r = 0; r < 6; r++) {
    let line = "  ";
    for (let c = 0; c < W; c++) {
      if (on(r, c)) line += rgb(face(r, c))("█");
      else if (on(r - 1, c - 1)) line += rgb([60, 60, 72])(COLOR ? "▓" : "░");
      else line += " ";
    }
    console.log(line);
    sleep(35);
  }
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
const findPython = () => ["python", "py -3"].find((py) => run(`${py} -c "import sys; assert sys.version_info >= (3, 11)"`));

// Python, 패키지가 없으면 설치 (사용자 범위, 관리자 권한 불필요)
function ensurePython() {
  let py = findPython();
  if (!py && process.platform === "win32") {
    busy("Python 설치 중 (winget)");
    run("winget install -e --id Python.Python.3.12 --scope user --silent " +
        "--accept-package-agreements --accept-source-agreements");
    const exe = path.join(os.homedir(), "AppData", "Local", "Programs", "Python", "Python312", "python.exe");
    py = findPython() || (fs.existsSync(exe) ? `"${exe}"` : null);
  }
  if (!py) return step("Python", false, "설치 실패: python.org 에서 설치 후 다시 실행");
  const has = () => run(`${py} -c "import pptx, lxml, PIL, pdfplumber"`);
  if (!has()) {
    busy("Python 패키지 설치 중");
    run(`${py} -m pip install --user -q ${PY_PKGS}`);
  }
  if (has()) step("Python", true, "패키지 준비됨");
  else step("Python", false, `패키지 설치 실패: ${py} -m pip install ${PY_PKGS}`);
}

// 파일이 아니라 등록 여부로 판단 (등록 안 된 파일은 안 보임)
const hasFont = () => ["HKLM", "HKCU"].some((root) =>
  run(`reg query "${root}\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts" /f "Noto Sans KR"`));

// 관리자 권한 없이 사용자 글꼴로 설치 (OFL)
function installFont() {
  const file = "NotoSansKR-VF.ttf";
  const dir = path.join(os.homedir(), "AppData", "Local", "Microsoft", "Windows", "Fonts");
  try {
    fs.mkdirSync(dir, { recursive: true });
    fs.copyFileSync(path.join(SRC, "assets", "fonts", file), path.join(dir, file));
    execSync(`reg add "HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts" /v "Noto Sans KR (TrueType)" ` +
             `/t REG_SZ /d "${path.join(dir, file)}" /f`, { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

function checks() {
  ensurePython();
  if (process.platform !== "win32") return step("PowerPoint", false, "렌더 확인은 Windows 필요");
  if (hasFont()) return step("글꼴", true, "Noto Sans KR");
  if (installFont()) step("글꼴", true, "Noto Sans KR 설치 (안 보이면 다시 로그인)");
  else step("글꼴", false, "Noto Sans KR 설치 실패");
}

const cmd = process.argv[2] || "install";
if (cmd === "--version" || cmd === "-v") {
  console.log(PKG.version);
} else if (cmd === "install" || cmd === "uninstall") {
  banner();
  cmd === "install" ? install() : uninstall();
} else {
  fail("사용: npx --allow-git=all github:muramasaDev/interx-deck-skill [install|uninstall|--version]");
}
