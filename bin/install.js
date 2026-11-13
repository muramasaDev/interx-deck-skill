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
const ITEMS = [".claude-plugin", "hooks", "skills", "scripts", "assets", "README.md", "CHANGELOG.md"];

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

// 村正: 세로로 서서 모빌처럼 도는 카타나 (반 칸 블록, 정사각 픽셀)
const W = 46;                  // 정보 카드 폭
const SPINE = [96, 104, 118];  // 칼등
const STEEL = [168, 178, 192];
const EDGE = [226, 232, 240];
const HAMON = [250, 252, 255]; // 물결 무늬
const BLOOD = [150, 0, 14];
const BLOOD_D = [92, 0, 8];
const GOLD = [176, 132, 40];
const TSUBA = [24, 22, 26];
const ITO = [120, 8, 20];      // 손잡이 끈 (붉은 실)
const SAME = [20, 18, 22];     // 손잡이 가죽 (검정)
const RED = [200, 30, 40];

// 픽셀 함수 px(x, y) 를 반 칸 블록 줄로 (y 는 위가 큼)
function blocks(w, h, px, indent = "  ") {
  const lines = [];
  for (let r = 0; r < h; r += 2) {
    let line = indent;
    for (let x = 0; x < w; x++) {
      const top = px(x, r), bot = r + 1 < h ? px(x, r + 1) : null;
      if (!COLOR) line += top && bot ? "█" : top ? "▀" : bot ? "▄" : " ";
      else if (top && bot) line += `\x1b[38;2;${top.join(";")};48;2;${bot.join(";")}m▀\x1b[0m`;
      else if (top) line += `\x1b[38;2;${top.join(";")}m▀\x1b[0m`;
      else if (bot) line += `\x1b[38;2;${bot.join(";")}m▄\x1b[0m`;
      else line += " ";
    }
    lines.push(line.replace(/\s+$/, ""));
  }
  return lines;
}

// 세로로 선 카타나를 세로축으로 돌린 한 장면 (L: 길이 픽셀, phi: 돌린 각)
// - 칼날은 판 (폭 3 픽셀, 두께 0.7), 손잡이 등은 원기둥, 츠바는 원판
function katana(L, phi, glint) {
  const pix = new Map();
  const put = (x, y, z, c) => {
    const k = `${Math.round(x)},${Math.round(y)}`, o = pix.get(k);
    if (!o || z > o.z) pix.set(k, { z, c });
  };
  const cs = Math.cos(phi), sn = Math.sin(phi), face = Math.abs(cs);
  const lit = (c, k) => c.map((v) => Math.min(255, Math.round(v * k)));
  const b0 = L * 0.27;                                   // 칼날 시작 (아래에서)
  for (let u = b0; u <= L; u += 0.25) {                  // 칼날: 휨, 칼끝, 물결 무늬, 피
    const s = (u - b0) / (L - b0);
    const half = s > 0.93 ? 1.5 * (1 - s) / 0.07 : 1.5 - 0.35 * s;
    const c = 2.4 * s * s;                               // 휨 (칼등 쪽으로)
    for (let v = -half; v <= half; v += 0.25) {
      for (const w of [-0.35, 0.35]) {
        const x = (c + v) * cs + w * sn, z = -(c + v) * sn + w * cs;
        let col = v < -half * 0.62 + 0.2 * Math.sin(u * 0.9) ? HAMON : v > half * 0.4 ? SPINE : STEEL;
        if (s > 0.5 && v < -half * 0.35 && Math.sin(u * 1.7) + Math.cos(u * 0.9) > 1.25) col = BLOOD;
        const k = face > 0.93 && Math.abs(u - glint) < 2.5 ? 1.25 : 0.42 + 0.62 * face;   // 정면일수록 밝음
        put(x, L - u, z, col === BLOOD ? col : lit(col, k));
      }
    }
  }
  const cyl = (u0, u1, r, color) => {                   // 원기둥: 가운데 밝게
    for (let u = u0; u < u1; u += 0.5) {
      for (let x = -r; x <= r; x += 0.25) put(x, L - u, 9, lit(color(u, Math.asin(x / r) + phi), 0.6 + 0.5 * Math.sqrt(1 - (x / r) ** 2)));
    }
  };
  cyl(b0 - 1, b0, 1.9, () => GOLD);                     // 하바키
  for (let x = -5; x <= 5; x += 0.25) {                 // 츠바 (원판, 옆에서 보면 띠)
    for (const du of [1, 1.5, 2]) put(x, L - (b0 - du), 9, Math.abs(x) > 4.4 ? GOLD : TSUBA);
  }
  cyl(1.2, b0 - 2, 1.5, (u, th) => {                    // 손잡이: 돌면 마름모 무늬가 흐름
    const a = (((u / 2.6 + th / Math.PI) % 1) + 1) % 1, b = (((u / 2.6 - th / Math.PI) % 1) + 1) % 1;
    return Math.min(Math.abs(a - 0.5), Math.abs(b - 0.5)) < 0.17 ? ITO : SAME;
  });
  cyl(0, 1.2, 1.7, () => GOLD);                         // 카시라
  const X = 8;
  return blocks(2 * X + 1, Math.ceil(L) + 1, (x0, y) => (pix.get(`${x0 - X},${y}`) || {}).c || null, " ".repeat(20));
}

// 7x5 픽셀 글꼴 (이름용)
const GLYPH = {
  C: [" ####", "#    ", "#    ", "#    ", "#    ", "#    ", " ####"],
  H: ["#   #", "#   #", "#   #", "#####", "#   #", "#   #", "#   #"],
  A: [" ### ", "#   #", "#   #", "#####", "#   #", "#   #", "#   #"],
  N: ["#   #", "##  #", "##  #", "# # #", "#  ##", "#  ##", "#   #"],
  G: [" ####", "#    ", "#    ", "# ###", "#   #", "#   #", " ### "],
  U: ["#   #", "#   #", "#   #", "#   #", "#   #", "#   #", " ### "],
  J: ["  ###", "    #", "    #", "    #", "    #", "#   #", " ### "],
  E: ["#####", "#    ", "#    ", "#### ", "#    ", "#    ", "#####"],
  O: [" ### ", "#   #", "#   #", "#   #", "#   #", "#   #", " ### "],
  " ": ["  ", "  ", "  ", "  ", "  ", "  ", "  "],
};

// 이름: 강철 글자 (글자 한 칸 = 한 줄)
// - cut 이면 베인 선 위쪽이 한 칸 밀리고 선에 피가 맺힘
function nameArt(text, cut) {
  const rows = [0, 1, 2, 3, 4, 5, 6].map((r) => [...text].map((ch) => GLYPH[ch][r]).join(" "));
  const w = rows[0].length + 5, h = 14;
  const slash = (x) => 10.5 - x * 7 / w;        // 베인 선 (왼쪽 아래에서 오른쪽 위)
  return blocks(w, h, (x, y) => {
    const gy = Math.floor(y / 2), above = cut && y < slash(x);
    const sx = x - 2 - (above ? 1 : 0);          // 베인 위쪽은 한 칸 밀림
    const on = rows[gy][sx] === "#";
    if (!on) return null;
    if (cut && Math.abs(y - slash(x)) < 0.7) return BLOOD;
    if (cut && y > slash(x) && y < slash(x) + 3 && (x * 5) % 7 === 0) return BLOOD_D;   // 흘러내림
    return mix(EDGE, SPINE, gy / 6);
  });
}

function banner() {
  const L = Math.max(30, Math.min(64, 2 * ((process.stdout.rows || 34) - 8)));   // 화면 높이에 맞춤
  const end = 0.35;                              // 멈춘 자세: 칼날 면이 거의 정면
  const steps = COLOR ? 72 : 0;                  // 색이 없으면 돌지 않음
  console.log();
  if (COLOR) process.stdout.write("\x1b[?25l"); // 커서 숨김
  for (let i = 1; i <= steps; i++) {             // 모빌: 천천히 멎으며 몇 번 되돌아 흔들림
    const t = i / steps;
    const phi = end + 5 * Math.PI * (1 - t) ** 3 + 0.9 * Math.sin(5 * Math.PI * t) * t * (1 - t) ** 1.5;
    const lines = katana(L, phi, L * (1.1 - t));
    process.stdout.write(lines.join("\n") + `\n\x1b[${lines.length}A`);
    sleep(34);
  }
  const fin = katana(L, end, -9).filter((l) => l.trim());   // 빈 줄은 뺌
  process.stdout.write((COLOR ? "\x1b[J" : "") + fin.join("\n") + "\n");
  console.log();
  console.log(`  ${grad("村正", RED, [120, 10, 20])}  ${bold("M U R A M A S A")}  ${dim("伊勢国 桑名")}`);
  console.log();
  console.log(`  ${dim("crafted by")}`);
  if ((process.stdout.columns || 100) >= 92) {    // 큰 이름: 먼저 온전히, 다음 베임
    const whole = nameArt("CHANGHUN JEONG", false);
    if (COLOR) {
      process.stdout.write(whole.join("\n") + `\n\x1b[${whole.length}A`);
      sleep(380);
    }
    process.stdout.write((COLOR ? "\x1b[J" : "") + nameArt("CHANGHUN JEONG", true).join("\n") + "\n");
  } else {
    console.log(`  ${grad("Changhun Jeong", EDGE, BLOOD)}`);
  }
  if (COLOR) process.stdout.write("\x1b[?25h");
  const head = `${dim("─ ")}${bold("interx-deck")} `;
  const ver = ` ${rgb(RED, "1;")("v" + PKG.version)}${dim(" ─")}`;
  console.log();
  console.log(`  ${dim("╭")}${head}${dim("─".repeat(Math.max(0, W - 2 - width(head) - width(ver))))}${ver}${dim("╮")}`);
  console.log(`  ${dim("│")} ${pad(`${dim("PPT engine for")} ${bold("Claude Code")}`, W - 4)} ${dim("│")}`);
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

// 설치본 점검: 엔진 자체 시험, 예제 덱 관문 (렌더와 PDF 포함)
function doctor() {
  const py = findPython();
  if (!py) fail("Python 3.11 이상이 없음: 먼저 설치 명령 실행");
  const root = fs.existsSync(path.join(DEST, "scripts", "gate.py")) ? DEST : SRC;
  const mark = path.join(root, MARK);
  step("설치본", true, `${fs.existsSync(mark) ? fs.readFileSync(mark, "utf8").trim() : PKG.version} (${root})`);
  const env = { ...process.env, PYTHONIOENCODING: "utf-8" };
  const exec = (args) => require("child_process").spawnSync(`${py} -P ${args}`, { shell: true, env, encoding: "utf8" });
  const tail = (r) => ((r.stdout || "") + (r.stderr || "")).trim().split("\n")
    .filter((l) => !/^주의/.test(l)).slice(-14).join("\n    ");
  const self = exec(`"${path.join(root, "scripts", "selfcheck.py")}" "${path.join(os.tmpdir(), "ixdoctor.pptx")}"`);
  if (self.status === 0) step("엔진 시험", true, "통과");
  else step("엔진 시험", false, "실패\n    " + tail(self));
  busy("예제 덱 관문 (렌더 포함, 2 ~ 4분)");
  const gate = exec(`"${path.join(root, "scripts", "gate.py")}" "${path.join(root, "scripts", "example_build.py")}"`);
  if (gate.status === 0) step("예제 덱 관문", true, "통과 (렌더, PDF 포함)");
  else step("예제 덱 관문", false, "실패\n    " + tail(gate));
  if (self.status || gate.status) process.exitCode = 1;
}

const cmd = process.argv[2] || "install";
if (cmd === "--version" || cmd === "-v") {
  console.log(PKG.version);
} else if (cmd === "install" || cmd === "uninstall") {
  banner();
  cmd === "install" ? install() : uninstall();
} else if (cmd === "doctor") {
  doctor();
} else {
  fail("사용: npx --allow-git=all muramasaDev/interx-deck-skill [install|uninstall|doctor|--version]");
}
