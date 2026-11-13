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

// 칼 좌표 (u 칼 방향, v 수직)의 색, 없으면 null
// - 칼끝 u=17, 손잡이 끝 u=-9 (날 약 3.4배)
function part(u, v, glint) {
  if (u >= -2.5 && u <= 17) {                   // 칼날: 휨, 칼끝 좁힘, 물결 무늬, 피
    const vb = v - 1.8 * ((u + 2.5) / 19.5) ** 2;
    const half = u > 14 ? 1.25 * (17 - u) / 3 : 1.25;
    if (Math.abs(vb) > half) return null;
    if (Math.abs(u - glint) < 1) return [255, 255, 255];
    if (u > 7 && vb < 0 && Math.sin(u * 2.7) + Math.cos(u * 1.3 + v) > 1.1) return u > 11 ? BLOOD : BLOOD_D;
    if (vb > 0.45) return SPINE;
    if (vb < -0.1 + 0.25 * Math.sin(u * 2.4)) return vb < -half + 0.3 ? HAMON : EDGE;
    return STEEL;
  }
  if (u >= -3.3 && u < -2.5) return Math.abs(v) <= 1.5 ? GOLD : null;   // 하바키
  if (u >= -4.3 && u < -3.3) return Math.abs(v) <= 3.2 ? TSUBA : null;  // 츠바
  if (u >= -9 && u < -4.3) {                    // 손잡이, 마름모 무늬
    if (Math.abs(v) > 1.4) return null;
    if (u < -8.3) return GOLD;                  // 카시라
    return Math.abs(((u * 1.3) % 2 + 2) % 2 - 1) > Math.abs(v) / 1.6 ? ITO : SAME;
  }
  return null;
}

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

// 각도 a 의 한 장면. drip 이면 칼날에서 피가 떨어짐
function frame(R, a, glint, drip) {
  const c = Math.cos(a), s = Math.sin(a), k = 17.5 / R, N = 2 * R + 1;
  const drops = drip ? [9.5, 12, 14.5].map((u, i) => {   // 날 끝 점의 화면 위치
    const v = -1.25 + 1.8 * ((u + 2.5) / 19.5) ** 2;
    return [Math.round((u * c - v * s) / k), Math.round((u * s + v * c) / k), 2 + i * 2];
  }) : [];
  return blocks(N, N, (x0, r) => {
    const x = x0 - R, y = R - r;
    const hit = part((x * c + y * s) * k, (-x * s + y * c) * k, glint);
    if (hit) return hit;
    for (const [dx, dy, len] of drops) if (x === dx && y < dy && y >= dy - len) return y === dy - len ? BLOOD : BLOOD_D;
    return null;
  });
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
  const R = Math.max(14, Math.min(30, (process.stdout.rows || 34) - 6));   // 화면 높이에 맞춤
  const end = Math.PI / 12;                      // 멈춘 자세: 칼끝 위로 15도
  const steps = COLOR ? 44 : 0;          // 색이 없으면 돌지 않음
  console.log();
  if (COLOR) process.stdout.write("\x1b[?25l"); // 커서 숨김
  for (let i = 1; i <= steps; i++) {
    const t = i / steps, ease = 1 - (1 - t) ** 3;
    const lines = frame(R, end - 4 * Math.PI * (1 - ease), -3 + 24 * t, i === steps);
    process.stdout.write(lines.join("\n") + `\n\x1b[${lines.length}A`);
    sleep(26);
  }
  // 멈춘 뒤 화면 폭까지 커지고 피가 떨어짐 (빈 줄은 뺌)
  const Rf = Math.max(R, Math.min(46, Math.floor(((process.stdout.columns || 100) - 6) / 2.3)));
  let up = 0;
  for (let j = COLOR ? 1 : 8; j <= 8; j++) {
    const lines = frame(Math.round(R + (Rf - R) * j / 8), end, 99, j === 8).filter((l) => l.trim());
    process.stdout.write((up ? `\x1b[${up}A` : "") + (COLOR ? "\x1b[J" : "") + lines.join("\n") + "\n");
    up = lines.length;
    sleep(22);
  }
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

const cmd = process.argv[2] || "install";
if (cmd === "--version" || cmd === "-v") {
  console.log(PKG.version);
} else if (cmd === "install" || cmd === "uninstall") {
  banner();
  cmd === "install" ? install() : uninstall();
} else {
  fail("사용: npx --allow-git=all muramasaDev/interx-deck-skill [install|uninstall|--version]");
}
