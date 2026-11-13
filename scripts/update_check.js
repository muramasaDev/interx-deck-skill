// 새 판 알림 (SessionStart 훅), 하루 한 번, 4초 제한
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..");
const CACHE = path.join(process.env.CLAUDE_CONFIG_DIR || path.join(os.homedir(), ".claude"), "interx-deck", "update.json");
const REPO = "https://github.com/muramasaDev/interx-deck-skill";
const num = (v) => v.replace(/^v/, "").split(".").map(Number);
const save = (c) => { fs.mkdirSync(path.dirname(CACHE), { recursive: true }); fs.writeFileSync(CACHE, JSON.stringify(c)); };
const newer = (a, b) => { const x = num(a), y = num(b); for (let i = 0; i < 3; i++) if (x[i] !== y[i]) return x[i] > y[i]; return false; };

try {
  const mine = JSON.parse(fs.readFileSync(path.join(ROOT, ".claude-plugin", "plugin.json"), "utf8")).version;
  let c = {};
  try { c = JSON.parse(fs.readFileSync(CACHE, "utf8")); } catch { /* 처음 */ }
  if (!c.checked || Date.now() - c.checked > 864e5) {
    save({ ...c, checked: Date.now() });   // 실패해도 하루 한 번만 시도
    const out = execSync(`git ls-remote --tags --refs ${REPO}`, {
      timeout: 4000, stdio: ["ignore", "pipe", "ignore"],
      env: { ...process.env, GIT_TERMINAL_PROMPT: "0", GCM_INTERACTIVE: "never", GIT_ASKPASS: "", SSH_ASKPASS: "" },
    }).toString();
    const tags = [...out.matchAll(/refs\/tags\/(v\d+\.\d+\.\d+)$/gm)].map((m) => m[1]);
    c = { checked: Date.now(), latest: tags.reduce((a, b) => (newer(b, a) ? b : a), "v0.0.0") };
    save(c);
  }
  if (c.latest && newer(c.latest, mine)) {
    process.stdout.write(JSON.stringify({
      systemMessage: `interx-deck 새 판 ${c.latest} (지금 ${mine}): npx --allow-git=all muramasaDev/interx-deck-skill`,
    }));
  }
} catch { /* 네트워크, 권한 문제는 알리지 않음 */ }
