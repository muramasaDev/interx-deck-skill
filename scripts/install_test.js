// 배포 뒤 새 폴더 설치 시험: 빈 CLAUDE_CONFIG_DIR 에 GitHub 의 태그 판을 설치하고 doctor 로 점검
// 사용: node scripts/install_test.js [vX.Y.Z]   (없으면 plugin.json 판, 태그를 push 한 뒤)
// 격리는 CLAUDE_CONFIG_DIR 뿐: Python 패키지와 글꼴은 이 PC 에 이미 있는 것을 씀 (깨끗한 PC 시험 아님)
const { execSync } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");

const ver = process.argv[2] ||
  "v" + JSON.parse(fs.readFileSync(path.join(__dirname, "..", ".claude-plugin", "plugin.json"), "utf8")).version;
const cfg = fs.mkdtempSync(path.join(os.tmpdir(), "ixinstall_"));
const env = { ...process.env, CLAUDE_CONFIG_DIR: cfg };
const run = (args) => execSync(`npx --yes --allow-git=all muramasaDev/interx-deck-skill#${ver} ${args}`,
  { env, stdio: "inherit", cwd: cfg });
try {
  run("");
  run("doctor");
  const got = JSON.parse(fs.readFileSync(path.join(cfg, "skills", "interx-deck", ".claude-plugin", "plugin.json"), "utf8")).version;
  if ("v" + got !== ver) throw new Error(`설치된 판 ${got} 이 ${ver} 와 다름`);
  console.log(`설치 시험 통과: ${ver} (${cfg})`);
} catch (e) {
  console.error(`설치 시험 실패: ${e.message}`);
  process.exit(1);
}
