import { spawnSync } from "child_process";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, "..");
const py = process.env.PYTHON || (process.platform === "win32" ? "python" : "python3");
const reqFile = path.join(__dirname, "requirements-presidio.txt");

console.log(`Using Python: ${py} (set PYTHON=… to override)`);
console.log(
  "Tip: if pip fails on macOS, use Homebrew Python 3.11+ in a venv:\n" +
    "  python3.11 -m venv .venv && source .venv/bin/activate && npm run presidio:install\n",
);

function run(args, label) {
  console.log(`\n→ ${label}\n`);
  const r = spawnSync(py, args, { stdio: "inherit", cwd: root, env: process.env });
  if (r.status !== 0) {
    process.exit(r.status ?? 1);
  }
}

run(["-m", "pip", "install", "-r", reqFile], `pip install -r server/requirements-presidio.txt`);
run(["-m", "spacy", "download", "en_core_web_sm"], "python -m spacy download en_core_web_sm");
for (const model of ["de_core_news_sm", "fr_core_news_sm", "it_core_news_sm"]) {
  const r = spawnSync(py, ["-m", "spacy", "download", model], { stdio: "inherit", cwd: root, env: process.env });
  if (r.status !== 0) {
    console.warn(`\n(Optional) spacy download ${model} failed — Presidio stays English-only for that language.\n`);
  }
}
console.log("\nPresidio dependencies installed. Restart npm run server (or dev:all).\n");
