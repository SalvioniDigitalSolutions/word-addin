import { spawn } from "child_process";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const py = process.env.PYTHON || (process.platform === "win32" ? "python" : "python3");
const script = path.join(__dirname, "presidio_server.py");
const child = spawn(py, [script], {
  stdio: "inherit",
  env: { ...process.env },
});
child.on("exit", (code) => process.exit(code ?? 0));
