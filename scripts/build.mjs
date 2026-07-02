import { spawnSync } from "node:child_process";
import path from "node:path";

const portal = process.argv[2] === "prod" ? "prod" : "test";
const env = { ...process.env, GM_PORTAL: portal, VITE_GM_PORTAL: portal };
const tsc = path.resolve("node_modules", "typescript", "bin", "tsc");
const vite = path.resolve("node_modules", "vite", "bin", "vite.js");

for (const [script, args] of [[tsc, ["-b"]], [vite, ["build"]]]) {
  const result = spawnSync(process.execPath, [script, ...args], { stdio: "inherit", env, shell: false });
  if (result.error) {
    console.error(result.error.message);
    process.exit(1);
  }
  if (result.status !== 0) process.exit(result.status ?? 1);
}
