import { spawnSync } from "node:child_process";
import path from "node:path";

const portal = process.argv[2] === "prod" ? "prod" : "test";
const commandName = process.argv[3] || "dev";
const args = process.argv.slice(4);
const env = { ...process.env, GM_PORTAL: portal, VITE_GM_PORTAL: portal };
const vite = path.resolve("node_modules", "vite", "bin", "vite.js");
const result = spawnSync(process.execPath, [vite, commandName, ...args], { stdio: "inherit", env, shell: false });

if (result.error) {
  console.error(result.error.message);
  process.exit(1);
}

process.exit(result.status ?? 1);
