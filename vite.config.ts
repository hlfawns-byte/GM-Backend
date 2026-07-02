import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

type ManagedAccountRecord = {
  id: number;
  account: string;
  password?: string;
  displayName: string;
  role: string;
  games: string[];
  permissions: string[];
  status: "启用" | "停用" | "鍚敤" | "鍋滅敤";
};

type AdminCredentials = {
  account: string;
  password: string;
};

type GameConfigRecord = {
  id: number;
  name: string;
  serverName: string;
  serverUrl: string;
  environment: string;
  logo: string;
  iconUrl?: string;
  backgroundUrl?: string;
};

type ItemRecord = {
  id: number;
  name: string;
  icon?: string;
};

type UserLogRecord = {
  id: number;
  time: string;
  operator: string;
  role: string;
  action: string;
  target: string;
  game?: string;
  serverName?: string;
  detail?: string;
  result: "成功" | "失败";
};

const portal = process.env.GM_PORTAL === "prod" ? "prod" : "test";
const accountsFile = path.resolve(`data/${portal}/accounts.json`);
const adminCredentialFile = path.resolve(`data/${portal}/admin-credentials.json`);
const gamesFile = path.resolve(`data/${portal}/games.json`);
const itemsFile = path.resolve(`data/${portal}/items.json`);
const itemUploadFile = path.resolve(`data/${portal}/uploads/Item.xlsx`);
const mailTemplatesFile = path.resolve(`data/${portal}/mail-templates.json`);
const rewardTemplatesFile = path.resolve(`data/${portal}/reward-templates.json`);
const userLogsFile = path.resolve(`data/${portal}/user-logs.json`);
const noticesFile = path.resolve(`data/${portal}/notices.json`);
const gmServerTarget = "http://52.77.195.98:9089";
const bundledPython = "C:\\Users\\Touka\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe";
const defaultItemTable = "D:\\Project\\Bb_Main\\Config\\Excel\\Item.xlsx";
const defaultAccounts: ManagedAccountRecord[] = [];
const defaultGames: GameConfigRecord[] = [
  portal === "prod"
    ? { id: 1, name: "包包4", serverName: "正式服", serverUrl: "/gm-api", environment: "Prod", logo: "包" }
    : { id: 1, name: "包包4", serverName: "测试服", serverUrl: "/gm-api", environment: "Test", logo: "包" },
];

function readJsonBody(req: import("node:http").IncomingMessage): Promise<Record<string, unknown>> {
  return new Promise((resolve) => {
    let raw = "";
    req.on("data", (chunk) => {
      raw += chunk;
    });
    req.on("end", () => {
      try {
        resolve(raw ? JSON.parse(raw) : {});
      } catch {
        resolve({});
      }
    });
  });
}

function readBufferBody(req: import("node:http").IncomingMessage): Promise<Buffer> {
  return new Promise((resolve) => {
    const chunks: Buffer[] = [];
    req.on("data", (chunk) => chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)));
    req.on("end", () => resolve(Buffer.concat(chunks)));
  });
}

function readJsonFile<T>(file: string, fallback: T): T {
  if (!fs.existsSync(file)) return fallback;
  try {
    return JSON.parse(fs.readFileSync(file, "utf8")) as T;
  } catch {
    return fallback;
  }
}

function writeJsonFile(file: string, data: unknown) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(data, null, 2), "utf8");
}

function ensureAccountsFile() {
  if (!fs.existsSync(accountsFile)) {
    fs.mkdirSync(path.dirname(accountsFile), { recursive: true });
    fs.writeFileSync(accountsFile, JSON.stringify(defaultAccounts, null, 2), "utf8");
  }
}

function readAccounts(): ManagedAccountRecord[] {
  ensureAccountsFile();
  try {
    const parsed = JSON.parse(fs.readFileSync(accountsFile, "utf8"));
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeAccounts(accounts: ManagedAccountRecord[]) {
  ensureAccountsFile();
  fs.writeFileSync(accountsFile, JSON.stringify(accounts, null, 2), "utf8");
}

function ensureGamesFile() {
  if (!fs.existsSync(gamesFile)) {
    fs.mkdirSync(path.dirname(gamesFile), { recursive: true });
    fs.writeFileSync(gamesFile, JSON.stringify(defaultGames, null, 2), "utf8");
  }
}

function readGames(): GameConfigRecord[] {
  ensureGamesFile();
  try {
    const parsed = JSON.parse(fs.readFileSync(gamesFile, "utf8"));
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeGames(games: GameConfigRecord[]) {
  ensureGamesFile();
  fs.writeFileSync(gamesFile, JSON.stringify(games, null, 2), "utf8");
}

function extractMultipartFile(buffer: Buffer, contentType: string) {
  const boundary = contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/)?.[1] ?? contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/)?.[2];
  if (!boundary) return null;
  const raw = buffer.toString("binary");
  const marker = `--${boundary}`;
  const part = raw.split(marker).find((chunk) => chunk.includes("filename="));
  if (!part) return null;
  const headerEnd = part.indexOf("\r\n\r\n");
  if (headerEnd < 0) return null;
  const header = part.slice(0, headerEnd);
  const filename = header.match(/filename="([^"]+)"/)?.[1] ?? "Item.xlsx";
  let body = part.slice(headerEnd + 4);
  if (body.endsWith("\r\n")) body = body.slice(0, -2);
  return { filename, data: Buffer.from(body, "binary") };
}

function parseItemWorkbook(file: string): ItemRecord[] {
  const script = [
    "import json, sys, openpyxl",
    "path=sys.argv[1]",
    "wb=openpyxl.load_workbook(path, read_only=True, data_only=True)",
    "ws=wb['Item'] if 'Item' in wb.sheetnames else wb[wb.sheetnames[0]]",
    "rows=list(ws.iter_rows(values_only=True))",
    "headers=[str(v).strip() if v is not None else '' for v in rows[0]] if rows else []",
    "id_col=headers.index('Id') if 'Id' in headers else 1",
    "name_col=2",
    "icon_col=headers.index('Icon') if 'Icon' in headers else -1",
    "out=[]",
    "for row in rows[4:]:",
    "    if id_col >= len(row) or row[id_col] in (None, ''): continue",
    "    try: item_id=int(row[id_col])",
    "    except Exception: continue",
    "    name=str(row[name_col]).strip() if name_col < len(row) and row[name_col] not in (None, '') else ''",
    "    icon=str(row[icon_col]).strip() if icon_col >= 0 and icon_col < len(row) and row[icon_col] not in (None, '') else ''",
    "    out.append({'id': item_id, 'name': name, 'icon': icon})",
    "print(json.dumps(out, ensure_ascii=False))",
  ].join("\n");
  const python = fs.existsSync(bundledPython) ? bundledPython : "python";
  const result = spawnSync(python, ["-c", script, file], { encoding: "utf8", env: { ...process.env, PYTHONIOENCODING: "utf-8" }, maxBuffer: 10 * 1024 * 1024 });
  if (result.status !== 0) throw new Error(result.stderr || "解析 Item.xlsx 失败");
  const parsed = JSON.parse(result.stdout || "[]");
  return Array.isArray(parsed) ? parsed : [];
}

function readItems(): ItemRecord[] {
  const current = readJsonFile<ItemRecord[]>(itemsFile, []);
  if (current.length) return current;
  if (fs.existsSync(defaultItemTable)) {
    try {
      const parsed = parseItemWorkbook(defaultItemTable);
      writeJsonFile(itemsFile, parsed);
      return parsed;
    } catch {
      return [];
    }
  }
  return [];
}

function appendUserLog(log: Omit<UserLogRecord, "id" | "time">) {
  const logs = readJsonFile<UserLogRecord[]>(userLogsFile, []);
  const record: UserLogRecord = {
    id: Date.now(),
    time: new Date().toISOString().slice(0, 19).replace("T", " "),
    operator: String(log.operator || "unknown"),
    role: String(log.role || "用户"),
    action: String(log.action || "操作"),
    target: String(log.target || "-"),
    game: log.game,
    serverName: log.serverName,
    detail: log.detail,
    result: log.result,
  };
  writeJsonFile(userLogsFile, [record, ...logs].slice(0, 1000));
  return record;
}

function readAdminCredentials(): AdminCredentials | null {
  const envAccount = process.env.GM_ADMIN_ACC?.trim();
  const envPassword = process.env.GM_ADMIN_PWD?.trim();
  if (envAccount && envPassword) {
    return { account: envAccount, password: envPassword };
  }
  if (!fs.existsSync(adminCredentialFile)) {
    return null;
  }
  try {
    const parsed = JSON.parse(fs.readFileSync(adminCredentialFile, "utf8"));
    const account = String(parsed.account ?? "").trim();
    const password = String(parsed.password ?? "").trim();
    return account && password ? { account, password } : null;
  } catch {
    return null;
  }
}

function writeAdminCredentials(credentials: AdminCredentials) {
  fs.mkdirSync(path.dirname(adminCredentialFile), { recursive: true });
  fs.writeFileSync(adminCredentialFile, JSON.stringify(credentials, null, 2), "utf8");
}

function sendJson(res: import("node:http").ServerResponse, status: number, data: unknown) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(data));
}

function gmApiBase(serverUrl: unknown) {
  const raw = String(serverUrl ?? "").trim();
  if (!raw || raw.startsWith("/gm-api")) {
    return gmServerTarget;
  }
  return raw.replace(/\/$/, "");
}

async function loginToGameServer(serverUrl: unknown, credentials: AdminCredentials) {
  const apiBase = gmApiBase(serverUrl);
  const response = await fetch(`${apiBase}/login`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ Acc: credentials.account, Pwd: credentials.password }),
  });
  const token = response.headers.get("Token") ?? "";
  const text = await response.text();
  let payload: { data?: { Result?: number }; result?: string } = {};
  try {
    payload = text ? JSON.parse(text) : {};
  } catch {
    payload = {};
  }
  if (!response.ok) {
    throw new Error(`服务端返回 ${response.status}`);
  }
  if (payload.data?.Result !== 0) {
    throw new Error(payload.result || "账号或密码错误");
  }
  if (!token) {
    throw new Error("登录成功但响应 Header 中没有 Token");
  }

  const tokenCheck = await fetch(`${apiBase}/gmStateInfo`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Token: token },
    body: "{}",
  });
  if (!tokenCheck.ok) {
    throw new Error(`Token 验证失败：${tokenCheck.status}`);
  }

  return token;
}

function localAccountPlugin() {
  return {
    name: "local-account-api",
    configureServer(server: import("vite").ViteDevServer) {
      server.middlewares.use(async (req, res, next) => {
        const url = req.url ?? "";
        if (!url.startsWith("/local-api/")) {
          next();
          return;
        }

        if (url === "/local-api/accounts" && req.method === "GET") {
          sendJson(res, 200, { accounts: readAccounts().map(({ password, ...account }) => account) });
          return;
        }

        if (url === "/local-api/user-logs" && req.method === "GET") {
          sendJson(res, 200, { logs: readJsonFile<UserLogRecord[]>(userLogsFile, []) });
          return;
        }

        if (url === "/local-api/user-logs" && req.method === "POST") {
          const body = await readJsonBody(req);
          const record = appendUserLog({
            operator: String(body.operator ?? "unknown"),
            role: String(body.role ?? "用户"),
            action: String(body.action ?? "操作"),
            target: String(body.target ?? "-"),
            game: body.game ? String(body.game) : undefined,
            serverName: body.serverName ? String(body.serverName) : undefined,
            detail: body.detail ? String(body.detail) : undefined,
            result: body.result === "失败" ? "失败" : "成功",
          });
          sendJson(res, 200, { log: record });
          return;
        }

        if (url === "/local-api/games" && req.method === "GET") {
          sendJson(res, 200, { games: readGames() });
          return;
        }

        if (url === "/local-api/items" && req.method === "GET") {
          sendJson(res, 200, { items: readItems() });
          return;
        }

        if (url === "/local-api/notices" && req.method === "GET") {
          const defaults = [1, 2, 3].map((slot) => ({ slot, title: "", body: "", imagePath: "" }));
          const saved = readJsonFile<Array<Record<string, unknown>>>(noticesFile, []);
          sendJson(res, 200, { notices: defaults.map((item) => ({ ...item, ...(saved.find((notice) => Number(notice.slot) === item.slot) ?? {}) })) });
          return;
        }

        if (url === "/local-api/notices" && req.method === "POST") {
          const body = await readJsonBody(req);
          const notices = readJsonFile<Array<Record<string, unknown>>>(noticesFile, []);
          const slot = Math.min(3, Math.max(1, Number(body.slot) || 1));
          const next = {
            slot,
            title: String(body.title ?? ""),
            body: String(body.body ?? ""),
            imagePath: String(body.imagePath ?? ""),
            updatedAt: new Date().toISOString().slice(0, 19).replace("T", " "),
          };
          writeJsonFile(noticesFile, [next, ...notices.filter((notice) => Number(notice.slot) !== slot)]);
          sendJson(res, 200, { notice: next });
          return;
        }

        if (url === "/local-api/items/upload" && req.method === "POST") {
          try {
            const buffer = await readBufferBody(req);
            const uploaded = extractMultipartFile(buffer, String(req.headers["content-type"] ?? ""));
            if (!uploaded) {
              sendJson(res, 400, { error: "没有读取到上传文件" });
              return;
            }
            fs.mkdirSync(path.dirname(itemUploadFile), { recursive: true });
            fs.writeFileSync(itemUploadFile, uploaded.data);
            const items = parseItemWorkbook(itemUploadFile);
            writeJsonFile(itemsFile, items);
            sendJson(res, 200, { filename: uploaded.filename, items });
          } catch (error) {
            sendJson(res, 400, { error: error instanceof Error ? error.message : "解析道具表失败" });
          }
          return;
        }

        if (url === "/local-api/mail-templates" && req.method === "GET") {
          sendJson(res, 200, { templates: readJsonFile(mailTemplatesFile, []) });
          return;
        }

        if (url === "/local-api/mail-templates" && req.method === "POST") {
          const body = await readJsonBody(req);
          const templates = readJsonFile<Record<string, unknown>[]>(mailTemplatesFile, []);
          const now = new Date().toISOString().slice(0, 19).replace("T", " ");
          const id = String(body.id ?? `t-${Date.now()}`);
          const existing = templates.find((template) => template.id === id);
          const next = { id, name: String(body.name ?? "未命名模板"), title: String(body.title ?? ""), body: String(body.body ?? ""), contents: body.contents && typeof body.contents === "object" ? body.contents : undefined, createdAt: String(existing?.createdAt ?? now), updatedAt: now };
          writeJsonFile(mailTemplatesFile, [next, ...templates.filter((template) => template.id !== id)]);
          sendJson(res, 200, { template: next });
          return;
        }

        if (url === "/local-api/reward-templates" && req.method === "GET") {
          sendJson(res, 200, { templates: readJsonFile(rewardTemplatesFile, []) });
          return;
        }

        if (url === "/local-api/reward-templates" && req.method === "POST") {
          const body = await readJsonBody(req);
          const templates = readJsonFile<Record<string, unknown>[]>(rewardTemplatesFile, []);
          const now = new Date().toISOString().slice(0, 19).replace("T", " ");
          const id = String(body.id ?? `r-${Date.now()}`);
          const next = { id, title: String(body.title ?? "未命名奖励模板"), items: Array.isArray(body.items) ? body.items : [], createdAt: now, updatedAt: now };
          writeJsonFile(rewardTemplatesFile, [next, ...templates.filter((template) => template.id !== id)]);
          sendJson(res, 200, { template: next });
          return;
        }

        if (url === "/local-api/games" && req.method === "POST") {
          const body = await readJsonBody(req);
          const games = readGames();
          const name = String(body.name ?? "").trim();
          const serverName = String(body.serverName ?? "").trim();
          const serverUrl = String(body.serverUrl ?? "").trim();
          if (!name || !serverName || !serverUrl) {
            sendJson(res, 400, { error: "游戏名、区服和服务端地址不能为空" });
            return;
          }
          if (games.some((item) => item.name === name && item.serverName === serverName)) {
            sendJson(res, 409, { error: "该游戏区服已存在" });
            return;
          }
          const created: GameConfigRecord = {
            id: Date.now(),
            name,
            serverName,
            serverUrl,
            environment: String(body.environment ?? "Test").trim() || "Test",
            logo: String(body.logo ?? name.slice(0, 1)).trim() || "G",
            iconUrl: String(body.iconUrl ?? "").trim() || undefined,
            backgroundUrl: String(body.backgroundUrl ?? "").trim() || undefined,
          };
          writeGames([created, ...games]);
          sendJson(res, 200, { game: created });
          return;
        }

        const deleteGameMatch = url.match(/^\/local-api\/games\/(\d+)$/);
        const updateGameMatch = url.match(/^\/local-api\/games\/(\d+)$/);
        if (updateGameMatch && req.method === "PUT") {
          const id = Number(updateGameMatch[1]);
          const body = await readJsonBody(req);
          const games = readGames();
          const existing = games.find((game) => game.id === id);
          if (!existing) {
            sendJson(res, 404, { error: "游戏区服不存在" });
            return;
          }
          const name = String(body.name ?? existing.name).trim();
          const serverName = String(body.serverName ?? existing.serverName).trim();
          const serverUrl = String(body.serverUrl ?? existing.serverUrl).trim();
          if (!name || !serverName || !serverUrl) {
            sendJson(res, 400, { error: "游戏名、区服和服务端地址不能为空" });
            return;
          }
          if (games.some((item) => item.id !== id && item.name === name && item.serverName === serverName)) {
            sendJson(res, 409, { error: "该游戏区服已存在" });
            return;
          }
          const previousKey = `${existing.name}/${existing.serverName}`;
          const nextKey = `${name}/${serverName}`;
          const updated: GameConfigRecord = {
            ...existing,
            name,
            serverName,
            serverUrl,
            environment: String(body.environment ?? existing.environment).trim() || "Test",
            logo: String(body.logo ?? existing.logo ?? name.slice(0, 1)).trim() || "G",
            iconUrl: String(body.iconUrl ?? existing.iconUrl ?? "").trim() || undefined,
            backgroundUrl: String(body.backgroundUrl ?? existing.backgroundUrl ?? "").trim() || undefined,
          };
          writeGames(games.map((game) => (game.id === id ? updated : game)));
          if (previousKey !== nextKey) {
            writeAccounts(
              readAccounts().map((account) => ({
                ...account,
                games: account.games.map((game) => (game === previousKey ? nextKey : game)),
              })),
            );
          }
          sendJson(res, 200, { game: updated });
          return;
        }

        if (deleteGameMatch && req.method === "DELETE") {
          const id = Number(deleteGameMatch[1]);
          const deleted = readGames().find((game) => game.id === id);
          writeGames(readGames().filter((game) => game.id !== id));
          if (deleted) {
            const deletedKey = `${deleted.name}/${deleted.serverName}`;
            writeAccounts(readAccounts().map((account) => ({ ...account, games: account.games.filter((game) => game !== deletedKey) })));
          }
          sendJson(res, 200, { Result: 0, id });
          return;
        }

        if (url === "/local-api/accounts" && req.method === "POST") {
          const body = await readJsonBody(req);
          const accounts = readAccounts();
          const account = String(body.account ?? "").trim();
          const password = String(body.password ?? "").trim();
          if (!account || !password) {
            sendJson(res, 400, { error: "账号和密码不能为空" });
            return;
          }
          if (accounts.some((item) => item.account === account)) {
            sendJson(res, 409, { error: "账号已存在" });
            return;
          }
          const created: ManagedAccountRecord = {
            id: Date.now(),
            account,
            password,
            displayName: String(body.displayName ?? account),
            role: String(body.role ?? "运营"),
            games: Array.isArray(body.games) ? body.games.map(String) : ["包包4/测试服"],
            permissions: Array.isArray(body.permissions) ? body.permissions.map(String) : [],
            status: "启用",
          };
          writeAccounts([created, ...accounts]);
          const { password: _password, ...safeAccount } = created;
          sendJson(res, 200, { account: safeAccount });
          return;
        }

        const deleteMatch = url.match(/^\/local-api\/accounts\/(\d+)$/);
        if (deleteMatch && req.method === "DELETE") {
          const id = Number(deleteMatch[1]);
          writeAccounts(readAccounts().filter((account) => account.id !== id));
          sendJson(res, 200, { Result: 0, id });
          return;
        }

        if (url === "/local-api/operator-login" && req.method === "POST") {
          const body = await readJsonBody(req);
          const account = String(body.account ?? "").trim();
          const password = String(body.password ?? "").trim();
          const found = readAccounts().find((item) => item.account === account && item.password === password && (item.status === "启用" || item.status === "鍚敤"));
          if (!found) {
            sendJson(res, 401, { error: "账号或密码错误" });
            return;
          }
          const { password: _password, ...safeAccount } = found;
          sendJson(res, 200, { account: safeAccount });
          return;
        }

        if (url === "/local-api/gm-login" && req.method === "POST") {
          const body = await readJsonBody(req);
          const account = String(body.account ?? "").trim();
          const password = String(body.password ?? "").trim();
          const operator = body.operator as ManagedAccountRecord | undefined;

          try {
            const credentials = operator ? readAdminCredentials() : { account, password };
            if (!credentials) {
              sendJson(res, 428, { error: "请先使用游戏服务端管理员账号登录一次，再让子账号登录" });
              return;
            }
            const token = await loginToGameServer(body.serverUrl, credentials);
            if (!operator) {
              writeAdminCredentials({ account, password });
            }
            sendJson(res, 200, { token, adminAccount: credentials.account });
          } catch (error) {
            sendJson(res, 401, { error: error instanceof Error ? error.message : "登录失败" });
          }
          return;
        }

        if (url === "/local-api/gm-token" && req.method === "POST") {
          const body = await readJsonBody(req);
          const credentials = readAdminCredentials();
          if (!credentials) {
            sendJson(res, 428, { error: "请先使用游戏服务端管理员账号登录一次" });
            return;
          }
          try {
            const token = await loginToGameServer(body.serverUrl, credentials);
            sendJson(res, 200, { token, adminAccount: credentials.account });
          } catch (error) {
            sendJson(res, 401, { error: error instanceof Error ? error.message : "获取 Token 失败" });
          }
          return;
        }

        sendJson(res, 404, { error: "Not found" });
      });
    },
  };
}

export default defineConfig({
  plugins: [react(), localAccountPlugin()],
  server: {
    proxy: {
      "/gm-api": {
        target: gmServerTarget,
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/gm-api/, ""),
      },
    },
  },
});
