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
  avatarUrl?: string;
  role: string;
  games: string[];
  permissions: string[];
  isManager?: boolean;
  status: string;
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
  serverAccount?: string;
  serverPassword?: string;
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
  result: string;
};

type ScheduledMailRecord = {
  id: string;
  serverUrl: string;
  game?: string;
  serverName?: string;
  operator?: string;
  body: Record<string, unknown>;
  scheduleAt: number;
  createdAt: number;
  status: "pending" | "sending" | "sent" | "failed";
  error?: string;
  lastAttemptAt?: number;
  retryCount?: number;
  sentAt?: number;
  response?: unknown;
};

const portal = process.env.GM_PORTAL === "prod" ? "prod" : "test";
const accountsFile = path.resolve(`data/${portal}/accounts.json`);
const adminCredentialFile = path.resolve(`data/${portal}/admin-credentials.json`);
const adminProfileFile = path.resolve(`data/${portal}/admin-profile.json`);
const gamesFile = path.resolve(`data/${portal}/games.json`);
const itemsFile = path.resolve(`data/${portal}/items.json`);
const itemUploadFile = path.resolve(`data/${portal}/uploads/Item.xlsx`);
const mailTemplatesFile = path.resolve(`data/${portal}/mail-templates.json`);
const rewardTemplatesFile = path.resolve(`data/${portal}/reward-templates.json`);
const scheduledMailsFile = path.resolve(`data/${portal}/scheduled-mails.json`);
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

function readScheduledMails(): ScheduledMailRecord[] {
  return readJsonFile<ScheduledMailRecord[]>(scheduledMailsFile, []);
}

function writeScheduledMails(mails: ScheduledMailRecord[]) {
  writeJsonFile(scheduledMailsFile, mails);
}

function credentialsForServer(serverUrl: unknown): AdminCredentials | null {
  const normalized = String(serverUrl ?? "").trim().replace(/\/$/, "");
  const game = readGames().find((item) => String(item.serverUrl ?? "").trim().replace(/\/$/, "") === normalized);
  const serverAccount = String(game?.serverAccount ?? "").trim();
  const serverPassword = String(game?.serverPassword ?? "").trim();
  if (serverAccount && serverPassword) return { account: serverAccount, password: serverPassword };
  return readAdminCredentials();
}

async function postToGameServer(serverUrl: unknown, endpoint: string, body: unknown) {
  const credentials = credentialsForServer(serverUrl);
  if (!credentials) throw new Error("缺少游戏服务端管理员账号密码");
  const token = await loginToGameServer(serverUrl, credentials);
  const upstream = await fetch(`${gmApiBase(serverUrl)}${endpoint}`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Token: token },
    body: JSON.stringify(body ?? {}),
  });
  const text = await upstream.text();
  let payload: unknown = {};
  try {
    payload = text ? JSON.parse(text) : {};
  } catch {
    payload = { raw: text };
  }
  return { ok: upstream.ok, status: upstream.status, payload };
}

function publicScheduledMail(task: ScheduledMailRecord) {
  return {
    ...task.body,
    Id: task.id,
    St: task.scheduleAt,
    CreateTime: task.createdAt,
    __scheduled: true,
    __scheduledStatus: task.status,
    __scheduledError: task.error,
    __serverName: task.serverName,
  };
}

function readAdminProfile(account: string) {
  const profile = readJsonFile<Record<string, unknown>>(adminProfileFile, {});
  return {
    account,
    displayName: String(profile.displayName ?? account ?? "admin"),
    avatarUrl: String(profile.avatarUrl ?? ""),
  };
}

function writeAdminProfile(profile: Record<string, unknown>) {
  writeJsonFile(adminProfileFile, profile);
}

function sendJson(res: import("node:http").ServerResponse, status: number, data: unknown) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  res.end(JSON.stringify(data));
}

function appVersionPayload() {
  const packageJson = readJsonFile<{ version?: string }>(path.resolve("package.json"), {});
  const version = String(packageJson.version ?? "0.0.0");
  const srcFile = path.resolve("src/main.tsx");
  const configFile = path.resolve("vite.config.ts");
  const srcMtime = fs.existsSync(srcFile) ? fs.statSync(srcFile).mtimeMs : 0;
  const configMtime = fs.existsSync(configFile) ? fs.statSync(configFile).mtimeMs : 0;
  return {
    version,
    build: `${version}-${Math.max(srcMtime, configMtime)}`,
    portal,
  };
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

let scheduledMailRunning = false;
async function processScheduledMails() {
  if (scheduledMailRunning) return;
  scheduledMailRunning = true;
  try {
    const now = Math.floor(Date.now() / 1000);
    const mails = readScheduledMails();
    let changed = false;
    for (const task of mails) {
      const canRecoverSending = task.status === "sending" && now - Number(task.lastAttemptAt ?? 0) > 60;
      const canRetryFailed = task.status === "failed" && Number(task.retryCount ?? 0) < 3 && now - Number(task.lastAttemptAt ?? 0) > 60;
      if (!["pending", "sending", "failed"].includes(task.status) || Number(task.scheduleAt) > now) continue;
      if (task.status === "sending" && !canRecoverSending) continue;
      if (task.status === "failed" && !canRetryFailed) continue;
      task.status = "sending";
      task.lastAttemptAt = now;
      task.retryCount = Number(task.retryCount ?? 0) + 1;
      task.error = undefined;
      changed = true;
      writeScheduledMails(mails);
      try {
        const result = await postToGameServer(task.serverUrl, "/gmMailAdd", task.body);
        const payload = result.payload && typeof result.payload === "object" ? result.payload as Record<string, unknown> : {};
        const data = payload.data && typeof payload.data === "object" ? payload.data as Record<string, unknown> : {};
        const businessResult = data.Result ?? payload.Result;
        if (!result.ok || (businessResult !== undefined && Number(businessResult) !== 0)) {
          throw new Error(String(payload.result ?? payload.error ?? `HTTP ${result.status}`));
        }
        task.status = "sent";
        task.sentAt = Math.floor(Date.now() / 1000);
        task.response = result.payload;
        appendUserLog({ operator: task.operator || "unknown", role: "管理员", action: "发送定时邮件", target: task.id, game: task.game, serverName: task.serverName, result: "成功" });
      } catch (error) {
        task.status = "failed";
        task.error = error instanceof Error ? error.message : "发送失败";
        appendUserLog({ operator: task.operator || "unknown", role: "管理员", action: "发送定时邮件", target: task.id, game: task.game, serverName: task.serverName, detail: task.error, result: "失败" });
      }
      changed = true;
      writeScheduledMails(mails);
    }
    if (changed) {
      writeScheduledMails(mails.filter((task) => task.status !== "sent").slice(0, 500));
    }
  } finally {
    scheduledMailRunning = false;
  }
}

function localAccountPlugin() {
  let timer: NodeJS.Timeout | undefined;
  return {
    name: "local-account-api",
    configureServer(server: import("vite").ViteDevServer) {
      timer = setInterval(() => void processScheduledMails(), 10_000);
      void processScheduledMails();
      server.httpServer?.once("close", () => {
        if (timer) clearInterval(timer);
      });
      server.middlewares.use(async (req, res, next) => {
        const url = req.url ?? "";
        const pathname = url.split("?")[0];
        if (!pathname.startsWith("/local-api/")) {
          next();
          return;
        }

        if (pathname === "/local-api/app-version" && req.method === "GET") {
          sendJson(res, 200, appVersionPayload());
          return;
        }

        if (pathname === "/local-api/accounts" && req.method === "GET") {
          sendJson(res, 200, { accounts: readAccounts().map(({ password, ...account }) => account) });
          return;
        }

        if (pathname === "/local-api/profile" && req.method === "GET") {
          const requestUrl = new URL(req.url ?? "/", "http://localhost");
          const account = String(requestUrl.searchParams.get("account") ?? "").trim();
          const found = readAccounts().find((item) => item.account === account);
          if (found) {
            const { password: _password, ...safeAccount } = found;
            sendJson(res, 200, { profile: safeAccount });
            return;
          }
          sendJson(res, 200, { profile: readAdminProfile(account || readAdminCredentials()?.account || "admin") });
          return;
        }

        if (pathname === "/local-api/profile" && req.method === "POST") {
          const body = await readJsonBody(req);
          const account = String(body.account ?? "").trim();
          const displayName = String(body.displayName ?? account).trim();
          const avatarUrl = String(body.avatarUrl ?? "");
          const oldPassword = String(body.oldPassword ?? "");
          const newPassword = String(body.newPassword ?? "");
          if (!account || !displayName) {
            sendJson(res, 400, { error: "账号和显示名称不能为空" });
            return;
          }
          const accounts = readAccounts();
          const index = accounts.findIndex((item) => item.account === account);
          if (index >= 0) {
            if (newPassword) {
              if (accounts[index].password !== oldPassword) {
                sendJson(res, 403, { error: "原密码不正确" });
                return;
              }
              accounts[index].password = newPassword;
            }
            accounts[index] = { ...accounts[index], displayName, avatarUrl };
            writeAccounts(accounts);
            const { password: _password, ...safeAccount } = accounts[index];
            sendJson(res, 200, { profile: safeAccount });
            return;
          }
          if (newPassword) {
            sendJson(res, 400, { error: "admin 游戏服务端密码不能在后台修改" });
            return;
          }
          const profile = { account, displayName, avatarUrl };
          writeAdminProfile(profile);
          sendJson(res, 200, { profile });
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

        if (url === "/local-api/game-servers" && req.method === "GET") {
          const servers = Array.from({ length: 200 }, (_, index) => ({ id: index + 1, name: `游戏内区服 ${index + 1}` }));
          sendJson(res, 200, { servers });
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
          const now = Math.floor(Date.now() / 1000);
          const id = String(body.id ?? `t-${Date.now()}`);
          const existing = templates.find((template) => template.id === id);
          const next = { id, name: String(body.name ?? "未命名模板").slice(0, 30), title: String(body.title ?? ""), body: String(body.body ?? ""), contents: body.contents && typeof body.contents === "object" ? body.contents : undefined, createdAt: existing?.createdAt ?? now, updatedAt: now };
          writeJsonFile(mailTemplatesFile, [next, ...templates.filter((template) => template.id !== id)]);
          sendJson(res, 200, { template: next });
          return;
        }

        const mailTemplateMatch = url.match(/^\/local-api\/mail-templates\/([^/]+)$/);
        if (mailTemplateMatch && req.method === "DELETE") {
          const id = decodeURIComponent(mailTemplateMatch[1]);
          const templates = readJsonFile<Record<string, unknown>[]>(mailTemplatesFile, []);
          writeJsonFile(mailTemplatesFile, templates.filter((template) => String(template.id) !== id));
          sendJson(res, 200, { Result: 0, id });
          return;
        }

        if (url === "/local-api/reward-templates" && req.method === "GET") {
          sendJson(res, 200, { templates: readJsonFile(rewardTemplatesFile, []) });
          return;
        }

        if (pathname === "/local-api/scheduled-mails" && req.method === "GET") {
          const serverUrl = String(new URL(req.url ?? "", "http://localhost").searchParams.get("serverUrl") ?? "").trim();
          const mails = readScheduledMails()
            .filter((task) => !serverUrl || String(task.serverUrl ?? "") === serverUrl)
            .filter((task) => task.status !== "sent")
            .map(publicScheduledMail);
          sendJson(res, 200, { mails });
          return;
        }

        if (pathname === "/local-api/scheduled-mails" && req.method === "POST") {
          const body = await readJsonBody(req);
          const serverUrl = String(body.serverUrl ?? "").trim();
          const mailBody = body.body && typeof body.body === "object" ? body.body as Record<string, unknown> : null;
          const scheduleAt = Number(body.scheduleAt);
          if (!serverUrl || !mailBody || !Number.isFinite(scheduleAt)) {
            sendJson(res, 400, { error: "定时邮件参数不完整" });
            return;
          }
          if (scheduleAt <= Math.floor(Date.now() / 1000)) {
            sendJson(res, 400, { error: "定时发送时间必须晚于当前时间" });
            return;
          }
          const task: ScheduledMailRecord = {
            id: `scheduled-${Date.now()}`,
            serverUrl,
            game: String(body.game ?? ""),
            serverName: String(body.serverName ?? ""),
            operator: String(body.operator ?? ""),
            body: { ...mailBody, St: 0 },
            scheduleAt,
            createdAt: Math.floor(Date.now() / 1000),
            status: "pending",
          };
          writeScheduledMails([task, ...readScheduledMails()]);
          appendUserLog({ operator: task.operator || "unknown", role: "管理员", action: "创建定时邮件", target: task.id, game: task.game, serverName: task.serverName, detail: `计划发送时间 ${scheduleAt}`, result: "成功" });
          sendJson(res, 200, { mail: publicScheduledMail(task) });
          return;
        }

        const scheduledMailMatch = pathname.match(/^\/local-api\/scheduled-mails\/([^/]+)$/);
        if (scheduledMailMatch && req.method === "DELETE") {
          const id = decodeURIComponent(scheduledMailMatch[1]);
          const mails = readScheduledMails();
          const task = mails.find((item) => String(item.id) === id);
          writeScheduledMails(mails.filter((item) => String(item.id) !== id));
          if (task) appendUserLog({ operator: task.operator || "unknown", role: "管理员", action: "撤回定时邮件", target: id, game: task.game, serverName: task.serverName, result: "成功" });
          sendJson(res, 200, { Result: 0, id });
          return;
        }

        if (url === "/local-api/reward-templates" && req.method === "POST") {
          const body = await readJsonBody(req);
          const templates = readJsonFile<Record<string, unknown>[]>(rewardTemplatesFile, []);
          const now = Math.floor(Date.now() / 1000);
          const id = String(body.id ?? `r-${Date.now()}`);
          const existing = templates.find((template) => template.id === id);
          const next = { id, title: String(body.title ?? "未命名奖励模板"), items: Array.isArray(body.items) ? body.items : [], createdAt: existing?.createdAt ?? now, updatedAt: now };
          writeJsonFile(rewardTemplatesFile, [next, ...templates.filter((template) => template.id !== id)]);
          sendJson(res, 200, { template: next });
          return;
        }

        const rewardTemplateMatch = url.match(/^\/local-api\/reward-templates\/([^/]+)$/);
        if (rewardTemplateMatch && req.method === "DELETE") {
          const id = decodeURIComponent(rewardTemplateMatch[1]);
          const templates = readJsonFile<Record<string, unknown>[]>(rewardTemplatesFile, []);
          writeJsonFile(rewardTemplatesFile, templates.filter((template) => String(template.id) !== id));
          sendJson(res, 200, { Result: 0, id });
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
            serverAccount: String(body.serverAccount ?? "").trim() || undefined,
            serverPassword: String(body.serverPassword ?? "").trim() || undefined,
            iconUrl: String(body.iconUrl ?? "").trim() || undefined,
            backgroundUrl: String(body.backgroundUrl ?? "").trim() || undefined,
          };
          writeGames([created, ...games]);
          sendJson(res, 200, { game: created });
          return;
        }

        if (url === "/local-api/games/reorder" && req.method === "POST") {
          const body = await readJsonBody(req);
          const ids = Array.isArray(body.ids) ? body.ids.map((id) => Number(id)).filter(Number.isFinite) : [];
          const uniqueIds = Array.from(new Set(ids));
          if (uniqueIds.length === 0) {
            sendJson(res, 400, { error: "游戏排序不能为空" });
            return;
          }
          const games = readGames();
          const orderedMap = new Map(games.map((game) => [Number(game.id), game]));
          const orderedGames = uniqueIds.map((id) => orderedMap.get(id)).filter((game): game is GameConfigRecord => Boolean(game));
          if (orderedGames.length !== uniqueIds.length) {
            sendJson(res, 400, { error: "游戏排序包含不存在的区服" });
            return;
          }
          const movedIds = new Set(uniqueIds);
          const orderedQueue = [...orderedGames];
          const nextGames = games.map((game) => (movedIds.has(Number(game.id)) ? orderedQueue.shift() ?? game : game));
          writeGames(nextGames);
          sendJson(res, 200, { games: nextGames });
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
            serverAccount: String(body.serverAccount ?? existing.serverAccount ?? "").trim() || undefined,
            serverPassword: String(body.serverPassword ?? existing.serverPassword ?? "").trim() || undefined,
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
            isManager: Boolean(body.isManager),
            status: "启用",
          };
          writeAccounts([created, ...accounts]);
          const { password: _password, ...safeAccount } = created;
          sendJson(res, 200, { account: safeAccount });
          return;
        }

        const deleteMatch = url.match(/^\/local-api\/accounts\/(\d+)$/);
        if (deleteMatch && req.method === "PUT") {
          const id = Number(deleteMatch[1]);
          const body = await readJsonBody(req);
          const accounts = readAccounts();
          const index = accounts.findIndex((account) => account.id === id);
          if (index < 0) {
            sendJson(res, 404, { error: "账号不存在" });
            return;
          }
          const next = {
            ...accounts[index],
            displayName: String(body.displayName ?? accounts[index].displayName ?? accounts[index].account),
            role: String(body.role ?? accounts[index].role ?? "运营"),
            games: Array.isArray(body.games) ? body.games.map(String) : accounts[index].games,
            permissions: Array.isArray(body.permissions) ? body.permissions.map(String) : accounts[index].permissions,
            isManager: Boolean(body.isManager),
            status: String(body.status ?? accounts[index].status ?? "启用") as ManagedAccountRecord["status"],
          };
          const password = String(body.password ?? "").trim();
          if (password) next.password = password;
          accounts[index] = next;
          writeAccounts(accounts);
          const { password: _password, ...safeAccount } = next;
          sendJson(res, 200, { account: safeAccount });
          return;
        }

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
          const found = readAccounts().find((item) => item.account === account && item.password === password && (item.status === "启用" || item.status === "鍚敤" || item.status === "閸氼垳鏁?"));
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
          const serverAccount = String(body.serverAccount ?? "").trim();
          const serverPassword = String(body.serverPassword ?? "").trim();
          const credentials = serverAccount && serverPassword ? { account: serverAccount, password: serverPassword } : readAdminCredentials();
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

        if (url === "/local-api/gm-post" && req.method === "POST") {
          const body = await readJsonBody(req);
          const endpoint = String(body.endpoint ?? "");
          const token = String(body.token ?? "");
          if (!endpoint.startsWith("/")) {
            sendJson(res, 400, { error: "接口地址不正确" });
            return;
          }
          if (!token) {
            sendJson(res, 401, { error: "Token 为空，请重新进入区服" });
            return;
          }
          try {
            const apiBase = gmApiBase(String(body.serverUrl ?? ""));
            const upstream = await fetch(`${apiBase}${endpoint}`, {
              method: "POST",
              headers: { "Content-Type": "application/json", Token: token },
              body: JSON.stringify(body.body ?? {}),
            });
            const text = await upstream.text();
            res.statusCode = upstream.status;
            res.setHeader("Content-Type", upstream.headers.get("content-type") || "application/json; charset=utf-8");
            res.end(text);
          } catch (error) {
            sendJson(res, 502, { error: error instanceof Error ? error.message : "游戏服务器连接失败" });
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

