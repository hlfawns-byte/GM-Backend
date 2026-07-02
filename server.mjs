import { createServer } from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const portalArg = process.argv.find((arg) => arg.startsWith("--portal="))?.split("=")[1];
const portal = (portalArg || process.env.GM_PORTAL) === "prod" ? "prod" : "test";
const host = process.env.HOST || "0.0.0.0";
const port = Number(process.env.PORT || (portal === "prod" ? 4174 : 4173));
const gmServerTarget = (process.env.GM_SERVER_TARGET || "http://52.77.195.98:9089").replace(/\/$/, "");
const distDir = path.resolve(__dirname, "dist");
const dataDir = path.resolve(__dirname, "data", portal);

const files = {
  accounts: path.join(dataDir, "accounts.json"),
  adminCredentials: path.join(dataDir, "admin-credentials.json"),
  games: path.join(dataDir, "games.json"),
  items: path.join(dataDir, "items.json"),
  itemUpload: path.join(dataDir, "uploads", "Item.xlsx"),
  mailTemplates: path.join(dataDir, "mail-templates.json"),
  rewardTemplates: path.join(dataDir, "reward-templates.json"),
  userLogs: path.join(dataDir, "user-logs.json"),
  notices: path.join(dataDir, "notices.json"),
};

const defaultAccounts = [];
const defaultGames = [
  portal === "prod"
    ? { id: 1, name: "包包4", serverName: "正式服", serverUrl: "/gm-api", environment: "Prod", logo: "包" }
    : { id: 1, name: "包包4", serverName: "测试服", serverUrl: "/gm-api", environment: "Test", logo: "包" },
];

function ensureJson(file, fallback) {
  if (!fs.existsSync(file)) writeJson(file, fallback);
}

function readJson(file, fallback) {
  if (!fs.existsSync(file)) return fallback;
  try {
    return JSON.parse(fs.readFileSync(file, "utf8"));
  } catch {
    return fallback;
  }
}

function writeJson(file, data) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(data, null, 2), "utf8");
}

function readAccounts() {
  ensureJson(files.accounts, defaultAccounts);
  const accounts = readJson(files.accounts, []);
  return Array.isArray(accounts) ? accounts : [];
}

function writeAccounts(accounts) {
  writeJson(files.accounts, accounts);
}

function readGames() {
  ensureJson(files.games, defaultGames);
  const games = readJson(files.games, []);
  return Array.isArray(games) ? games : [];
}

function writeGames(games) {
  writeJson(files.games, games);
}

function sendJson(res, status, data) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(data));
}

function readBody(req) {
  return new Promise((resolve) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)));
    req.on("end", () => resolve(Buffer.concat(chunks)));
  });
}

async function readJsonBody(req) {
  const buffer = await readBody(req);
  try {
    return buffer.length ? JSON.parse(buffer.toString("utf8")) : {};
  } catch {
    return {};
  }
}

function appendUserLog(log) {
  const logs = readJson(files.userLogs, []);
  const record = {
    id: Date.now(),
    time: new Date().toISOString().slice(0, 19).replace("T", " "),
    operator: String(log.operator || "unknown"),
    role: String(log.role || "用户"),
    action: String(log.action || "操作"),
    target: String(log.target || "-"),
    game: log.game,
    serverName: log.serverName,
    detail: log.detail,
    result: log.result === "失败" ? "失败" : "成功",
  };
  writeJson(files.userLogs, [record, ...(Array.isArray(logs) ? logs : [])].slice(0, 1000));
  return record;
}

function readAdminCredentials() {
  const envAccount = process.env.GM_ADMIN_ACC?.trim();
  const envPassword = process.env.GM_ADMIN_PWD?.trim();
  if (envAccount && envPassword) return { account: envAccount, password: envPassword };
  const parsed = readJson(files.adminCredentials, null);
  const account = String(parsed?.account ?? "").trim();
  const password = String(parsed?.password ?? "").trim();
  return account && password ? { account, password } : null;
}

function writeAdminCredentials(credentials) {
  writeJson(files.adminCredentials, credentials);
}

function gmApiBase(serverUrl) {
  const raw = String(serverUrl ?? "").trim();
  if (!raw || raw.startsWith("/gm-api")) return gmServerTarget;
  return raw.replace(/\/$/, "");
}

async function loginToGameServer(serverUrl, credentials) {
  const apiBase = gmApiBase(serverUrl);
  const response = await fetch(`${apiBase}/login`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ Acc: credentials.account, Pwd: credentials.password }),
  });
  const token = response.headers.get("Token") ?? "";
  const text = await response.text();
  let payload = {};
  try {
    payload = text ? JSON.parse(text) : {};
  } catch {
    payload = {};
  }
  if (!response.ok) throw new Error(`服务端返回 ${response.status}`);
  if (payload?.data?.Result !== 0) throw new Error(payload?.result || "账号或密码错误");
  if (!token) throw new Error("登录成功但响应 Header 中没有 Token");

  const tokenCheck = await fetch(`${apiBase}/gmStateInfo`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Token: token },
    body: "{}",
  });
  if (!tokenCheck.ok) throw new Error(`Token 验证失败：${tokenCheck.status}`);
  return token;
}

function extractMultipartFile(buffer, contentType) {
  const boundary = contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/)?.[1] ?? contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/)?.[2];
  if (!boundary) return null;
  const raw = buffer.toString("binary");
  const part = raw.split(`--${boundary}`).find((chunk) => chunk.includes("filename="));
  if (!part) return null;
  const headerEnd = part.indexOf("\r\n\r\n");
  if (headerEnd < 0) return null;
  const header = part.slice(0, headerEnd);
  const filename = header.match(/filename="([^"]+)"/)?.[1] ?? "Item.xlsx";
  let body = part.slice(headerEnd + 4);
  if (body.endsWith("\r\n")) body = body.slice(0, -2);
  return { filename, data: Buffer.from(body, "binary") };
}

function parseItemWorkbook(file) {
  const workbook = XLSX.readFile(file, { cellDates: false });
  const sheetName = workbook.SheetNames.includes("Item") ? "Item" : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error("Item.xlsx 中没有可读取的工作表");
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!Array.isArray(rows) || rows.length === 0) return [];

  const headers = rows[0].map((value) => String(value ?? "").trim());
  const findColumn = (names, fallback) => {
    const index = headers.findIndex((header) => names.some((name) => header.toLowerCase() === name.toLowerCase()));
    return index >= 0 ? index : fallback;
  };
  const idCol = findColumn(["Id", "ID", "ItemId", "ItemID", "道具ID"], 1);
  const nameCol = findColumn(["Name", "ItemName", "名称", "道具名"], 2);
  const iconCol = findColumn(["Icon", "图标"], -1);

  return rows.slice(1).flatMap((row) => {
    const itemId = Number(row[idCol]);
    if (!Number.isFinite(itemId) || itemId <= 0) return [];
    return [{
      id: itemId,
      name: String(row[nameCol] ?? "").trim(),
      icon: iconCol >= 0 ? String(row[iconCol] ?? "").trim() : "",
    }];
  });
}

async function handleLocalApi(req, res, pathname) {
  if (pathname === "/local-api/accounts" && req.method === "GET") {
    sendJson(res, 200, { accounts: readAccounts().map(({ password, ...account }) => account) });
    return true;
  }

  if (pathname === "/local-api/user-logs" && req.method === "GET") {
    sendJson(res, 200, { logs: readJson(files.userLogs, []) });
    return true;
  }

  if (pathname === "/local-api/user-logs" && req.method === "POST") {
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
    return true;
  }

  if (pathname === "/local-api/games" && req.method === "GET") {
    sendJson(res, 200, { games: readGames() });
    return true;
  }

  if (pathname === "/local-api/items" && req.method === "GET") {
    sendJson(res, 200, { items: readJson(files.items, []) });
    return true;
  }

  if (pathname === "/local-api/items/upload" && req.method === "POST") {
    try {
      const uploaded = extractMultipartFile(await readBody(req), String(req.headers["content-type"] ?? ""));
      if (!uploaded) {
        sendJson(res, 400, { error: "没有读取到上传文件" });
        return true;
      }
      fs.mkdirSync(path.dirname(files.itemUpload), { recursive: true });
      fs.writeFileSync(files.itemUpload, uploaded.data);
      const items = parseItemWorkbook(files.itemUpload);
      writeJson(files.items, items);
      sendJson(res, 200, { filename: uploaded.filename, items });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : "解析道具表失败" });
    }
    return true;
  }

  if (pathname === "/local-api/notices" && req.method === "GET") {
    const defaults = [1, 2, 3].map((slot) => ({ slot, title: "", body: "", imagePath: "" }));
    const saved = readJson(files.notices, []);
    sendJson(res, 200, { notices: defaults.map((item) => ({ ...item, ...((Array.isArray(saved) ? saved : []).find((notice) => Number(notice.slot) === item.slot) ?? {}) })) });
    return true;
  }

  if (pathname === "/local-api/notices" && req.method === "POST") {
    const body = await readJsonBody(req);
    const notices = readJson(files.notices, []);
    const slot = Math.min(3, Math.max(1, Number(body.slot) || 1));
    const next = {
      slot,
      title: String(body.title ?? ""),
      body: String(body.body ?? ""),
      imagePath: String(body.imagePath ?? ""),
      updatedAt: new Date().toISOString().slice(0, 19).replace("T", " "),
    };
    writeJson(files.notices, [next, ...(Array.isArray(notices) ? notices : []).filter((notice) => Number(notice.slot) !== slot)]);
    sendJson(res, 200, { notice: next });
    return true;
  }

  if (pathname === "/local-api/mail-templates" && req.method === "GET") {
    sendJson(res, 200, { templates: readJson(files.mailTemplates, []) });
    return true;
  }

  if (pathname === "/local-api/mail-templates" && req.method === "POST") {
    const body = await readJsonBody(req);
    const templates = readJson(files.mailTemplates, []);
    const now = new Date().toISOString().slice(0, 19).replace("T", " ");
    const id = String(body.id ?? `t-${Date.now()}`);
    const existing = Array.isArray(templates) ? templates.find((template) => template.id === id) : null;
    const next = { id, name: String(body.name ?? "未命名模板"), title: String(body.title ?? ""), body: String(body.body ?? ""), contents: body.contents && typeof body.contents === "object" ? body.contents : undefined, createdAt: String(existing?.createdAt ?? now), updatedAt: now };
    writeJson(files.mailTemplates, [next, ...(Array.isArray(templates) ? templates : []).filter((template) => template.id !== id)]);
    sendJson(res, 200, { template: next });
    return true;
  }

  if (pathname === "/local-api/reward-templates" && req.method === "GET") {
    sendJson(res, 200, { templates: readJson(files.rewardTemplates, []) });
    return true;
  }

  if (pathname === "/local-api/reward-templates" && req.method === "POST") {
    const body = await readJsonBody(req);
    const templates = readJson(files.rewardTemplates, []);
    const now = new Date().toISOString().slice(0, 19).replace("T", " ");
    const id = String(body.id ?? `r-${Date.now()}`);
    const next = { id, title: String(body.title ?? "未命名奖励模板"), items: Array.isArray(body.items) ? body.items : [], createdAt: now, updatedAt: now };
    writeJson(files.rewardTemplates, [next, ...(Array.isArray(templates) ? templates : []).filter((template) => template.id !== id)]);
    sendJson(res, 200, { template: next });
    return true;
  }

  if (pathname === "/local-api/games" && req.method === "POST") {
    const body = await readJsonBody(req);
    const games = readGames();
    const name = String(body.name ?? "").trim();
    const serverName = String(body.serverName ?? "").trim();
    const serverUrl = String(body.serverUrl ?? "").trim();
    if (!name || !serverName || !serverUrl) {
      sendJson(res, 400, { error: "游戏名、区服和服务端地址不能为空" });
      return true;
    }
    if (games.some((item) => item.name === name && item.serverName === serverName)) {
      sendJson(res, 409, { error: "该游戏区服已存在" });
      return true;
    }
    const created = {
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
    return true;
  }

  const gameMatch = pathname.match(/^\/local-api\/games\/(\d+)$/);
  if (gameMatch && req.method === "PUT") {
    const id = Number(gameMatch[1]);
    const body = await readJsonBody(req);
    const games = readGames();
    const existing = games.find((game) => game.id === id);
    if (!existing) {
      sendJson(res, 404, { error: "游戏区服不存在" });
      return true;
    }
    const name = String(body.name ?? existing.name).trim();
    const serverName = String(body.serverName ?? existing.serverName).trim();
    const serverUrl = String(body.serverUrl ?? existing.serverUrl).trim();
    if (!name || !serverName || !serverUrl) {
      sendJson(res, 400, { error: "游戏名、区服和服务端地址不能为空" });
      return true;
    }
    if (games.some((item) => item.id !== id && item.name === name && item.serverName === serverName)) {
      sendJson(res, 409, { error: "该游戏区服已存在" });
      return true;
    }
    const previousKey = `${existing.name}/${existing.serverName}`;
    const nextKey = `${name}/${serverName}`;
    const updated = {
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
      writeAccounts(readAccounts().map((account) => ({ ...account, games: account.games.map((game) => (game === previousKey ? nextKey : game)) })));
    }
    sendJson(res, 200, { game: updated });
    return true;
  }

  if (gameMatch && req.method === "DELETE") {
    const id = Number(gameMatch[1]);
    const games = readGames();
    const deleted = games.find((game) => game.id === id);
    writeGames(games.filter((game) => game.id !== id));
    if (deleted) {
      const deletedKey = `${deleted.name}/${deleted.serverName}`;
      writeAccounts(readAccounts().map((account) => ({ ...account, games: account.games.filter((game) => game !== deletedKey) })));
    }
    sendJson(res, 200, { Result: 0, id });
    return true;
  }

  if (pathname === "/local-api/accounts" && req.method === "POST") {
    const body = await readJsonBody(req);
    const accounts = readAccounts();
    const account = String(body.account ?? "").trim();
    const password = String(body.password ?? "").trim();
    if (!account || !password) {
      sendJson(res, 400, { error: "账号和密码不能为空" });
      return true;
    }
    if (accounts.some((item) => item.account === account)) {
      sendJson(res, 409, { error: "账号已存在" });
      return true;
    }
    const created = {
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
    return true;
  }

  const accountMatch = pathname.match(/^\/local-api\/accounts\/(\d+)$/);
  if (accountMatch && req.method === "DELETE") {
    const id = Number(accountMatch[1]);
    writeAccounts(readAccounts().filter((account) => account.id !== id));
    sendJson(res, 200, { Result: 0, id });
    return true;
  }

  if (pathname === "/local-api/operator-login" && req.method === "POST") {
    const body = await readJsonBody(req);
    const account = String(body.account ?? "").trim();
    const password = String(body.password ?? "").trim();
    const found = readAccounts().find((item) => item.account === account && item.password === password && (item.status === "启用" || item.status === "鍚敤"));
    if (!found) {
      sendJson(res, 401, { error: "账号或密码错误" });
      return true;
    }
    const { password: _password, ...safeAccount } = found;
    sendJson(res, 200, { account: safeAccount });
    return true;
  }

  if (pathname === "/local-api/gm-login" && req.method === "POST") {
    const body = await readJsonBody(req);
    const account = String(body.account ?? "").trim();
    const password = String(body.password ?? "").trim();
    const operator = body.operator;
    try {
      const credentials = operator ? readAdminCredentials() : { account, password };
      if (!credentials) {
        sendJson(res, 428, { error: "请先使用游戏服务端管理员账号登录一次，再让子账号登录" });
        return true;
      }
      const token = await loginToGameServer(body.serverUrl, credentials);
      if (!operator) writeAdminCredentials({ account, password });
      sendJson(res, 200, { token, adminAccount: credentials.account });
    } catch (error) {
      sendJson(res, 401, { error: error instanceof Error ? error.message : "登录失败" });
    }
    return true;
  }

  if (pathname === "/local-api/gm-token" && req.method === "POST") {
    const body = await readJsonBody(req);
    const credentials = readAdminCredentials();
    if (!credentials) {
      sendJson(res, 428, { error: "请先使用游戏服务端管理员账号登录一次" });
      return true;
    }
    try {
      const token = await loginToGameServer(body.serverUrl, credentials);
      sendJson(res, 200, { token, adminAccount: credentials.account });
    } catch (error) {
      sendJson(res, 401, { error: error instanceof Error ? error.message : "获取 Token 失败" });
    }
    return true;
  }

  sendJson(res, 404, { error: "Not found" });
  return true;
}

async function proxyGmApi(req, res, pathname, search) {
  const targetUrl = `${gmServerTarget}${pathname.replace(/^\/gm-api/, "")}${search}`;
  const headers = { ...req.headers, host: new URL(gmServerTarget).host };
  const body = req.method === "GET" || req.method === "HEAD" ? undefined : await readBody(req);
  try {
    const upstream = await fetch(targetUrl, { method: req.method, headers, body, redirect: "manual" });
    res.writeHead(upstream.status, Object.fromEntries(upstream.headers.entries()));
    if (upstream.body) {
      const buffer = Buffer.from(await upstream.arrayBuffer());
      res.end(buffer);
    } else {
      res.end();
    }
  } catch (error) {
    sendJson(res, 502, { error: error instanceof Error ? error.message : "游戏服务器连接失败" });
  }
}

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
};

function serveStatic(res, pathname) {
  const normalized = decodeURIComponent(pathname === "/" ? "/index.html" : pathname);
  const file = path.resolve(distDir, `.${normalized}`);
  if (!file.startsWith(distDir)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }
  const finalFile = fs.existsSync(file) && fs.statSync(file).isFile() ? file : path.join(distDir, "index.html");
  if (!fs.existsSync(finalFile)) {
    res.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("dist 目录不存在，请先执行 npm run build:test 或 npm run build:prod");
    return;
  }
  res.writeHead(200, { "Content-Type": mimeTypes[path.extname(finalFile).toLowerCase()] || "application/octet-stream" });
  fs.createReadStream(finalFile).pipe(res);
}

const server = createServer(async (req, res) => {
  const requestUrl = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);
  if (requestUrl.pathname.startsWith("/local-api/")) {
    await handleLocalApi(req, res, requestUrl.pathname);
    return;
  }
  if (requestUrl.pathname.startsWith("/gm-api")) {
    await proxyGmApi(req, res, requestUrl.pathname, requestUrl.search);
    return;
  }
  serveStatic(res, requestUrl.pathname);
});

server.listen(port, host, () => {
  console.log(`GM Admin Panel ${portal} server running at http://${host}:${port}`);
  console.log(`GM server target: ${gmServerTarget}`);
});
