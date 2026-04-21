import express from "express";
import { WebSocketServer } from "ws";
import { createServer } from "node:http";
import { randomUUID } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { readFileSync } from "node:fs";
import { bridgeInfo, tools, getToolByName } from "./tools.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ADDIN_DIR = path.resolve(__dirname, "..", "addin");

function resolvePort() {
  const argv = process.argv.slice(2);
  for (let i = 0; i < argv.length; i++) {
    if ((argv[i] === "--port" || argv[i] === "-p") && argv[i + 1]) return Number(argv[i + 1]);
    if (argv[i].startsWith("--port=")) return Number(argv[i].slice("--port=".length));
  }
  if (process.env.WORDBRIDGE_PORT) return Number(process.env.WORDBRIDGE_PORT);
  return 3001;
}
const PORT = resolvePort();

function resolveHost() {
  const argv = process.argv.slice(2);
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--host" && argv[i + 1]) return argv[i + 1];
    if (argv[i].startsWith("--host=")) return argv[i].slice("--host=".length);
  }
  if (process.env.WORDBRIDGE_HOST) return process.env.WORDBRIDGE_HOST;
  return "127.0.0.1";
}
const HOST = resolveHost();
if (!Number.isInteger(PORT) || PORT <= 0 || PORT > 65535) {
  console.error(`[wordbridge] invalid port: ${PORT}`);
  process.exit(2);
}

const app = express();
app.use(express.json({ limit: "20mb" }));

app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

app.get("/addin/manifest.xml", (_req, res) => {
  const proto = _req.get("x-forwarded-proto") || _req.protocol;
  const origin = `${proto}://${_req.get("host")}`;
  const xml = readFileSync(path.join(ADDIN_DIR, "manifest.xml"), "utf8")
    .replaceAll("http://localhost:3001", origin);
  res.type("application/xml").send(xml);
});
app.use("/addin", express.static(ADDIN_DIR));

const httpServer = createServer(app);
const wss = new WebSocketServer({ server: httpServer, path: "/ws" });

const clients = new Set();
const pending = new Map();

wss.on("connection", (ws) => {
  clients.add(ws);
  console.log(`[ws] client connected (total=${clients.size})`);
  ws.on("close", () => {
    clients.delete(ws);
    console.log(`[ws] client disconnected (total=${clients.size})`);
  });
  ws.on("message", (raw) => {
    let msg;
    try { msg = JSON.parse(raw.toString()); } catch { return; }
    if (msg.type === "result" && msg.id && pending.has(msg.id)) {
      const { resolve } = pending.get(msg.id);
      pending.delete(msg.id);
      resolve(msg);
    } else if (msg.type === "hello") {
      console.log(`[ws] hello from add-in: ${msg.info || ""}`);
    } else if (msg.type === "log") {
      console.log(`[addin] ${msg.text}`);
    }
  });
});

function sendOp(op, timeoutMs = 20000) {
  return new Promise((resolve, reject) => {
    if (clients.size === 0) {
      return reject(new Error("no add-in connected — open the Word Bridge task pane in Word"));
    }
    const id = randomUUID();
    const payload = JSON.stringify({ type: "op", id, op });
    pending.set(id, { resolve, reject });
    for (const ws of clients) ws.send(payload);
    setTimeout(() => {
      if (pending.has(id)) {
        pending.delete(id);
        reject(new Error(`op timed out after ${timeoutMs}ms`));
      }
    }, timeoutMs);
  });
}

app.get("/status", (_req, res) => {
  res.json({ ok: true, connectedClients: clients.size, port: PORT });
});

app.get("/tools", (_req, res) => {
  res.json({
    ...bridgeInfo,
    baseUrl: `http://127.0.0.1:${PORT}`,
    toolCount: tools.length,
    tools,
  });
});

app.get("/tools/:name", (req, res) => {
  const t = getToolByName(req.params.name);
  if (!t) return res.status(404).json({ ok: false, error: `unknown tool: ${req.params.name}` });
  res.json(t);
});

app.get("/", (_req, res) => {
  const proto = _req.get("x-forwarded-proto") || _req.protocol;
  const origin = `${proto}://${_req.get("host")}`;
  res.type("text/plain").send(
    [
      `wordbridge ${bridgeInfo.version} — Word live-editing bridge`,
      ``,
      `Endpoints:`,
      `  GET  /status        bridge health + connected add-in clients`,
      `  GET  /tools         full tool catalog (JSON Schema) for LLM callers`,
      `  GET  /tools/<name>  one tool`,
      `  POST /op            execute one op          body: { kind, ... }`,
      `  POST /ops           execute a batch of ops  body: [ {...}, ... ]`,
      `  GET  /addin/taskpane.html   Office.js task pane served to Word`,
      `  WS   /ws            task-pane connection (internal)`,
      ``,
      `Word Add-in:`,
      `  Manifest URL: ${origin}/addin/manifest.xml`,
      ``,
      `Sideload the add-in in Microsoft Word:`,
      ``,
      `  Option A — Upload via Word UI:`,
      `    1. Open Word > Insert > Get Add-ins (or Add-ins > My Add-ins)`,
      `    2. Click "Upload My Add-in" (under Manage My Add-ins or via the dropdown)`,
      `    3. Browse and upload the manifest URL or downloaded manifest.xml`,
      `    4. The Word Bridge task pane will appear on the right`,
      ``,
      `  Option B — Install via manifest folder (no UI needed):`,
      ``,
      `    macOS:`,
      `      1. Create the wef folder if it does not exist:`,
      `         mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef`,
      `      2. Download the manifest into that folder:`,
      `         curl -o ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml \\`,
      `              ${origin}/addin/manifest.xml`,
      `      3. Restart Word — the add-in will appear in Insert > My Add-ins`,
      ``,
      `    Windows:`,
      `      1. Create the WEF folder if it does not exist:`,
      `         %LOCALAPPDATA%\\Microsoft\\Office\\16.0\\WEF\\`,
      `      2. Save the manifest into that folder:`,
      `         curl -o "%LOCALAPPDATA%\\Microsoft\\Office\\16.0\\WEF\\manifest.xml" ^`,
      `              ${origin}/addin/manifest.xml`,
      `      3. Restart Word — the add-in will appear in Insert > My Add-ins`,
      ``,
      `  The add-in connects to this server via WebSocket automatically.`,
      ``,
      `Start with: curl ${origin}/tools | jq .`,
      `LLM callers: read /tools, pick a tool, POST its example to /op.`,
    ].join("\n"),
  );
});

app.post("/op", async (req, res) => {
  try {
    const result = await sendOp(req.body);
    res.json(result);
  } catch (err) {
    res.status(502).json({ ok: false, error: err.message });
  }
});

app.post("/ops", async (req, res) => {
  const ops = Array.isArray(req.body) ? req.body : req.body?.ops;
  if (!Array.isArray(ops)) return res.status(400).json({ ok: false, error: "body must be an array of ops or { ops: [...] }" });
  const results = [];
  for (const op of ops) {
    try {
      results.push(await sendOp(op));
    } catch (err) {
      results.push({ ok: false, error: err.message, op });
      if (req.query.stopOnError === "1") break;
    }
  }
  res.json({ ok: results.every((r) => r.ok !== false), results });
});

httpServer.on("upgrade", (req, socket, head) => {
  console.log("[http] upgrade request:", req.url, req.headers.upgrade);
});

httpServer.listen(PORT, HOST, () => {
  console.log(`[wordbridge] listening on http://${HOST}:${PORT}`);
  console.log(`[wordbridge] task pane: http://127.0.0.1:${PORT}/addin/taskpane.html`);
  console.log(`[wordbridge] status:    http://127.0.0.1:${PORT}/status`);
  console.log(`[wordbridge] tools:     http://127.0.0.1:${PORT}/tools`);
});
