import express from "express";
import { WebSocketServer } from "ws";
import { createServer } from "node:http";
import { randomUUID } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { readFileSync } from "node:fs";
import { bridgeInfo, tools, getToolByName, getToolsByHost } from "./tools.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ADDIN_DIR = path.resolve(__dirname, "..", "addin");

// Random word list for client IDs
const WORDS = [
  "alpha", "bravo", "charlie", "delta", "echo", "foxtrot", "golf", "hotel",
  "india", "juliet", "kilo", "lima", "mike", "november", "oscar", "papa",
  "quebec", "romeo", "sierra", "tango", "uniform", "victor", "whiskey",
  "xray", "yankee", "zulu", "anchor", "breeze", "coral", "drift", "ember",
  "flint", "grove", "haze", "ivory", "jade", "knot", "lumen", "maple",
  "nexus", "opal", "prism", "quartz", "ridge", "spark", "tide", "umbra",
  "vault", "wren", "zenith", "arrow", "bloom", "crest", "dusk", "fern",
  "glyph", "haven", "isle", "jewel", "kayak", "lotus", "mirth", "nova",
];

function generateClientId() {
  const w1 = WORDS[Math.floor(Math.random() * WORDS.length)];
  const w2 = WORDS[Math.floor(Math.random() * WORDS.length)];
  return `${w1}-${w2}`;
}

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

app.get("/addin/:app/manifest.xml", (req, res) => {
  const sub = req.params.app;
  const file = path.join(ADDIN_DIR, sub, "manifest.xml");
  try { readFileSync(file); } catch { return res.status(404).json({ error: `no manifest for ${sub}` }); }
  const proto = req.get("x-forwarded-proto") || req.protocol;
  const origin = `${proto}://${req.get("host")}`;
  const xml = readFileSync(file, "utf8").replaceAll("http://localhost:3001", origin);
  res.type("application/xml").send(xml);
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

// Client registry: Map<clientId, { ws, kind, connectedAt }>
const clients = new Map();
const pending = new Map();

function getClientsByKind(kind) {
  const result = [];
  for (const [id, client] of clients) {
    if (client.kind === kind) result.push({ id, ...client });
  }
  return result;
}

function getClientById(clientId) {
  return clients.get(clientId) || null;
}

wss.on("connection", (ws) => {
  const clientId = generateClientId();
  const client = { ws, kind: null, connectedAt: new Date().toISOString() };
  clients.set(clientId, client);
  console.log(`[ws] client connected: ${clientId} (total=${clients.size})`);

  // Send the assigned client ID to the add-in
  ws.send(JSON.stringify({ type: "welcome", clientId }));

  ws.on("close", () => {
    clients.delete(clientId);
    console.log(`[ws] client disconnected: ${clientId} (total=${clients.size})`);
  });

  ws.on("message", (raw) => {
    let msg;
    try { msg = JSON.parse(raw.toString()); } catch { return; }
    if (msg.type === "result" && msg.id && pending.has(msg.id)) {
      const { resolve } = pending.get(msg.id);
      pending.delete(msg.id);
      resolve(msg);
    } else if (msg.type === "hello") {
      // Update client kind from hello message
      if (msg.kind) {
        client.kind = msg.kind;
      }
      console.log(`[ws] hello from ${clientId}: kind=${client.kind}, info=${msg.info || ""}`);
    } else if (msg.type === "log") {
      console.log(`[${clientId}] ${msg.text}`);
    }
  });
});

function sendOp(op, { clientId, kind, timeoutMs = 20000 } = {}) {
  return new Promise((resolve, reject) => {
    // Resolve target client(s)
    let targets = [];

    if (clientId) {
      const client = getClientById(clientId);
      if (!client) return reject(new Error(`client not found: ${clientId}`));
      targets = [client];
    } else if (kind) {
      targets = getClientsByKind(kind);
      if (targets.length === 0) {
        return reject(new Error(`no ${kind} add-in connected — open the Bridge task pane in ${kind}`));
      }
      // Send to the first matching client
      targets = [targets[0]];
    } else {
      // No target specified — send to first available client
      if (clients.size === 0) {
        return reject(new Error("no add-in connected — open a Bridge task pane in an Office app"));
      }
      targets = [clients.values().next().value];
    }

    const id = randomUUID();
    const payload = JSON.stringify({ type: "op", id, op });
    pending.set(id, { resolve, reject });
    for (const t of targets) {
      const ws = t.ws || t;
      ws.send(payload);
    }
    setTimeout(() => {
      if (pending.has(id)) {
        pending.delete(id);
        reject(new Error(`op timed out after ${timeoutMs}ms`));
      }
    }, timeoutMs);
  });
}

app.get("/status", (_req, res) => {
  const summary = { word: 0, excel: 0, powerpoint: 0, unknown: 0 };
  const clientList = [];
  for (const [id, client] of clients) {
    const k = client.kind || "unknown";
    summary[k] = (summary[k] || 0) + 1;
    clientList.push({ clientId: id, kind: client.kind, connectedAt: client.connectedAt });
  }
  res.json({ ok: true, connectedClients: clients.size, summary, clients: clientList, port: PORT });
});

app.get("/tools", (req, res) => {
  const host = typeof req.query.host === "string" ? req.query.host : undefined;
  const filtered = getToolsByHost(host);
  res.json({
    ...bridgeInfo,
    baseUrl: `http://127.0.0.1:${PORT}`,
    host: host || undefined,
    toolCount: filtered.length,
    tools: filtered,
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
  const template = readFileSync(path.join(__dirname, "index.md"), "utf8");
  const body = template
    .replaceAll("{{origin}}", origin)
    .replaceAll("{{version}}", bridgeInfo.version);
  res.type("text/plain").send(body);
});

function inferHost(op) {
  if (!op || !op.kind) return undefined;
  const t = getToolByName(op.kind);
  if (!t) return undefined;
  return t.host && t.host !== "any" ? t.host : undefined;
}

app.post("/op", async (req, res) => {
  try {
    const { clientId, target, ...op } = req.body;
    const kind = target || inferHost(op);
    const result = await sendOp(op, { clientId, kind });
    res.json(result);
  } catch (err) {
    res.status(502).json({ ok: false, error: err.message });
  }
});

app.post("/ops", async (req, res) => {
  const body = req.body;
  const ops = Array.isArray(body) ? body : body?.ops;
  const clientId = body?.clientId;
  const target = body?.target;
  if (!Array.isArray(ops)) return res.status(400).json({ ok: false, error: "body must be an array of ops or { ops: [...], clientId?, target? }" });
  const results = [];
  for (const op of ops) {
    try {
      const opClientId = op.clientId || clientId;
      const opTarget = op.target || target || inferHost(op);
      results.push(await sendOp(op, { clientId: opClientId, kind: opTarget }));
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
