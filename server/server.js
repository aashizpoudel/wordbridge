import express from "express";
import { WebSocketServer } from "ws";
import { createServer } from "node:http";
import { randomUUID } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";

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
if (!Number.isInteger(PORT) || PORT <= 0 || PORT > 65535) {
  console.error(`[wordbridge] invalid port: ${PORT}`);
  process.exit(2);
}

const app = express();
app.use(express.json({ limit: "4mb" }));

app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
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

httpServer.listen(PORT, "127.0.0.1", () => {
  console.log(`[wordbridge] listening on http://127.0.0.1:${PORT}`);
  console.log(`[wordbridge] task pane: http://127.0.0.1:${PORT}/addin/taskpane.html`);
  console.log(`[wordbridge] status:    http://127.0.0.1:${PORT}/status`);
});
