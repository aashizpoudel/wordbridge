/* global Office, PowerPoint */

const WS_URL = `${window.location.protocol === "https:" ? "wss" : "ws"}://${window.location.host}/ws`;
let ws = null;
let reconnectTimer = null;

const logEl = () => document.getElementById("log");
const statusEl = () => document.getElementById("status");

function log(...args) {
  const line = args.map((a) => (typeof a === "string" ? a : JSON.stringify(a))).join(" ");
  const el = logEl();
  if (el) {
    el.textContent += `[${new Date().toLocaleTimeString()}] ${line}\n`;
    el.scrollTop = el.scrollHeight;
  }
  try { ws && ws.readyState === 1 && ws.send(JSON.stringify({ type: "log", text: line })); } catch {}
  console.log("[pptbridge]", line);
}

function serializeError(err) {
  if (!err) return "null";
  const out = {};
  for (const k of ["name", "code", "message", "stack"]) if (err[k] !== undefined) out[k] = err[k];
  if (err.debugInfo) {
    out.debugInfo = {};
    for (const k of ["code", "message", "errorLocation", "statement", "surroundingStatements", "fullStatements"]) {
      if (err.debugInfo[k] !== undefined) out.debugInfo[k] = err.debugInfo[k];
    }
  }
  try { return JSON.stringify(out, null, 2); } catch { return String(err); }
}

function setStatus(text, cls) {
  const el = statusEl();
  if (!el) return;
  el.textContent = text;
  el.className = `pill ${cls}`;
}

function connect() {
  if (ws && (ws.readyState === 0 || ws.readyState === 1)) return;
  setStatus("connecting…", "warn");
  log("connecting to: " + WS_URL);
  try {
    ws = new WebSocket(WS_URL);
  } catch (err) {
    log("WebSocket constructor error: " + (err.message || String(err)));
    setStatus("error", "err");
    scheduleReconnect();
    return;
  }
  ws.onopen = () => {
    setStatus("connected", "ok");
    log("ws connected");
    ws.send(JSON.stringify({ type: "hello", kind: "powerpoint", info: "PowerPoint task pane ready" }));
  };
  ws.onclose = () => {
    setStatus("disconnected", "warn");
    log("ws closed, retrying in 2s");
    scheduleReconnect();
  };
  ws.onerror = (ev) => {
    log("ws error: " + JSON.stringify({type: ev.type, url: WS_URL}));
    setStatus("error", "err");
  };
  ws.onmessage = async (evt) => {
    let msg;
    try { msg = JSON.parse(evt.data); } catch { return; }
    if (msg.type === "welcome") {
      log("assigned clientId: " + msg.clientId);
      return;
    }
    if (msg.type !== "op") return;
    const { id, op } = msg;
    try {
      const result = await applyOp(op);
      ws.send(JSON.stringify({ type: "result", id, ok: true, result }));
    } catch (err) {
      const detail = serializeError(err);
      log("op failed:\n" + detail);
      ws.send(JSON.stringify({ type: "result", id, ok: false, error: err.message || String(err), detail }));
    }
  };
}

function scheduleReconnect() {
  if (reconnectTimer) return;
  reconnectTimer = setTimeout(() => {
    reconnectTimer = null;
    connect();
  }, 2000);
}

async function applyOp(op) {
  log("op:", op.kind);
  return PowerPoint.run(async (context) => {
    switch (op.kind) {
      case "ping": {
        return { pong: true };
      }

      case "getSlides": {
        const { limit = 0 } = op;
        const slides = context.presentation.slides;
        slides.load("items/id,items/index");
        await context.sync();
        const items = [];
        for (let i = 0; i < slides.items.length; i++) {
          const slide = slides.items[i];
          items.push({ id: slide.id, index: i });
          if (limit > 0 && items.length >= limit) break;
        }
        return { total: slides.items.length, returned: items.length, slides: items };
      }

      case "getOutline": {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        const outline = [];
        for (let i = 0; i < slides.items.length; i++) {
          const shapes = slides.items[i].shapes;
          shapes.load("items/name,items/type");
          await context.sync();
          let title = null;
          for (const shape of shapes.items) {
            if (shape.name && shape.name.toLowerCase().includes("title")) {
              try {
                const tf = shape.textFrame;
                tf.load("textRange/text");
                await context.sync();
                title = tf.textRange.text;
              } catch { /* no text frame */ }
              break;
            }
          }
          outline.push({ index: i, title });
        }
        return { slideCount: slides.items.length, outline };
      }

      case "getSlideText": {
        const { slideIndex = 0, textLimit = 4000 } = op;
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`slide index out of range: ${slideIndex} (total: ${slides.items.length})`);
        }
        const slide = slides.items[slideIndex];
        const shapes = slide.shapes;
        shapes.load("items/id,items/name,items/type");
        await context.sync();
        const texts = [];
        for (const shape of shapes.items) {
          try {
            const tf = shape.textFrame;
            tf.load("textRange/text");
            await context.sync();
            texts.push({
              shapeId: shape.id,
              shapeName: shape.name,
              text: tf.textRange.text.slice(0, textLimit),
            });
          } catch { /* shape has no text frame */ }
        }
        return { slideIndex, shapeCount: shapes.items.length, texts };
      }

      case "setShapeText": {
        const { slideIndex = 0, shapeName, shapeId, text } = op;
        if (text === undefined) throw new Error("setShapeText: text required");
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`slide index out of range: ${slideIndex} (total: ${slides.items.length})`);
        }
        const slide = slides.items[slideIndex];
        const shapes = slide.shapes;
        shapes.load("items/id,items/name,items/type");
        await context.sync();
        let target = null;
        for (const shape of shapes.items) {
          if (shapeId && shape.id === shapeId) { target = shape; break; }
          if (shapeName && shape.name === shapeName) { target = shape; break; }
        }
        if (!target) throw new Error(`shape not found: ${shapeId || shapeName}`);
        target.textFrame.textRange.text = text;
        await context.sync();
        return { ok: true, slideIndex, shapeName: target.name };
      }

      case "findReplaceAll": {
        const { find, replace, matchCase = false } = op;
        if (!find) throw new Error("findReplaceAll: find required");
        if (replace === undefined) throw new Error("findReplaceAll: replace required");
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        let totalMatched = 0, totalReplaced = 0;
        for (const slide of slides.items) {
          const shapes = slide.shapes;
          shapes.load("items/type,items/name");
          await context.sync();
          for (const shape of shapes.items) {
            try {
              const tf = shape.textFrame;
              tf.load("textRange/text");
              await context.sync();
              const text = tf.textRange.text;
              const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
              const regex = new RegExp(escaped, matchCase ? "g" : "gi");
              const matches = text.match(regex);
              if (matches && matches.length > 0) {
                totalMatched += matches.length;
                tf.textRange.text = text.replace(regex, replace);
                totalReplaced += matches.length;
                await context.sync();
              }
            } catch { /* shape has no text frame */ }
          }
        }
        return { matched: totalMatched, replaced: totalReplaced };
      }

      case "addSlide": {
        context.presentation.slides.add();
        await context.sync();
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        return { ok: true, totalSlides: slides.items.length };
      }

      case "addSlideWithLayout": {
        const { layoutId, slideMasterId, index } = op;
        const options = {};
        if (layoutId) options.layoutId = layoutId;
        if (slideMasterId) options.slideMasterId = slideMasterId;
        context.presentation.slides.add(options);
        await context.sync();
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (index !== undefined) {
          slides.items[slides.items.length - 1].moveTo(index);
          await context.sync();
        }
        return { ok: true, totalSlides: slides.items.length };
      }

      case "deleteSlide": {
        const { slideIndex } = op;
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`slide index out of range: ${slideIndex} (total: ${slides.items.length})`);
        }
        slides.items[slideIndex].delete();
        await context.sync();
        return { ok: true, deleted: slideIndex };
      }

      case "duplicateSlide": {
        const { slideIndex = 0, targetIndex } = op;
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`slide index out of range: ${slideIndex} (total: ${slides.items.length})`);
        }
        const slide = slides.items[slideIndex];
        const exportResult = slide.exportAsBase64();
        await context.sync();
        const insertOptions = {};
        if (targetIndex !== undefined) insertOptions.index = targetIndex;
        context.presentation.insertSlidesFromBase64(exportResult.value, insertOptions);
        await context.sync();
        slides.load("items");
        await context.sync();
        return { ok: true, totalSlides: slides.items.length };
      }

      case "moveSlide": {
        const { slideIndex, targetIndex } = op;
        if (slideIndex === undefined) throw new Error("moveSlide: slideIndex required");
        if (targetIndex === undefined) throw new Error("moveSlide: targetIndex required");
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`slide index out of range: ${slideIndex} (total: ${slides.items.length})`);
        }
        const slide = slides.getItemAt(slideIndex);
        slide.moveTo(targetIndex);
        await context.sync();
        return { ok: true, from: slideIndex, to: targetIndex };
      }

      case "getShapes": {
        const { slideIndex = 0 } = op;
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`slide index out of range: ${slideIndex} (total: ${slides.items.length})`);
        }
        const shapes = slides.items[slideIndex].shapes;
        shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
        await context.sync();
        const items = shapes.items.map((s) => ({
          id: s.id,
          name: s.name,
          type: s.type,
          left: s.left,
          top: s.top,
          width: s.width,
          height: s.height,
        }));
        return { slideIndex, shapes: items };
      }

      case "insertImage": {
        const { slideIndex = 0, base64, left = 50, top = 50, width = 400, height = 300 } = op;
        if (!base64) throw new Error("insertImage: base64 required");
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`slide index out of range: ${slideIndex} (total: ${slides.items.length})`);
        }
        const slide = slides.items[slideIndex];
        slide.shapes.addImage(base64, { left, top, width, height });
        await context.sync();
        return { ok: true, slideIndex };
      }

      case "replaceImage": {
        const { slideIndex = 0, shapeName, shapeId, base64 } = op;
        if (!base64) throw new Error("replaceImage: base64 required");
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`slide index out of range: ${slideIndex} (total: ${slides.items.length})`);
        }
        const shapes = slides.items[slideIndex].shapes;
        shapes.load("items/id,items/name,items/type");
        await context.sync();
        let target = null;
        for (const shape of shapes.items) {
          if (shapeId && shape.id === shapeId) { target = shape; break; }
          if (shapeName && shape.name === shapeName) { target = shape; break; }
        }
        if (!target) throw new Error(`shape not found: ${shapeId || shapeName}`);
        target.fill.setImage(base64);
        await context.sync();
        return { ok: true, shapeName: target.name, shapeId: target.id };
      }

      case "getSlideImage": {
        const { slideIndex = 0, height = 300 } = op;
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`slide index out of range: ${slideIndex} (total: ${slides.items.length})`);
        }
        const imageResult = slides.items[slideIndex].getImageAsBase64({ height });
        await context.sync();
        return { slideIndex, base64: imageResult.value, height };
      }

      case "getLayouts": {
        const masters = context.presentation.slideMasters;
        masters.load("items/id,items/name,items/layouts/items/id,items/layouts/items/name");
        await context.sync();
        const result = masters.items.map((m) => ({
          id: m.id,
          name: m.name,
          layouts: m.layouts.items.map((l) => ({ id: l.id, name: l.name })),
        }));
        return { masters: result };
      }

      default:
        throw new Error(`unknown op: ${op.kind}`);
    }
  });
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    log("not running inside PowerPoint — host=", info.host);
    return;
  }
  log("Office ready, host=PowerPoint");
  connect();

  document.getElementById("pingBtn").addEventListener("click", async () => {
    try {
      await applyOp({ kind: "ping" });
      log("ping ok");
    } catch (err) {
      log("ping err:\n" + serializeError(err));
    }
  });
  document.getElementById("clearLogBtn").addEventListener("click", () => {
    logEl().textContent = "";
  });
});
