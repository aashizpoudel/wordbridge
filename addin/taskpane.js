/* global Office, Word */

// Derive WS URL from the page origin so the task pane adapts to whichever
// port served it — changing the bridge port only requires updating the
// manifest SourceLocation (and a Word relaunch to pick it up).
const WS_URL = `${window.location.protocol === "https:" ? "wss" : "ws"}://${window.location.host}/ws`;
let ws = null;
let reconnectTimer = null;

const logEl = () => document.getElementById("log");
const statusEl = () => document.getElementById("status");
const trackEl = () => document.getElementById("trackStatus");

function log(...args) {
  const line = args.map((a) => (typeof a === "string" ? a : JSON.stringify(a))).join(" ");
  const el = logEl();
  if (el) {
    el.textContent += `[${new Date().toLocaleTimeString()}] ${line}\n`;
    el.scrollTop = el.scrollHeight;
  }
  try { ws && ws.readyState === 1 && ws.send(JSON.stringify({ type: "log", text: line })); } catch {}
  console.log("[wordbridge]", line);
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

function setTrackStatus(text, cls) {
  const el = trackEl();
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
    ws.send(JSON.stringify({ type: "hello", info: "Word task pane ready" }));
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

async function readTrackState(context) {
  // On Mac Word, Document.changeTrackingMode is documented but neither readable
  // nor writable via Word.js. Try both property-name forms; if both fail,
  // degrade to manual mode and surface a banner so the user toggles in Review.
  const tryRead = async () => {
    try {
      context.document.load("changeTrackingMode");
      await context.sync();
      return context.document.changeTrackingMode;
    } catch { return undefined; }
  };
  let mode = await tryRead();
  if (mode === undefined) {
    try {
      const d = context.document;
      d.load({ select: "changeTrackingMode" });
      await context.sync();
      mode = d.changeTrackingMode;
    } catch { mode = undefined; }
  }
  const warn = document.getElementById("trackWarn");
  if (mode === undefined) {
    setTrackStatus("manual — verify Review ribbon", "warn");
    if (warn) warn.style.display = "block";
    return { mode: null, on: null, manual: true };
  }
  const on = mode === "TrackAll" || mode === "TrackMineOnly";
  setTrackStatus(on ? `ON (${mode})` : `OFF (${mode})`, on ? "ok" : "err");
  if (warn) warn.style.display = on ? "none" : "block";
  return { mode, on };
}

async function applyOp(op) {
  log("op:", op.kind);
  return Word.run(async (context) => {
    switch (op.kind) {
      case "ping": {
        return { pong: true };
      }
      case "setTrackChanges": {
        // Many Word builds (notably Mac) do not allow setting this property via the JS API.
        // Try it, but fall back to instructing the user to toggle manually.
        try {
          context.document.changeTrackingMode = op.on ? "TrackAll" : "Off";
          await context.sync();
          return await readTrackState(context);
        } catch (err) {
          log("setTrackChanges not supported on this Word build — toggle manually via Review ribbon.");
          return await readTrackState(context);
        }
      }
      case "getTrackChanges": {
        return await readTrackState(context);
      }
      case "findReplace": {
        const { find, replace, matchCase = false, matchWholeWord = false, maxReplacements = 0 } = op;
        const results = context.document.body.search(find, { matchCase, matchWholeWord });
        results.load("items");
        await context.sync();
        const n = maxReplacements > 0 ? Math.min(results.items.length, maxReplacements) : results.items.length;
        for (let i = 0; i < n; i++) {
          results.items[i].insertText(replace, Word.InsertLocation.replace);
        }
        await context.sync();
        return { matched: results.items.length, replaced: n };
      }
      case "insertAfterText": {
        const { anchor, text, asParagraph = false, style } = op;
        const results = context.document.body.search(anchor, { matchCase: true });
        results.load("items");
        await context.sync();
        if (results.items.length === 0) throw new Error(`anchor not found: ${anchor}`);
        const target = results.items[0];
        if (asParagraph) {
          const p = target.paragraphs.getFirst().insertParagraph(text, Word.InsertLocation.after);
          if (style) p.style = style;
        } else {
          target.insertText(text, Word.InsertLocation.after);
        }
        await context.sync();
        return { anchorMatches: results.items.length };
      }
      case "insertOoxml": {
        const { anchor, ooxml, location = "after" } = op;
        const results = context.document.body.search(anchor, { matchCase: true });
        results.load("items");
        await context.sync();
        if (results.items.length === 0) throw new Error(`anchor not found: ${anchor}`);
        const target = results.items[0];
        const locMap = { after: Word.InsertLocation.after, before: Word.InsertLocation.before, replace: Word.InsertLocation.replace };
        const loc = locMap[location];
        if (!loc) throw new Error(`bad location: ${location}`);
        target.insertOoxml(ooxml, loc);
        await context.sync();
        return { ok: true };
      }
      case "deleteText": {
        const { find, maxDeletions = 0 } = op;
        const results = context.document.body.search(find, { matchCase: true });
        results.load("items");
        await context.sync();
        const n = maxDeletions > 0 ? Math.min(results.items.length, maxDeletions) : results.items.length;
        for (let i = 0; i < n; i++) results.items[i].delete();
        await context.sync();
        return { matched: results.items.length, deleted: n };
      }
      case "setParagraphStyle": {
        const { anchor, style } = op;
        const results = context.document.body.search(anchor, { matchCase: true });
        results.load("items");
        await context.sync();
        if (results.items.length === 0) throw new Error(`anchor not found: ${anchor}`);
        const p = results.items[0].paragraphs.getFirst();
        p.style = style;
        await context.sync();
        return { ok: true };
      }
      case "getText": {
        const { limit = 4000 } = op;
        const body = context.document.body;
        body.load("text");
        await context.sync();
        return { text: body.text.slice(0, limit), totalLength: body.text.length };
      }
      case "getParagraphs": {
        const { styleFilter, limit = 0, includeEmpty = false, textLimit = 500 } = op;
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load(["text", "style", "styleBuiltIn"]);
        await context.sync();
        const items = [];
        for (let i = 0; i < paragraphs.items.length; i++) {
          const p = paragraphs.items[i];
          const text = p.text || "";
          if (!includeEmpty && text.trim() === "") continue;
          if (styleFilter && p.style !== styleFilter && p.styleBuiltIn !== styleFilter) continue;
          items.push({
            index: i,
            style: p.style,
            styleBuiltIn: p.styleBuiltIn,
            text: text.length > textLimit ? text.slice(0, textLimit) + "…" : text,
            length: text.length,
          });
          if (limit > 0 && items.length >= limit) break;
        }
        return { total: paragraphs.items.length, returned: items.length, paragraphs: items };
      }
      case "getOoxml": {
        const { anchor, scope = "body", charLimit = 200000 } = op;
        let range;
        if (anchor) {
          const results = context.document.body.search(anchor, { matchCase: true });
          results.load("items");
          await context.sync();
          if (results.items.length === 0) throw new Error(`anchor not found: ${anchor}`);
          range = results.items[0].paragraphs.getFirst().getRange();
        } else if (scope === "body") {
          range = context.document.body.getRange();
        } else if (scope === "selection") {
          range = context.document.getSelection();
        } else {
          throw new Error(`bad scope: ${scope}`);
        }
        const ooxmlResult = range.getOoxml();
        await context.sync();
        const value = ooxmlResult.value || "";
        return {
          ooxml: value.length > charLimit ? value.slice(0, charLimit) : value,
          totalLength: value.length,
          truncated: value.length > charLimit,
        };
      }
      case "insertImage": {
        // Insert an inline picture at/around an anchor.
        //   anchor: text to search for
        //   base64: image bytes, base64-encoded (PNG/JPEG)
        //   location: "after" | "before" | "replace" (default "after")
        //     - "after"/"before" insert a new paragraph next to the anchor's paragraph
        //       and place the picture inside it
        //     - "replace" replaces the anchor text itself with the picture inline
        //   widthPoints?: if set, resize the picture width (height scales proportionally)
        //   alignment?: "left"|"center"|"right" applied to the containing paragraph
        const { anchor, base64, location = "after", widthPoints, alignment } = op;
        if (!anchor) throw new Error("insertImage: anchor required");
        if (!base64) throw new Error("insertImage: base64 required");

        const results = context.document.body.search(anchor, { matchCase: true });
        results.load("items");
        await context.sync();
        if (results.items.length === 0) throw new Error(`anchor not found: ${anchor}`);
        const target = results.items[0];

        let picture;
        if (location === "replace") {
          picture = target.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
        } else {
          const parent = target.paragraphs.getFirst();
          const loc = location === "before" ? Word.InsertLocation.before : Word.InsertLocation.after;
          const newPara = parent.insertParagraph("", loc);
          if (alignment === "left") newPara.alignment = Word.Alignment.left;
          else if (alignment === "right") newPara.alignment = Word.Alignment.right;
          else newPara.alignment = Word.Alignment.centered;
          picture = newPara.insertInlinePictureFromBase64(base64, Word.InsertLocation.start);
        }
        if (widthPoints && Number.isFinite(widthPoints)) {
          picture.width = Number(widthPoints);
        }
        await context.sync();
        return { ok: true, widthPoints: widthPoints ?? null, location };
      }
      case "replaceParagraphByIndex": {
        // Replace the entire content of a paragraph identified by its index
        // in body.paragraphs. Preserves paragraph style and properties; clears
        // inline math, fields, tracked-change markers, etc. inside the paragraph.
        const { index, newText } = op;
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items");
        await context.sync();
        if (index < 0 || index >= paragraphs.items.length) {
          throw new Error(`paragraph index out of range: ${index} (total: ${paragraphs.items.length})`);
        }
        const p = paragraphs.items[index];
        p.getRange().insertText(newText, Word.InsertLocation.replace);
        await context.sync();
        return { index, totalParagraphs: paragraphs.items.length };
      }
      case "getTrackedChanges": {
        const { timeoutMs = 3000, textLimit = 200 } = op;
        const inner = (async () => {
          const tc = context.document.body.getTrackedChanges();
          tc.load(["items/author", "items/date", "items/type", "items/text"]);
          await context.sync();
          return tc.items.map((c, i) => ({
            index: i,
            type: c.type,
            author: c.author,
            date: c.date,
            text: (c.text || "").length > textLimit ? c.text.slice(0, textLimit) + "…" : c.text,
            length: (c.text || "").length,
          }));
        })();
        // getTrackedChanges is known to hang on some Mac builds (office-js #5535, #6514).
        const timeout = new Promise((_, reject) =>
          setTimeout(() => reject(new Error(`getTrackedChanges timed out after ${timeoutMs}ms (known Mac Word.js bug)`)), timeoutMs)
        );
        const changes = await Promise.race([inner, timeout]);
        return { count: changes.length, changes };
      }
      case "reviewChanges": {
        // Unified accept/reject with multiple selector strategies.
        //   action: "accept" | "reject"
        //   selector:
        //     { kind: "all" }                                  — every change in body
        //     { kind: "index", value: N }                      — Nth change (0-based) in body
        //     { kind: "text", value: "...", matchCase?: bool }  — changes whose text contains substring
        //     { kind: "paragraph", anchor: "..." }             — changes inside the paragraph containing anchor
        //   authorFilter?: "name"                              — further narrow to one author
        //   timeoutMs?: default 15000
        //   maxMatches?: 0 for all
        const { action, selector, authorFilter, timeoutMs = 15000, maxMatches = 0 } = op;
        if (action !== "accept" && action !== "reject") throw new Error(`bad action: ${action}`);
        const doAction = (change) => action === "accept" ? change.accept() : change.reject();

        const inner = (async () => {
          let source, sourceScope;
          if (selector?.kind === "paragraph") {
            const anchor = selector.anchor;
            if (!anchor) throw new Error("selector paragraph requires anchor");
            const results = context.document.body.search(anchor, { matchCase: true });
            results.load("items");
            await context.sync();
            if (results.items.length === 0) throw new Error(`anchor not found: ${anchor}`);
            const para = results.items[0].paragraphs.getFirst();
            source = para.getRange().getTrackedChanges();
            sourceScope = "paragraph";
          } else {
            source = context.document.body.getTrackedChanges();
            sourceScope = "body";
          }
          source.load(["items/author", "items/type", "items/text"]);
          await context.sync();

          const all = source.items || [];
          const candidates = [];
          if (!selector || selector.kind === "all" || selector.kind === "paragraph") {
            for (const c of all) candidates.push(c);
          } else if (selector.kind === "index") {
            const i = Number(selector.value);
            if (i < 0 || i >= all.length) {
              throw new Error(`index out of range: ${i} (found ${all.length} changes in ${sourceScope})`);
            }
            candidates.push(all[i]);
          } else if (selector.kind === "text") {
            const needle = String(selector.value);
            const caseSensitive = !!selector.matchCase;
            const hay = caseSensitive ? (t => t) : (t => (t || "").toLowerCase());
            const n2 = caseSensitive ? needle : needle.toLowerCase();
            for (const c of all) {
              if ((hay(c.text) || "").includes(n2)) candidates.push(c);
            }
          } else {
            throw new Error(`unknown selector: ${selector.kind}`);
          }

          const filtered = authorFilter
            ? candidates.filter((c) => c.author === authorFilter)
            : candidates;
          const toTouch = maxMatches > 0 ? filtered.slice(0, maxMatches) : filtered;
          for (const c of toTouch) doAction(c);
          await context.sync();

          return {
            scope: sourceScope,
            totalInScope: all.length,
            matched: filtered.length,
            touched: toTouch.length,
            action,
          };
        })();

        const timeout = new Promise((_, reject) =>
          setTimeout(
            () => reject(new Error(`reviewChanges timed out after ${timeoutMs}ms (known Mac Word.js bug — try --paragraph scope or accept manually)`)),
            timeoutMs,
          ),
        );
        return await Promise.race([inner, timeout]);
      }
      default:
        throw new Error(`unknown op: ${op.kind}`);
    }
  });
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    log("not running inside Word — host=", info.host);
    return;
  }
  log("Office ready, host=Word");
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
  document.getElementById("refreshTrackBtn").addEventListener("click", async () => {
    try { await applyOp({ kind: "getTrackChanges" }); } catch (err) { log("refresh err:\n" + serializeError(err)); }
  });

  // initial track-changes probe
  applyOp({ kind: "getTrackChanges" }).catch((e) => log("init probe failed:\n" + serializeError(e)));
});
