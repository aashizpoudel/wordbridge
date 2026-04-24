/* global Office, Excel */

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
  console.log("[excelbridge]", line);
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
    ws.send(JSON.stringify({ type: "hello", kind: "excel", info: "Excel task pane ready" }));
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

async function applyOp(op) {
  log("op:", op.kind);
  return Excel.run(async (context) => {
    switch (op.kind) {
      case "ping": {
        return { pong: true };
      }

      case "getSheets": {
        // List all worksheets
        const sheets = context.workbook.worksheets;
        sheets.load("items/id,items/name,items/position,items/visibility");
        await context.sync();
        const items = sheets.items.map((s) => ({
          id: s.id,
          name: s.name,
          position: s.position,
          visibility: s.visibility,
        }));
        return { count: items.length, sheets: items };
      }

      case "getRange": {
        // Read values from a range
        const { sheet, address = "A1:Z50", valuesOnly = false } = op;
        const ws = sheet
          ? context.workbook.worksheets.getItem(sheet)
          : context.workbook.worksheets.getActiveWorksheet();
        const range = ws.getRange(address);
        const loads = ["values", "address", "rowCount", "columnCount"];
        if (!valuesOnly) loads.push("formulas", "numberFormat");
        range.load(loads.join(","));
        await context.sync();
        const result = {
          address: range.address,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          values: range.values,
        };
        if (!valuesOnly) {
          result.formulas = range.formulas;
          result.numberFormat = range.numberFormat;
        }
        return result;
      }

      case "setRange": {
        // Write values to a range
        const { sheet, address, values } = op;
        if (!address) throw new Error("setRange: address required");
        if (!values) throw new Error("setRange: values required (2D array)");
        const ws = sheet
          ? context.workbook.worksheets.getItem(sheet)
          : context.workbook.worksheets.getActiveWorksheet();
        const range = ws.getRange(address);
        range.values = values;
        await context.sync();
        return { ok: true, address };
      }

      case "setFormulas": {
        // Write formulas to a range
        const { sheet, address, formulas } = op;
        if (!address) throw new Error("setFormulas: address required");
        if (!formulas) throw new Error("setFormulas: formulas required (2D array)");
        const ws = sheet
          ? context.workbook.worksheets.getItem(sheet)
          : context.workbook.worksheets.getActiveWorksheet();
        const range = ws.getRange(address);
        range.formulas = formulas;
        await context.sync();
        return { ok: true, address };
      }

      case "getCellValue": {
        // Read a single cell
        const { sheet, address = "A1" } = op;
        const ws = sheet
          ? context.workbook.worksheets.getItem(sheet)
          : context.workbook.worksheets.getActiveWorksheet();
        const range = ws.getRange(address);
        range.load("values,formulas,numberFormat,address");
        await context.sync();
        return {
          address: range.address,
          value: range.values[0][0],
          formula: range.formulas[0][0],
          numberFormat: range.numberFormat[0][0],
        };
      }

      case "setCellValue": {
        // Write a single cell
        const { sheet, address = "A1", value } = op;
        if (value === undefined) throw new Error("setCellValue: value required");
        const ws = sheet
          ? context.workbook.worksheets.getItem(sheet)
          : context.workbook.worksheets.getActiveWorksheet();
        const range = ws.getRange(address);
        range.values = [[value]];
        await context.sync();
        return { ok: true, address };
      }

      case "addSheet": {
        // Add a new worksheet
        const { name } = op;
        const sheet = name
          ? context.workbook.worksheets.add(name)
          : context.workbook.worksheets.add();
        sheet.load("id,name,position");
        await context.sync();
        return { ok: true, id: sheet.id, name: sheet.name, position: sheet.position };
      }

      case "deleteSheet": {
        // Delete a worksheet by name
        const { name } = op;
        if (!name) throw new Error("deleteSheet: name required");
        const sheet = context.workbook.worksheets.getItem(name);
        sheet.delete();
        await context.sync();
        return { ok: true, deleted: name };
      }

      case "formatRange": {
        // Apply formatting to a range
        const { sheet, address, bold, italic, fontSize, fontColor, fillColor, numberFormat, horizontalAlignment } = op;
        if (!address) throw new Error("formatRange: address required");
        const ws = sheet
          ? context.workbook.worksheets.getItem(sheet)
          : context.workbook.worksheets.getActiveWorksheet();
        const range = ws.getRange(address);
        if (bold !== undefined) range.format.font.bold = bold;
        if (italic !== undefined) range.format.font.italic = italic;
        if (fontSize !== undefined) range.format.font.size = fontSize;
        if (fontColor !== undefined) range.format.font.color = fontColor;
        if (fillColor !== undefined) range.format.fill.color = fillColor;
        if (numberFormat !== undefined) range.numberFormat = [[numberFormat]];
        if (horizontalAlignment !== undefined) range.format.horizontalAlignment = horizontalAlignment;
        await context.sync();
        return { ok: true, address };
      }

      case "getUsedRange": {
        // Get the used range of a worksheet
        const { sheet, valuesOnly = true } = op;
        const ws = sheet
          ? context.workbook.worksheets.getItem(sheet)
          : context.workbook.worksheets.getActiveWorksheet();
        const range = ws.getUsedRange();
        const loads = ["values", "address", "rowCount", "columnCount"];
        if (!valuesOnly) loads.push("formulas", "numberFormat");
        range.load(loads.join(","));
        await context.sync();
        const result = {
          address: range.address,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          values: range.values,
        };
        if (!valuesOnly) {
          result.formulas = range.formulas;
          result.numberFormat = range.numberFormat;
        }
        return result;
      }

      case "insertRow": {
        // Insert a row of values at a specific position
        const { sheet, row, values } = op;
        if (row === undefined) throw new Error("insertRow: row required (0-based)");
        if (!values) throw new Error("insertRow: values required (array)");
        const ws = sheet
          ? context.workbook.worksheets.getItem(sheet)
          : context.workbook.worksheets.getActiveWorksheet();
        const insertRange = ws.getRange(`${row + 1}:${row + 1}`);
        insertRange.insert(Excel.InsertShiftDirection.down);
        const cols = values.length;
        const addr = `A${row + 1}:${String.fromCharCode(64 + cols)}${row + 1}`;
        const target = ws.getRange(addr);
        target.values = [values];
        await context.sync();
        return { ok: true, row, address: addr };
      }

      case "deleteRow": {
        // Delete a row by index
        const { sheet, row } = op;
        if (row === undefined) throw new Error("deleteRow: row required (0-based)");
        const ws = sheet
          ? context.workbook.worksheets.getItem(sheet)
          : context.workbook.worksheets.getActiveWorksheet();
        const range = ws.getRange(`${row + 1}:${row + 1}`);
        range.delete(Excel.DeleteShiftDirection.up);
        await context.sync();
        return { ok: true, deleted: row };
      }

      default:
        throw new Error(`unknown op: ${op.kind}`);
    }
  });
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    log("not running inside Excel — host=", info.host);
    return;
  }
  log("Office ready, host=Excel");
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
