#!/usr/bin/env node
// Minimal CLI — POSTs ops to the bridge.
// Usage:
//   wordbridge status
//   wordbridge ping
//   wordbridge track on|off
//   wordbridge find-replace "old" "new" [--case] [--whole] [--max N]
//   wordbridge insert-after "anchor" "text" [--paragraph] [--style "Heading1"]
//   wordbridge delete "text" [--max N]
//   wordbridge set-style "anchor" "StyleName"
//   wordbridge get-text [--limit N]
//   wordbridge ops-file path/to/ops.json     # batch from JSON file
//   wordbridge raw '{"kind":"insertOoxml","anchor":"...","ooxml":"..."}'

function parseFlags(args) {
  const pos = [];
  const flags = {};
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a.startsWith("--")) {
      const eq = a.indexOf("=");
      if (eq >= 0) { flags[a.slice(2, eq)] = a.slice(eq + 1); continue; }
      const key = a.slice(2);
      const next = args[i + 1];
      if (next !== undefined && !next.startsWith("--")) { flags[key] = next; i++; }
      else flags[key] = true;
    } else pos.push(a);
  }
  return { pos, flags };
}

// Resolve --port globally (strip it from argv so subcommand parsing ignores it).
function resolveGlobalPort(argv) {
  const out = [];
  let port = null;
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--port" || a === "-p") { port = Number(argv[i + 1]); i++; }
    else if (a.startsWith("--port=")) { port = Number(a.slice("--port=".length)); }
    else out.push(a);
  }
  if (port === null && process.env.WORDBRIDGE_PORT) port = Number(process.env.WORDBRIDGE_PORT);
  if (port === null) port = 3001;
  return { port, rest: out };
}

const { port: PORT, rest: ARGV } = resolveGlobalPort(process.argv.slice(2));
const BASE = `http://127.0.0.1:${PORT}`;

async function post(pathname, body) {
  const res = await fetch(BASE + pathname, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const text = await res.text();
  try { return { status: res.status, json: JSON.parse(text) }; }
  catch { return { status: res.status, text }; }
}

async function get(pathname) {
  const res = await fetch(BASE + pathname);
  const text = await res.text();
  try { return { status: res.status, json: JSON.parse(text) }; }
  catch { return { status: res.status, text }; }
}

async function sendOp(op) {
  const r = await post("/op", op);
  return r;
}

async function main() {
  const [cmd, ...rest] = ARGV;
  const { pos, flags } = parseFlags(rest);

  try {
    switch (cmd) {
      case "status": {
        const r = await get("/status");
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "ping": {
        const r = await sendOp({ kind: "ping" });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "track": {
        const on = pos[0] === "on";
        const r = await sendOp({ kind: "setTrackChanges", on });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "find-replace": {
        const [find, replace] = pos;
        if (!find || replace === undefined) throw new Error("usage: wordbridge find-replace <find> <replace>");
        const r = await sendOp({
          kind: "findReplace",
          find, replace,
          matchCase: !!flags.case,
          matchWholeWord: !!flags.whole,
          maxReplacements: flags.max ? Number(flags.max) : 0,
        });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "insert-after": {
        const [anchor, text] = pos;
        if (!anchor || !text) throw new Error("usage: wordbridge insert-after <anchor> <text>");
        const r = await sendOp({
          kind: "insertAfterText",
          anchor, text,
          asParagraph: !!flags.paragraph,
          style: flags.style || undefined,
        });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "delete": {
        const [find] = pos;
        if (!find) throw new Error("usage: wordbridge delete <text>");
        const r = await sendOp({ kind: "deleteText", find, maxDeletions: flags.max ? Number(flags.max) : 0 });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "set-style": {
        const [anchor, style] = pos;
        if (!anchor || !style) throw new Error("usage: wordbridge set-style <anchor> <style>");
        const r = await sendOp({ kind: "setParagraphStyle", anchor, style });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "get-text": {
        const r = await sendOp({ kind: "getText", limit: flags.limit ? Number(flags.limit) : 4000 });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "get-paragraphs": {
        const r = await sendOp({
          kind: "getParagraphs",
          styleFilter: flags.style || undefined,
          limit: flags.limit ? Number(flags.limit) : 0,
          includeEmpty: !!flags.empty,
          textLimit: flags["text-limit"] ? Number(flags["text-limit"]) : 500,
        });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "get-ooxml": {
        const r = await sendOp({
          kind: "getOoxml",
          anchor: flags.anchor || undefined,
          scope: flags.scope || "body",
          charLimit: flags["char-limit"] ? Number(flags["char-limit"]) : 200000,
        });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "insert-image": {
        // Usage: wordbridge insert-image "anchor" path/to/img.png [--location after|before|replace] [--width 450] [--align center]
        const [anchor, imagePath] = pos;
        if (!anchor || !imagePath) throw new Error("usage: wordbridge insert-image <anchor> <imagePath> [--location after|before|replace] [--width POINTS] [--align left|center|right]");
        const fs = await import("node:fs/promises");
        const bytes = await fs.readFile(imagePath);
        const base64 = bytes.toString("base64");
        const r = await sendOp({
          kind: "insertImage",
          anchor,
          base64,
          location: flags.location || "after",
          widthPoints: flags.width ? Number(flags.width) : undefined,
          alignment: flags.align || "center",
        });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "replace-paragraph": {
        const [index, newText] = pos;
        if (index === undefined || newText === undefined) throw new Error("usage: wordbridge replace-paragraph <index> <newText>");
        const r = await sendOp({ kind: "replaceParagraphByIndex", index: Number(index), newText });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "accept":
      case "reject": {
        // Usage:
        //   wordbridge accept --all
        //   wordbridge accept --index 3
        //   wordbridge accept --text "Datasets"
        //   wordbridge accept --paragraph "Materials and Methods"
        //   wordbridge reject --text "foo" --author apoudel6 --max 1 --timeout 20000
        let selector;
        if (flags.all) selector = { kind: "all" };
        else if (flags.index !== undefined) selector = { kind: "index", value: Number(flags.index) };
        else if (flags.text) selector = { kind: "text", value: flags.text, matchCase: !!flags.case };
        else if (flags.paragraph) selector = { kind: "paragraph", anchor: flags.paragraph };
        else throw new Error(`usage: wordbridge ${cmd} <--all | --index N | --text STR | --paragraph ANCHOR> [--author X] [--max N] [--timeout MS]`);
        const r = await sendOp({
          kind: "reviewChanges",
          action: cmd,
          selector,
          authorFilter: flags.author || undefined,
          maxMatches: flags.max ? Number(flags.max) : 0,
          timeoutMs: flags.timeout ? Number(flags.timeout) : 15000,
        });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      // Legacy aliases
      case "accept-all":
      case "reject-all": {
        const r = await sendOp({
          kind: "reviewChanges",
          action: cmd === "accept-all" ? "accept" : "reject",
          selector: { kind: "all" },
          authorFilter: flags.author || undefined,
          timeoutMs: flags.timeout ? Number(flags.timeout) : 15000,
        });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "get-tracked-changes": {
        const r = await sendOp({
          kind: "getTrackedChanges",
          timeoutMs: flags.timeout ? Number(flags.timeout) : 3000,
          textLimit: flags["text-limit"] ? Number(flags["text-limit"]) : 200,
        });
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "raw": {
        const [json] = pos;
        if (!json) throw new Error("usage: wordbridge raw '<json op>'");
        const op = JSON.parse(json);
        const r = await sendOp(op);
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      case "ops-file": {
        const [p] = pos;
        if (!p) throw new Error("usage: wordbridge ops-file <path>");
        const fs = await import("node:fs/promises");
        const ops = JSON.parse(await fs.readFile(p, "utf8"));
        const r = await post("/ops" + (flags.stopOnError ? "?stopOnError=1" : ""), ops);
        console.log(JSON.stringify(r.json ?? r.text, null, 2));
        return;
      }
      default:
        console.error(`usage: wordbridge <status|ping|track|find-replace|insert-after|delete|set-style|get-text|get-paragraphs|get-ooxml|get-tracked-changes|raw|ops-file> ...`);
        process.exit(2);
    }
  } catch (err) {
    console.error("error:", err.message || String(err));
    process.exit(1);
  }
}

main();
