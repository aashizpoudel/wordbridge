#!/usr/bin/env node
// Rewrites addin/manifest.xml to use a new port and reinstalls it into
// Word's sideload folder. Usage: node tools/set-port.js 4000
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const MANIFEST = path.join(ROOT, "addin", "manifest.xml");
const WEF = path.join(os.homedir(), "Library", "Containers", "com.microsoft.Word", "Data", "Documents", "wef");
const INSTALLED = path.join(WEF, "wordbridge.manifest.xml");

const port = Number(process.argv[2]);
if (!Number.isInteger(port) || port <= 0 || port > 65535) {
  console.error(`usage: node tools/set-port.js <port>`);
  process.exit(2);
}

const xml = fs.readFileSync(MANIFEST, "utf8");
const patched = xml.replace(
  /(127\.0\.0\.1|localhost):\d+/g,
  (_m, host) => `${host}:${port}`,
);

fs.writeFileSync(MANIFEST, patched);
console.log(`[set-port] updated ${MANIFEST} -> port ${port}`);

if (!fs.existsSync(WEF)) {
  console.warn(`[set-port] Word wef folder not found at ${WEF}; skipping sideload install`);
} else {
  fs.copyFileSync(MANIFEST, INSTALLED);
  console.log(`[set-port] installed -> ${INSTALLED}`);
}

console.log(`\nNext steps:`);
console.log(`  1. Fully quit Word (Cmd+Q) so it re-scans wef/ on next launch.`);
console.log(`  2. Start the bridge on the new port:`);
console.log(`       node ~/wordbridge/server/server.js --port ${port}`);
console.log(`  3. Re-open your doc in Word, reopen the Word Bridge task pane.`);
console.log(`  4. From the CLI, use --port ${port} or export WORDBRIDGE_PORT=${port}.`);
