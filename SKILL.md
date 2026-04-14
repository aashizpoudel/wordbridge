---
name: wordbridge
description: Co-edit a Microsoft Word document that the user has open, using the Word Bridge HTTP server (default http://127.0.0.1:3001) to read structure and apply live, tracked-change edits via Office.js. TRIGGER when the user asks you to revise, rewrite, annotate, fix, or restructure a .docx they are working on in Word (e.g. "fix the abstract", "renumber the references", "tighten the methods section", "accept the edits I approved"). SKIP for plain-text files, markdown, or docs the user has not confirmed are currently open in Word.
---

# Word Bridge — editing a live Word document

You are editing a document the user has **open in Microsoft Word right now**. Every edit you apply is rendered live in Word and, if the user has Track Changes enabled, recorded as a tracked change attributed to the user's Word identity. Multiple assistants (Claude, Codex, other LLMs) may be driving the same document at the same time — so keep edits small, anchored, and reversible.

## 0. Is Word Bridge installed?

Before assuming you can edit, verify the bridge exists and can be reached. Do **not** run install/setup steps silently — always confirm with the user before touching their system.

```bash
# Is the HTTP bridge already running?
curl -sS http://127.0.0.1:3001/status 2>/dev/null
```

- **Response with `ok: true` and `clients >= 1`** → you're ready. Jump to §2.
- **Response with `ok: true` and `clients: 0`** → bridge is up but the Word task pane isn't connected. Ask the user to open Word → Insert → (My Add-ins | Shared Folder) → **Word Bridge**, wait for `Bridge: connected`, then proceed.
- **Connection refused / no response** → bridge is not running. Check whether the repo exists (`ls ~/wordbridge` or the project path the user mentioned). If it does, go to §1.2 to start it. If it doesn't, go to §1.1 to install.

## 1. Install and set up Word Bridge

Requires **Node 18+** and **Microsoft Word desktop (2016 or later)**.

### 1.1 First-time install

1. Clone the repo and install deps:
   ```bash
   git clone https://github.com/<user>/wordbridge ~/wordbridge
   cd ~/wordbridge && npm install
   ```
2. Sideload the Office Add-in manifest into Word. **This step is platform-specific**:

   **macOS:**
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   cp ~/wordbridge/addin/manifest.xml \
      ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/wordbridge.manifest.xml
   ```

   **Windows** (Word uses a *trusted shared folder catalog*, not a per-app `wef/` folder):
   1. Copy `addin\manifest.xml` into a folder on disk, e.g. `C:\wordbridge-manifests\`.
   2. In Word: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
   3. Paste the folder path into **Catalog Url**, click **Add catalog**, tick **Show in Menu**, click **OK**.

3. **Fully quit Word** so it re-scans the manifest on next launch:
   - macOS: ⌘Q (not just close the window)
   - Windows: close all Word windows, or `taskkill /f /im WINWORD.EXE` if stuck

4. Re-open Word and load the add-in:
   - macOS: **Insert → Add-ins → My Add-ins → Developer Add-ins → Word Bridge**
   - Windows: **Insert → My Add-ins → Shared Folder → Word Bridge**

   The task pane should show `Bridge: disconnected` at this point — expected, because the server isn't running yet.

### 1.2 Start the bridge server

```bash
# Default port 3001
node ~/wordbridge/server/server.js

# Custom port
node ~/wordbridge/server/server.js --port 4000
# or
WORDBRIDGE_PORT=4000 node ~/wordbridge/server/server.js
```

On Windows (PowerShell):
```powershell
node C:\path\to\wordbridge\server\server.js
$env:WORDBRIDGE_PORT=4000; node C:\path\to\wordbridge\server\server.js
```

Leave this running in a terminal. The task pane pill should switch to `Bridge: connected` within a second or two. Re-run `curl http://127.0.0.1:<port>/status` to confirm `clients >= 1`.

### 1.3 Enable tracked changes and set the author

- **Track Changes**: **Review → Track Changes → For Everyone** (macOS) or **Review → Track Changes** (Windows). The Office.js setter is broken on Mac, so the user must do this manually — *ask them to toggle it*, don't try to automate it.
- **Author identity** (every tracked edit is attributed to this name):
  - macOS: **Word → Preferences → User Information**
  - Windows: **File → Options → General → Personalize your copy of Microsoft Office**

### 1.4 Changing the port later

The task pane's URL is baked in by Word at launch time, so changing the port requires rewriting the manifest *and* restarting Word:

```bash
node ~/wordbridge/tools/set-port.js 4000
# macOS: auto-reinstalls to wef/
# Windows: manually re-copy addin\manifest.xml to your trusted-catalog folder
# then fully quit Word, restart it, and start the bridge on the new port
```

Once these steps are done and `GET /status` returns `clients >= 1`, the bridge is ready.

## 2. Before you touch the document

1. **Confirm the bridge is up and an add-in is connected.**
   ```bash
   curl -s http://127.0.0.1:3001/status
   # Or: node <repo>/cli/wordbridge.js status
   ```
   The response should show `ok: true` and `clients >= 1`. If `clients` is `0`, tell the user to open the Word Bridge task pane (Insert → My Add-ins → Word Bridge) and wait for `Bridge: connected` before retrying.

2. **Discover available tools** via `GET /tools` (or `GET /tools/<name>` for a single tool). Don't hard-code op schemas — the catalog is the source of truth, and new ops can appear. Each tool comes with `inputSchema`, `outputSchema`, and an example.
   ```bash
   curl -s http://127.0.0.1:3001/tools | jq '.tools[].name'
   ```

3. **Scout the document before editing.** You are flying blind otherwise. Good scouting ops:
   - `getText { limit }` — quick plain-text dump to understand the shape.
   - `getParagraphs { styleFilter?, limit?, textLimit? }` — structured list with styles and indices. Use this to find anchors and to count paragraphs before a `replaceParagraphByIndex`.
   - `getOoxml { anchor?, scope: "body"|"selection", charLimit? }` — raw OOXML when you need exact tag-level detail (fields, tables, tracked-change markup).

4. **Check Track Changes state with the user, not the API.** `getTrackChanges` / `setTrackChanges` are non-functional on Word for Mac (office-js #2797/#6246). If the user needs edits recorded as tracked changes, they must toggle **Review → Track Changes** in Word themselves. On Mac this cannot be automated.

## 3. How to send ops

Two equivalent paths:

- **HTTP** (most portable — works from any language/agent):
  ```
  POST http://127.0.0.1:3001/op
  Content-Type: application/json
  { "kind": "<tool name>", ...fields }
  ```
  Batch mode: `POST /ops` with `[{...}, {...}]` or `{ ops: [...] }`. Add `?stopOnError=1` to halt on first failure.

- **CLI** (wraps HTTP; convenient for shell workflows):
  ```bash
  node <repo>/cli/wordbridge.js <command> [args] [--port N]
  # e.g.
  node <repo>/cli/wordbridge.js find-replace "[1]" "(Cheng et al., 2022)" --max 1
  node <repo>/cli/wordbridge.js ops-file edits.json --stopOnError
  ```

On success you get `{ ok: true, result: <tool-specific object> }`. On failure: `{ ok: false, error, detail? }`. If the server returns 502 `"no add-in connected"`, the task pane was closed — ask the user to reopen it.

## 4. Editing ops — when to use each

| Intent | Op | Notes |
|---|---|---|
| Change exact text N times | `findReplace` | Exact-string match. Use `matchCase`/`matchWholeWord` to narrow. `maxReplacements: 0` means *replace all* — default to a small cap unless the user asked for all. |
| Delete a snippet | `deleteText` | Same anchoring rules as `findReplace`. |
| Insert after/before an anchor | `insertAfterText` | Set `asParagraph: true` + `style` to add a new styled paragraph (e.g. a reference entry). Otherwise it's an inline insert. |
| Replace an entire paragraph by index | `replaceParagraphByIndex` | Use after `getParagraphs` to get the index. Preserves paragraph style. Clears inline math/fields inside the paragraph — don't use on complex structures. |
| Re-style a paragraph | `setParagraphStyle` | Tracked as a format change. |
| Insert an image | `insertImage` | `base64` of PNG/JPEG + `anchor`. `location`: `"after"` / `"before"` (new paragraph next to anchor) or `"replace"` (inline, replaces the anchor text). Optional `widthPoints`, `alignment`. |
| Anything complex (tables, fields, literal `<w:ins>` with a spoofed author, `pPrChange`) | `insertOoxml` | Escape hatch. You must emit a well-formed OOXML fragment. This is the **only way to produce a tracked insertion when Track Changes is off**, via an explicit `<w:ins w:author="…">` wrapper. |

## 5. Reviewing tracked changes

- **Enumerate**: `getTrackedChanges { timeoutMs?, textLimit? }`. Known to hang on some Word-for-Mac builds; the server wraps it in a timeout. On hang, tell the user to review manually.
- **Accept / reject**: `reviewChanges { action: "accept"|"reject", selector, authorFilter?, maxMatches?, timeoutMs? }` with one of these selectors:
  - `{ kind: "all" }` — every change in body.
  - `{ kind: "index", value: N }` — Nth change (0-based).
  - `{ kind: "text", value: "...", matchCase?: bool }` — changes whose text contains substring.
  - `{ kind: "paragraph", anchor: "..." }` — changes inside the paragraph containing anchor.

  CLI shortcuts: `wordbridge accept --all`, `wordbridge accept --text "foo"`, `wordbridge reject --paragraph "Methods" --author someone`. On Mac, prefer `--paragraph` scope — full-body accept/reject can time out due to the same Office.js bug that affects enumeration.

## 6. Hard rules

- **Never guess anchors.** Anchor-based ops do exact-text search. If you haven't scouted the text via `getText` / `getParagraphs` / `getOoxml` in this session, scout first. An anchor that matches zero or multiple paragraphs will either error out or edit the wrong place.
- **Keep anchors short and unique.** Long anchors break on hidden characters, tracked-change markers, or soft hyphens. Aim for a distinctive 3–8 word fragment.
- **Prefer small, composable ops over large OOXML blobs.** `insertOoxml` is powerful and dangerous — one malformed tag can wedge the document. Use it only when no higher-level op fits.
- **Don't accept or reject tracked changes without being asked.** When the user asks you to revise, *leave the edits as pending tracked changes* so they can review. Only call `reviewChanges` when the user explicitly says to accept/reject something.
- **Don't toggle Track Changes for the user** — the API is broken on Mac. If they need it on, ask them to toggle it in Review → Track Changes.
- **Serialize ops, don't parallelize them.** Even if another assistant is on the same document, send one op at a time and wait for the result. Concurrent edits from two agents can race and corrupt the tracked-change log.
- **Batch with `/ops` when order matters.** If you have a series of edits that assume each other's results (e.g. insert a placeholder, then re-style it), send them as a single batch with `stopOnError=1` so a mid-batch failure doesn't leave the doc half-edited.
- **Watch for OneDrive-synced docs.** Don't recommend quitting Word mid-session on a OneDrive file — OneDrive merge can silently drop tracked-change markup.

## 7. When things go wrong

- `502 no add-in connected` — task pane is closed or WebSocket died. Ask the user to open Word Bridge from Insert → My Add-ins.
- `anchor not found: <str>` — your anchor didn't exact-match. Re-scout with `getParagraphs` (look at `text` previews) and pick a shorter, more unique snippet.
- `paragraph index out of range` — the doc changed since your last `getParagraphs` call. Re-scout before retrying.
- `reviewChanges timed out` — known Mac Word.js hang. Retry with `selector.kind: "paragraph"` to narrow scope, or tell the user to accept/reject manually.
- The bridge itself is 500-ing — `GET /status` and check the server logs. Don't retry blindly.

## 8. Reporting back to the user

When reporting what you did, be specific: which op(s) you ran, how many anchors matched, how many replacements/insertions landed, and — critically — whether the edits are recorded as tracked changes (you can infer this from whether the user had Track Changes on when they asked, but say so explicitly and tell them to check the Review ribbon if unsure). If you left any changes pending for their review, say that too.
