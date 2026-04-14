# Word Bridge

Apply programmatic edits to a Microsoft Word document **while it's open**, rendered live and recorded as **tracked changes** attributed to the Word user.

## Architecture

```
  CLI / Claude Code / any script / MCP client
            │  POST /op   (JSON over HTTP)
            ▼
   Bridge server (Node, 127.0.0.1:<port>)
            │  WebSocket push
            ▼
   Word task-pane add-in (Office.js / Word.js)
            │  live edits
            ▼
   The document you have open in Word
```

- **server/** — Node/Express + `ws` bridge. Default port `3001`, overridable.
- **addin/** — Office Add-in (task pane) with a WebSocket client. Served by the bridge at `/addin/taskpane.html`. Derives its WS URL from `window.location`, so it adapts to whichever port served it.
- **cli/wordbridge.js** — thin HTTP client for the bridge.
- **tools/set-port.js** — helper that rewrites the manifest to a new port and reinstalls it to Word's sideload folder.

## One-time setup

1. Install deps:
   ```bash
   cd ~/wordbridge && npm install
   ```
2. Sideload the manifest into Word's developer folder:
   ```bash
   cp ~/wordbridge/addin/manifest.xml \
      ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/wordbridge.manifest.xml
   ```
3. **Fully quit Word** (⌘Q — not just close windows) before the next launch, otherwise Word won't pick up the new manifest.

## Every session

1. Start the bridge:
   ```bash
   node ~/wordbridge/server/server.js
   # or on a different port:
   node ~/wordbridge/server/server.js --port 4000
   # or via env var:
   WORDBRIDGE_PORT=4000 node ~/wordbridge/server/server.js
   ```
2. Open your `.docx` in Word.
3. **Insert → Add-ins → My Add-ins → Developer Add-ins → Word Bridge.** The task pane opens on the right.
4. The task pane shows `Bridge: connected` when the WebSocket is up.
5. **Enable Track Changes manually**: **Review → Track Changes → For Everyone**. The Office.js setter for `changeTrackingMode` is broken on Word for Mac (see *Known bugs*), so the add-in cannot toggle this for you.
6. Set your author identity: **Word → Preferences → User Information**. Every tracked change will be attributed to this name.

## Changing the port

The bridge, CLI, and task pane all accept a custom port. Changing the port that the **task pane is loaded from** requires updating the manifest and relaunching Word, because Word bakes the task-pane URL at launch time.

```bash
# 1. rewrite manifest + reinstall to wef/
node ~/wordbridge/tools/set-port.js 4000

# 2. fully quit Word (⌘Q)

# 3. start bridge on the new port
node ~/wordbridge/server/server.js --port 4000

# 4. reopen the doc, reopen the Word Bridge task pane
# 5. CLI now needs --port 4000 or WORDBRIDGE_PORT=4000
```

Once the task pane is loaded, it reads its own `window.location.host` to build the WebSocket URL, so it always connects to the bridge it was served from. No client-side config.

## Using the CLI

```bash
# is the bridge running?
node ~/wordbridge/cli/wordbridge.js status

# round-trip (requires the task pane to be open)
node ~/wordbridge/cli/wordbridge.js ping

# try to toggle track changes (no-op on Mac Word — use Review ribbon instead)
node ~/wordbridge/cli/wordbridge.js track on

# find & replace (tracked when Track Changes is on in Word)
node ~/wordbridge/cli/wordbridge.js find-replace "[1]" "(Cheng et al., 2022)"
node ~/wordbridge/cli/wordbridge.js find-replace "old phrase" "new phrase" --case --whole --max 1

# insert text after an anchor
node ~/wordbridge/cli/wordbridge.js insert-after "Abstract." " This paper…"

# insert a new paragraph after a heading, with a style
node ~/wordbridge/cli/wordbridge.js insert-after "References" "New ref entry." --paragraph --style "RefListing"

# delete text
node ~/wordbridge/cli/wordbridge.js delete "stray sentence." --max 1

# set paragraph style by anchor text
node ~/wordbridge/cli/wordbridge.js set-style "References" "RefTitle"

# read body text (plain, lossy)
node ~/wordbridge/cli/wordbridge.js get-text --limit 8000

# read structured paragraph list (index, style, text preview, length)
node ~/wordbridge/cli/wordbridge.js get-paragraphs --style RefListing --text-limit 150
node ~/wordbridge/cli/wordbridge.js get-paragraphs --limit 50 --empty

# read raw OOXML for a range
node ~/wordbridge/cli/wordbridge.js get-ooxml --anchor "Abstract"
node ~/wordbridge/cli/wordbridge.js get-ooxml --scope body --char-limit 50000
node ~/wordbridge/cli/wordbridge.js get-ooxml --scope selection

# list pending tracked changes (may time out on Mac — known Word.js bug)
node ~/wordbridge/cli/wordbridge.js get-tracked-changes --timeout 5000

# use a non-default port
node ~/wordbridge/cli/wordbridge.js --port 4000 status
WORDBRIDGE_PORT=4000 node ~/wordbridge/cli/wordbridge.js ping

# batch ops from a JSON file (array of op objects)
node ~/wordbridge/cli/wordbridge.js ops-file edits.json --stopOnError

# raw op (escape hatch for anything the CLI doesn't wrap)
node ~/wordbridge/cli/wordbridge.js raw '{"kind":"insertOoxml","anchor":"Conclusion","ooxml":"<w:p .../>","location":"after"}'
```

Put `~/wordbridge/cli` on your `PATH` (or symlink `wordbridge.js` into `/usr/local/bin/wordbridge`) to drop the `node …` prefix.

## Supported ops

| kind | fields | notes |
|---|---|---|
| `ping` | — | round-trip check |
| `setTrackChanges` | `on: bool` | **broken on Mac Word** — use Review ribbon; op is a best-effort no-op |
| `getTrackChanges` | — | **broken on Mac Word** — returns `manual` if unreadable |
| `findReplace` | `find`, `replace`, `matchCase?`, `matchWholeWord?`, `maxReplacements?` | exact-text search |
| `insertAfterText` | `anchor`, `text`, `asParagraph?`, `style?` | searches for `anchor`, inserts after |
| `deleteText` | `find`, `maxDeletions?` | |
| `setParagraphStyle` | `anchor`, `style` | |
| `insertOoxml` | `anchor`, `ooxml`, `location: "before"\|"after"\|"replace"` | escape hatch for complex XML, including `<w:ins>` |
| `getText` | `limit?` | plain body text (no structure) |
| `getParagraphs` | `styleFilter?`, `limit?`, `includeEmpty?`, `textLimit?` | structured paragraph list with style + index |
| `getOoxml` | `anchor?`, `scope? = body\|selection`, `charLimit?` | raw OOXML for a range |
| `getTrackedChanges` | `timeoutMs? = 3000`, `textLimit?` | pending tracked changes (timed out on some Mac builds) |

Text edits are recorded as tracked changes **if** Track Changes is enabled in Word's Review ribbon. The add-in cannot force this state on Mac Word (Office.js bug). The `insertOoxml` op is the escape hatch that can produce tracked insertions even when Track Changes is off, by emitting a literal `<w:ins>` fragment.

## Known bugs (upstream, not this project)

`Word.Document.changeTrackingMode` is spec'd in `WordApi 1.4` but is non-functional on Word for Mac as of February 2026:

- [office-js#2797](https://github.com/OfficeDev/office-js/issues/2797) — `PropertyNotLoaded` on read, fixed on Windows, not Mac.
- [office-js#6246](https://github.com/officedev/office-js/issues/6246) — cross-platform "retrieve the markup revision mode" gap, open.
- [office-js#6514](https://github.com/OfficeDev/office-js/issues/6514) — `getTrackedChanges()` hangs silently on Word for Mac 16.106, open.

The add-in catches these gracefully and degrades to manual-mode.

## Troubleshooting

- **Task pane doesn't list the add-in.** You didn't fully quit Word before launching after the manifest was installed. ⌘Q, then reopen.
- **Task pane is blank.** Word's webview refused `http://127.0.0.1`. Switch to HTTPS with a self-signed cert (not shipped yet in this scaffold).
- **`no add-in connected`.** The task pane isn't open, or its WebSocket disconnected (check the pill in the task pane).
- **Author on tracked changes is wrong.** The add-in uses whoever is signed into Word. Set your name in **Word → Preferences → User Information** *before* opening the document.
- **Edits don't appear as tracked.** Check that the Review ribbon shows Track Changes toggled on for this document.
- **Anchor not found.** `insertAfterText` / `setParagraphStyle` use exact string search. Shorten the anchor to a unique snippet.
- **`get-tracked-changes` times out.** Known Mac Word.js bug (#6514). Increase `--timeout` or accept that you can't enumerate changes on this build.
- **Bridge port in use.** Another process is on `3001`. Either kill it (`lsof -ti :3001 | xargs kill`) or start the bridge on a different port with `--port`.

## What this can't do (limitations)

- The add-in must be loaded — if Word is closed, ops fail with "no add-in connected". No file-edit fallback; edit the `.docx` directly (e.g., raw OOXML) when Word is closed.
- Office.js doesn't let you set an arbitrary tracked-change author via the normal path — the author is whoever is signed into Word. `insertOoxml` with a literal `<w:ins w:author="…">` fragment is the only way to override it.
- Very exotic OOXML (`pPrChange` with complex children, custom field codes, author-index tables) needs the `insertOoxml` escape hatch.
- Accepting or rejecting changes in Word while an op is mid-flight can race. Let ops finish before reviewing.
- If you edit a OneDrive-synced `.docx`, don't quit Word and let OneDrive merge mid-session — that can silently drop tracked-change markup.

## Directory layout

```
wordbridge/
├── package.json
├── README.md
├── addin/
│   ├── manifest.xml         # Office Add-in manifest (sideloaded into Word's wef/)
│   ├── taskpane.html        # task pane UI
│   └── taskpane.js          # Office.js + WebSocket client
├── server/
│   └── server.js            # Express + ws bridge
├── cli/
│   └── wordbridge.js        # CLI client
├── tools/
│   └── set-port.js          # manifest port rewriter
└── test/                    # throwaway .docx copies for testing
```
