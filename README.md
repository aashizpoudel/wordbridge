# Word Bridge

Apply programmatic edits to a Microsoft Word document **while it's open**, rendered live and recorded as **tracked changes** attributed to the Word user.

## Use case — let Claude and Codex co-edit your Word documents

Word Bridge turns Word into a shared surface where **you and one or more AI assistants (Claude, Codex / GPT, or any tool-using LLM) can co-edit the same document at the same time**. You keep Word open and see every change land live in the page; the AI gets a stable, programmatic interface to the document without any copy-paste loop.

Concrete flows this unlocks:

- **Track-changes collaboration with Claude or Codex.** Ask the assistant to revise a paragraph, fix citations, tighten the abstract, or rewrite a section. Every edit shows up as a tracked insertion/deletion attributed to *you* (Word's user identity), so you can accept/reject each change from Word's Review ribbon exactly like a human collaborator's edits.
- **Co-pilot for long-form writing.** Draft in Word while Claude watches the structure via `getParagraphs` / `getOoxml` and suggests or applies edits to specific anchors you call out ("fix the methods section", "renumber references"). No context switching into a chat window.
- **LLM-to-LLM pipelines.** Because the bridge exposes its tool catalog at `GET /tools` in JSON Schema, Claude, Codex, and MCP clients can all talk to the same document simultaneously — one handling citations, another handling prose — without stomping each other, as long as you serialize their ops through the HTTP endpoint.
- **Scripted, auditable revisions.** Run a CLI or a `.json` batch of ops (`wordbridge ops-file edits.json`) to apply a reviewed set of changes in one pass, and use tracked changes as the audit log.
- **Review automation.** Let an assistant enumerate tracked changes (`getTrackedChanges`) and selectively `accept`/`reject` them by text, paragraph, or author — useful when you have a PR-style review loop on a doc.

The add-in and bridge are **platform-agnostic** to the caller: anything that can POST JSON to `http://127.0.0.1:3001/op` can drive Word.

### Pointing an LLM at this project

[`SKILL.md`](./SKILL.md) is a self-contained, skill-format instruction file that teaches any tool-using LLM (Claude Code, Codex, MCP clients, custom agents) how to install, run, and drive Word Bridge end-to-end. **Load it directly into the model's context** and the LLM gets:

- setup + install steps (clone, `npm install`, sideload the manifest, start the bridge),
- how to discover the op catalog at `GET /tools`,
- how to scout a document safely before editing,
- when to use each op (findReplace, insertAfterText, replaceParagraphByIndex, insertImage, insertOoxml, reviewChanges, …),
- hard rules to avoid corrupting tracked changes or stomping on other agents editing the same doc.

Ways to load it:

- **Claude Code**: drop `SKILL.md` into `~/.claude/skills/wordbridge/SKILL.md` (or the project-local `.claude/skills/` folder), and Claude will auto-trigger it when you ask it to edit a Word document.
- **Codex / GPT / MCP clients**: paste or include the contents of `SKILL.md` as a system/developer message, or have the agent fetch the raw file from the repo at session start.
- **Any agent**: `curl https://raw.githubusercontent.com/<your-fork>/wordbridge/main/SKILL.md` and inject the text into the model's prompt.

Once loaded, the LLM can call `GET http://127.0.0.1:3001/tools` at runtime for the live, version-accurate op schemas — the skill file tells it how, so it never needs to hard-code them.

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

Works on **macOS** and **Windows**. Requires Node 18+ and Microsoft Word (desktop, 2016 or later).

1. Install deps:
   ```bash
   cd wordbridge && npm install
   ```
2. Sideload the manifest into Word:

   **macOS** — drop the manifest into Word's developer folder:
   ```bash
   cp ~/wordbridge/addin/manifest.xml \
      ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/wordbridge.manifest.xml
   ```

   **Windows** — Word uses a *trusted shared folder catalog* instead of a per-app `wef/` folder:
   1. Create a folder anywhere, e.g. `C:\wordbridge-manifests\`, and copy `addin\manifest.xml` into it.
   2. In Word, **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
   3. Paste the folder path (e.g. `C:\wordbridge-manifests\`) into the **Catalog Url** box, click **Add catalog**, tick **Show in Menu**, click **OK**.
   4. Close and reopen Word.
3. **Fully quit Word before the next launch** — on macOS ⌘Q (not just close windows); on Windows right-click the taskbar icon → Close all windows, or `taskkill /f /im WINWORD.EXE` if it's stuck. Otherwise Word won't re-scan the manifest.

## Every session

1. Start the bridge:
   ```bash
   # macOS / Linux
   node ~/wordbridge/server/server.js
   node ~/wordbridge/server/server.js --port 4000
   WORDBRIDGE_PORT=4000 node ~/wordbridge/server/server.js
   ```
   ```powershell
   # Windows (PowerShell)
   node C:\path\to\wordbridge\server\server.js
   node C:\path\to\wordbridge\server\server.js --port 4000
   $env:WORDBRIDGE_PORT=4000; node C:\path\to\wordbridge\server\server.js
   ```
2. Open your `.docx` in Word.
3. Open the task pane:
   - **macOS**: **Insert → Add-ins → My Add-ins → Developer Add-ins → Word Bridge.**
   - **Windows**: **Insert → My Add-ins → Shared Folder → Word Bridge.**

   The task pane opens on the right.
4. The task pane shows `Bridge: connected` when the WebSocket is up.
5. **Enable Track Changes manually**: **Review → Track Changes → For Everyone** (Mac) / **Review → Track Changes** (Windows). The Office.js setter for `changeTrackingMode` is broken on Word for Mac (see *Known bugs*), so the add-in cannot toggle this for you on Mac. On Windows the underlying API works, but the add-in currently treats both platforms the same and expects you to toggle manually.
6. Set your author identity — every tracked change is attributed to this name:
   - **macOS**: **Word → Preferences → User Information**
   - **Windows**: **File → Options → General → Personalize your copy of Microsoft Office**

## Changing the port

The bridge, CLI, and task pane all accept a custom port. Changing the port that the **task pane is loaded from** requires updating the manifest and relaunching Word, because Word bakes the task-pane URL at launch time.

```bash
# 1. rewrite the manifest (and, on macOS, auto-reinstall it to wef/)
node ~/wordbridge/tools/set-port.js 4000
# On Windows: also manually copy addin\manifest.xml back into your
# trusted-catalog folder (e.g. C:\wordbridge-manifests\manifest.xml).

# 2. fully quit Word (⌘Q on Mac, close all Word windows on Windows)

# 3. start bridge on the new port
node ~/wordbridge/server/server.js --port 4000

# 4. reopen the doc, reopen the Word Bridge task pane
# 5. CLI now needs --port 4000 or WORDBRIDGE_PORT=4000
```

Note: `tools/set-port.js` currently auto-installs to Word's macOS `wef/` folder only. On Windows it still rewrites `addin/manifest.xml`, but you need to copy the updated manifest into your trusted-catalog folder yourself.

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

Put `~/wordbridge/cli` on your `PATH` (or symlink `wordbridge.js` into `/usr/local/bin/wordbridge`) to drop the `node …` prefix. On Windows, add `C:\path\to\wordbridge\cli` to `PATH` and invoke it as `node wordbridge.js …`, or create a `wordbridge.cmd` wrapper.

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

`Word.Document.changeTrackingMode` is spec'd in `WordApi 1.4` but is non-functional on Word for Mac as of February 2026. Windows is partially fixed:

- [office-js#2797](https://github.com/OfficeDev/office-js/issues/2797) — `PropertyNotLoaded` on read; **fixed on Windows**, still broken on Mac.
- [office-js#6246](https://github.com/officedev/office-js/issues/6246) — cross-platform "retrieve the markup revision mode" gap, open.
- [office-js#6514](https://github.com/OfficeDev/office-js/issues/6514) — `getTrackedChanges()` hangs silently on Word for Mac 16.106, open. Does not reproduce on Windows as of this writing.

The add-in catches these gracefully and degrades to manual-mode on both platforms.

## Troubleshooting

- **Task pane doesn't list the add-in.** You didn't fully quit Word before launching after the manifest was installed. ⌘Q (Mac) / close all Word windows (Windows), then reopen. On Windows, also confirm the trusted catalog folder is registered in **File → Options → Trust Center → Trusted Add-in Catalogs** with **Show in Menu** ticked.
- **Task pane is blank.** Word's webview refused `http://127.0.0.1`. Switch to HTTPS with a self-signed cert (not shipped yet in this scaffold). Windows is stricter than Mac here — loopback HTTP may be blocked entirely depending on Edge WebView2 policy.
- **`no add-in connected`.** The task pane isn't open, or its WebSocket disconnected (check the pill in the task pane).
- **Author on tracked changes is wrong.** The add-in uses whoever is signed into Word. Set your name in **Word → Preferences → User Information** (Mac) or **File → Options → General** (Windows) *before* opening the document.
- **Edits don't appear as tracked.** Check that the Review ribbon shows Track Changes toggled on for this document.
- **Anchor not found.** `insertAfterText` / `setParagraphStyle` use exact string search. Shorten the anchor to a unique snippet.
- **`get-tracked-changes` times out.** Known Mac Word.js bug (#6514). Increase `--timeout` or accept that you can't enumerate changes on this build. Should work on Windows.
- **Bridge port in use.**
  - macOS / Linux: `lsof -ti :3001 | xargs kill`
  - Windows (PowerShell): `Get-NetTCPConnection -LocalPort 3001 | Select-Object -ExpandProperty OwningProcess | ForEach-Object { Stop-Process -Id $_ -Force }`

  Or start the bridge on a different port with `--port`.

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
