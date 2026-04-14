// Tool catalog for the Word Bridge /op endpoint.
// Exposed via GET /tools so LLM callers can self-discover how to use the bridge.
// Schema style: JSON Schema (subset) — understood by Claude, OpenAI, and MCP clients.

export const bridgeInfo = {
  name: "wordbridge",
  version: "0.1.0",
  description:
    "Apply programmatic edits to a Microsoft Word document that the user has open in Word. " +
    "Edits are applied live through an Office.js task-pane add-in and are recorded as " +
    "tracked changes when Track Changes is enabled in Word's Review ribbon. " +
    "Use GET /tools to discover available operations.",
  requestEndpoint: {
    method: "POST",
    path: "/op",
    contentType: "application/json",
    body: {
      kind: "<tool name from the tools list>",
      "<field>": "<value — fields vary by tool, see tool.inputSchema>",
    },
    description:
      "Send one op at a time. The op is forwarded to the Word task-pane add-in over a WebSocket " +
      "and executed via Word.js. If no add-in is connected, the server returns 502 with " +
      "'no add-in connected'.",
  },
  batchEndpoint: {
    method: "POST",
    path: "/ops",
    contentType: "application/json",
    body: "array of op objects, OR { ops: [...] }",
    query: { stopOnError: "1 to halt the batch on first failure" },
    description: "Apply an ordered batch of ops. Each result is returned in order.",
  },
  responseFormat: {
    success: {
      type: "result",
      id: "<uuid>",
      ok: true,
      result: "<tool-specific result object, see tool.outputSchema>",
    },
    failure: {
      type: "result",
      id: "<uuid>",
      ok: false,
      error: "<string>",
      detail: "<optional Office.js error detail JSON>",
    },
  },
  statusEndpoint: { method: "GET", path: "/status", description: "Reports bridge health and number of connected add-in clients." },
  toolsEndpoint: { method: "GET", path: "/tools", description: "Returns this catalog. Also accepts GET /tools/<name> for a single tool." },
  notes: [
    "Every op runs against whichever Word document is currently focused in Word. Open only the target document to avoid editing the wrong file.",
    "Text-mutating ops are recorded as tracked changes if the user has Track Changes enabled in the Review ribbon. Mac Word's Office.js API cannot set this state programmatically — the user must toggle it manually.",
    "The tracked-change author is the user's Word identity (Word → Preferences → User Information). The insertOoxml op is the only way to override this, by emitting a literal <w:ins w:author='...'> fragment.",
    "Anchor-based ops (insertAfterText, setParagraphStyle, deleteText, insertOoxml, findReplace) do exact-text search. Keep anchors short and unique.",
    "Before editing, use getText / getParagraphs / getOoxml to scout the document. Never guess anchors blindly.",
  ],
};

export const tools = [
  {
    name: "ping",
    description: "Round-trip sanity check. Returns { pong: true }.",
    inputSchema: {
      type: "object",
      properties: { kind: { const: "ping" } },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { pong: { type: "boolean" } },
      required: ["pong"],
    },
    example: {
      request: { kind: "ping" },
      response: { ok: true, result: { pong: true } },
    },
  },

  {
    name: "getText",
    description:
      "Read the document body as plain text. Strips styles, tables, images, headers/footers, and tracked-change markers. " +
      "Best for a quick scout; use getParagraphs for structure or getOoxml for raw XML.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "getText" },
        limit: { type: "integer", minimum: 1, default: 4000, description: "Max characters to return." },
      },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: {
        text: { type: "string" },
        totalLength: { type: "integer", description: "Full body length before truncation." },
      },
    },
    example: {
      request: { kind: "getText", limit: 2000 },
      response: { ok: true, result: { text: "Paper number and page range\r...", totalLength: 24725 } },
    },
  },

  {
    name: "getParagraphs",
    description:
      "Return a structured list of paragraphs with their style, text preview, and length. " +
      "Filter by style name to isolate a section (e.g. all RefListing entries).",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "getParagraphs" },
        styleFilter: { type: "string", description: "Return only paragraphs with this style (exact match on style or styleBuiltIn)." },
        limit: { type: "integer", minimum: 0, default: 0, description: "Cap on paragraphs returned. 0 = no cap." },
        includeEmpty: { type: "boolean", default: false, description: "Include blank paragraphs." },
        textLimit: { type: "integer", minimum: 0, default: 500, description: "Per-paragraph text snippet size; longer paragraphs are truncated with '…'." },
      },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: {
        total: { type: "integer", description: "Total paragraphs in the body." },
        returned: { type: "integer" },
        paragraphs: {
          type: "array",
          items: {
            type: "object",
            properties: {
              index: { type: "integer" },
              style: { type: "string" },
              styleBuiltIn: { type: "string" },
              text: { type: "string" },
              length: { type: "integer" },
            },
          },
        },
      },
    },
    example: {
      request: { kind: "getParagraphs", styleFilter: "RefListing", textLimit: 120 },
      response: {
        ok: true,
        result: {
          total: 180,
          returned: 11,
          paragraphs: [
            { index: 162, style: "RefListing", styleBuiltIn: "", text: "[1] Cheng et al. …", length: 210 },
          ],
        },
      },
    },
  },

  {
    name: "getOoxml",
    description:
      "Return the raw OOXML for a range. Use anchor to target a specific paragraph, " +
      "scope='body' for the whole body, scope='selection' for the user's current selection. " +
      "Output is truncated at charLimit.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "getOoxml" },
        anchor: { type: "string", description: "Exact-text anchor; OOXML returned is for the paragraph containing this string." },
        scope: { type: "string", enum: ["body", "selection"], default: "body" },
        charLimit: { type: "integer", minimum: 1, default: 200000 },
      },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: {
        ooxml: { type: "string" },
        totalLength: { type: "integer" },
        truncated: { type: "boolean" },
      },
    },
    example: {
      request: { kind: "getOoxml", anchor: "Abstract" },
      response: { ok: true, result: { ooxml: "<pkg:package …>", totalLength: 12344, truncated: false } },
    },
  },

  {
    name: "getTrackedChanges",
    description:
      "Enumerate pending tracked changes (insertions, deletions, formats) with author, date, type, and text. " +
      "Known issue: this call hangs on some Word-for-Mac builds (office-js #5535, #6514); the server " +
      "wraps it in a timeout and returns an error if it hangs.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "getTrackedChanges" },
        timeoutMs: { type: "integer", minimum: 500, default: 3000 },
        textLimit: { type: "integer", minimum: 0, default: 200 },
      },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: {
        count: { type: "integer" },
        changes: {
          type: "array",
          items: {
            type: "object",
            properties: {
              index: { type: "integer" },
              type: { type: "string" },
              author: { type: "string" },
              date: { type: "string" },
              text: { type: "string" },
              length: { type: "integer" },
            },
          },
        },
      },
    },
    example: {
      request: { kind: "getTrackedChanges", timeoutMs: 5000 },
      response: { ok: true, result: { count: 2, changes: [] } },
    },
  },

  {
    name: "findReplace",
    description:
      "Exact-text search and replace over the body. Recorded as a tracked change if Track Changes is on.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "findReplace" },
        find: { type: "string" },
        replace: { type: "string" },
        matchCase: { type: "boolean", default: false },
        matchWholeWord: { type: "boolean", default: false },
        maxReplacements: { type: "integer", minimum: 0, default: 0, description: "0 = replace all." },
      },
      required: ["kind", "find", "replace"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { matched: { type: "integer" }, replaced: { type: "integer" } },
    },
    example: {
      request: { kind: "findReplace", find: "[1]", replace: "(Cheng et al., 2022)", maxReplacements: 1 },
      response: { ok: true, result: { matched: 3, replaced: 1 } },
    },
  },

  {
    name: "insertAfterText",
    description:
      "Find an anchor string and insert new text after it. If asParagraph=true, inserts a new paragraph " +
      "(optionally with a named style) after the paragraph containing the anchor.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "insertAfterText" },
        anchor: { type: "string" },
        text: { type: "string" },
        asParagraph: { type: "boolean", default: false },
        style: { type: "string", description: "Applied only if asParagraph=true." },
      },
      required: ["kind", "anchor", "text"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { anchorMatches: { type: "integer" } },
    },
    example: {
      request: { kind: "insertAfterText", anchor: "References", text: "New entry.", asParagraph: true, style: "RefListing" },
      response: { ok: true, result: { anchorMatches: 1 } },
    },
  },

  {
    name: "deleteText",
    description: "Exact-text delete. Recorded as a tracked deletion if Track Changes is on.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "deleteText" },
        find: { type: "string" },
        maxDeletions: { type: "integer", minimum: 0, default: 0, description: "0 = delete all matches." },
      },
      required: ["kind", "find"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { matched: { type: "integer" }, deleted: { type: "integer" } },
    },
    example: {
      request: { kind: "deleteText", find: "[placeholder]", maxDeletions: 1 },
      response: { ok: true, result: { matched: 2, deleted: 1 } },
    },
  },

  {
    name: "setParagraphStyle",
    description: "Apply a paragraph style to the paragraph containing an anchor string. Recorded as a tracked format change.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "setParagraphStyle" },
        anchor: { type: "string" },
        style: { type: "string", description: "Style name, e.g. 'Heading1', 'RefListing', 'RefTitle'." },
      },
      required: ["kind", "anchor", "style"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { ok: { type: "boolean" } },
    },
    example: {
      request: { kind: "setParagraphStyle", anchor: "References", style: "RefTitle" },
      response: { ok: true, result: { ok: true } },
    },
  },

  {
    name: "insertOoxml",
    description:
      "Escape hatch: inject a raw OOXML fragment at an anchor. Use this when no higher-level op fits " +
      "(complex tables, pPrChange, field codes, or a literal <w:ins w:author='...'> tracked insertion " +
      "that bypasses the Review-ribbon toggle entirely). The fragment must be a well-formed OOXML snippet.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "insertOoxml" },
        anchor: { type: "string" },
        ooxml: { type: "string", description: "Well-formed OOXML fragment, e.g. '<w:p xmlns:w=\"…\"><w:r><w:t>hi</w:t></w:r></w:p>'." },
        location: { type: "string", enum: ["before", "after", "replace"], default: "after" },
      },
      required: ["kind", "anchor", "ooxml"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { ok: { type: "boolean" } },
    },
    example: {
      request: {
        kind: "insertOoxml",
        anchor: "Conclusion",
        location: "after",
        ooxml: '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:ins w:id="1" w:author="apoudel6" w:date="2026-04-14T00:00:00Z"><w:r><w:t>Thanks to the reviewers.</w:t></w:r></w:ins></w:p>',
      },
      response: { ok: true, result: { ok: true } },
    },
  },

  {
    name: "setTrackChanges",
    description:
      "Attempt to toggle Word's Track Changes state. BROKEN on Word for Mac (office-js #2797, #6246) — " +
      "the op does not throw, but it has no effect. Users on Mac must toggle manually via Review → Track Changes.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "setTrackChanges" },
        on: { type: "boolean" },
      },
      required: ["kind", "on"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { mode: { type: ["string", "null"] }, on: { type: ["boolean", "null"] }, manual: { type: "boolean" } },
    },
    example: {
      request: { kind: "setTrackChanges", on: true },
      response: { ok: true, result: { mode: null, on: null, manual: true } },
    },
  },

  {
    name: "getTrackChanges",
    description:
      "Report the current Track Changes state. BROKEN on Word for Mac — returns { manual: true } and " +
      "expects the caller to verify state via the Review ribbon.",
    inputSchema: {
      type: "object",
      properties: { kind: { const: "getTrackChanges" } },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { mode: { type: ["string", "null"] }, on: { type: ["boolean", "null"] }, manual: { type: "boolean" } },
    },
    example: {
      request: { kind: "getTrackChanges" },
      response: { ok: true, result: { mode: null, on: null, manual: true } },
    },
  },
];

export function getToolByName(name) {
  return tools.find((t) => t.name === name) || null;
}
