// Tool catalog for the Office Bridge /op endpoint.
// Exposed via GET /tools so LLM callers can self-discover how to use the bridge.
// Schema style: JSON Schema (subset) — understood by Claude, OpenAI, and MCP clients.
//
// Every tool has a `host` field identifying which Office app it runs against:
//   "word"       — routed to a connected Word task pane (Word.js)
//   "powerpoint" — routed to a connected PowerPoint task pane (PowerPoint.js)
//   "any"        — implemented by every task pane (e.g. ping)
// Callers can filter via GET /tools?host=<name>, or pass `target: "<host>"`
// on /op to force routing when multiple hosts are connected.

export const bridgeInfo = {
  name: "officebridge",
  version: "0.2.0",
  description:
    "Apply programmatic edits to a Microsoft Office document that the user has open. " +
    "Word and PowerPoint are supported — each tool is tagged with a `host` field. " +
    "Edits are applied live through an Office.js task-pane add-in over a WebSocket. " +
    "Word text-mutating ops are recorded as tracked changes when Track Changes is on. " +
    "Use GET /tools to discover available operations (optionally ?host=word|powerpoint).",
  requestEndpoint: {
    method: "POST",
    path: "/op",
    contentType: "application/json",
    body: {
      kind: "<tool name from the tools list>",
      target: "<optional: 'word' or 'powerpoint' to force routing>",
      clientId: "<optional: route to a specific connected client>",
      "<field>": "<value — fields vary by tool, see tool.inputSchema>",
    },
    description:
      "Send one op at a time. The op is forwarded over a WebSocket to the matching task-pane " +
      "add-in and executed via Office.js. If no add-in of the right host is connected, the server " +
      "returns 502 with 'no <host> add-in connected'.",
  },
  batchEndpoint: {
    method: "POST",
    path: "/ops",
    contentType: "application/json",
    body: "array of op objects, OR { ops: [...], target?, clientId? }",
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
  statusEndpoint: { method: "GET", path: "/status", description: "Reports bridge health and connected add-in clients (by kind)." },
  toolsEndpoint: { method: "GET", path: "/tools", description: "Returns this catalog. GET /tools?host=<name> filters by host. GET /tools/<name> returns a single tool." },
  notes: [
    "Each op runs against whichever document is focused in its host app. Open only the target document to avoid editing the wrong file.",
    "Word: text-mutating ops are recorded as tracked changes if Track Changes is on in the Review ribbon. Mac Word's Office.js API cannot set this programmatically — toggle manually.",
    "Word: the tracked-change author is the user's Word identity (Word → Preferences → User Information). Use insertOoxml with a <w:ins w:author='...'> fragment to override.",
    "Word anchor-based ops (insertAfterText, setParagraphStyle, deleteText, insertOoxml, findReplace) do exact-text search. Keep anchors short and unique.",
    "PowerPoint: ops are indexed by slide position (0-based). Use getOutline or getSlides first to scout.",
    "PowerPoint: shapes are addressed by shapeId (preferred, stable) or shapeName. Use getShapes to list them.",
    "Before editing, scout with getText/getParagraphs/getOoxml (Word) or getOutline/getSlideText/getShapes (PowerPoint).",
  ],
};

export const tools = [
  // ---------------- shared ----------------
  {
    name: "ping",
    host: "any",
    description: "Round-trip sanity check. Returns { pong: true }. Implemented by every task pane.",
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

  // ---------------- Word ----------------
  {
    name: "getText",
    host: "word",
    description:
      "Read the Word document body as plain text. Strips styles, tables, images, headers/footers, and tracked-change markers. " +
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
    host: "word",
    description:
      "Return a structured list of Word paragraphs with their style, text preview, and length. " +
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
    host: "word",
    description:
      "Return the raw OOXML for a Word range. Use anchor to target a specific paragraph, " +
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
    host: "word",
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
    host: "word",
    description:
      "Exact-text search and replace over the Word body. Recorded as a tracked change if Track Changes is on.",
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
    host: "word",
    description:
      "Find an anchor string in Word and insert new text after it. If asParagraph=true, inserts a new paragraph " +
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
    host: "word",
    description: "Exact-text delete in Word. Recorded as a tracked deletion if Track Changes is on.",
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
    host: "word",
    description: "Apply a paragraph style in Word to the paragraph containing an anchor string. Recorded as a tracked format change.",
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
    host: "word",
    description:
      "Escape hatch for Word: inject a raw OOXML fragment at an anchor. Use this when no higher-level op fits " +
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
    host: "word",
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
    host: "word",
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

  // ---------------- PowerPoint ----------------
  {
    name: "getSlides",
    host: "powerpoint",
    description:
      "List slides in the active PowerPoint presentation (id + zero-based index). Use as a cheap scout before any slide-targeted op.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "getSlides" },
        limit: { type: "integer", minimum: 0, default: 0, description: "Cap on slides returned. 0 = all." },
      },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: {
        total: { type: "integer" },
        returned: { type: "integer" },
        slides: {
          type: "array",
          items: {
            type: "object",
            properties: { id: { type: "string" }, index: { type: "integer" } },
          },
        },
      },
    },
    example: {
      request: { kind: "getSlides" },
      response: { ok: true, result: { total: 3, returned: 3, slides: [{ id: "256", index: 0 }, { id: "257", index: 1 }, { id: "258", index: 2 }] } },
    },
  },

  {
    name: "getOutline",
    host: "powerpoint",
    description:
      "Return the title of every slide in order — the presentation's outline. Titles are detected by shape name containing 'title'. Null when a slide has no title placeholder or its title is empty.",
    inputSchema: {
      type: "object",
      properties: { kind: { const: "getOutline" } },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: {
        slideCount: { type: "integer" },
        outline: {
          type: "array",
          items: {
            type: "object",
            properties: { index: { type: "integer" }, title: { type: ["string", "null"] } },
          },
        },
      },
    },
    example: {
      request: { kind: "getOutline" },
      response: { ok: true, result: { slideCount: 2, outline: [{ index: 0, title: "Introduction" }, { index: 1, title: "Methods" }] } },
    },
  },

  {
    name: "getSlideText",
    host: "powerpoint",
    description: "Dump all text from every shape on one slide. Shapes without a text frame are skipped.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "getSlideText" },
        slideIndex: { type: "integer", minimum: 0, default: 0, description: "Zero-based slide index." },
        textLimit: { type: "integer", minimum: 1, default: 4000, description: "Per-shape text cap." },
      },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: {
        slideIndex: { type: "integer" },
        shapeCount: { type: "integer" },
        texts: {
          type: "array",
          items: {
            type: "object",
            properties: {
              shapeId: { type: "string" },
              shapeName: { type: "string" },
              text: { type: "string" },
            },
          },
        },
      },
    },
    example: {
      request: { kind: "getSlideText", slideIndex: 0 },
      response: { ok: true, result: { slideIndex: 0, shapeCount: 2, texts: [{ shapeId: "4", shapeName: "Title 1", text: "Introduction" }, { shapeId: "5", shapeName: "Content 2", text: "- bullet one\n- bullet two" }] } },
    },
  },

  {
    name: "getShapes",
    host: "powerpoint",
    description: "List every shape on a slide with its id, name, type, and geometry (points). Use shapeId to target a shape precisely in other ops.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "getShapes" },
        slideIndex: { type: "integer", minimum: 0, default: 0 },
      },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: {
        slideIndex: { type: "integer" },
        shapes: {
          type: "array",
          items: {
            type: "object",
            properties: {
              id: { type: "string" },
              name: { type: "string" },
              type: { type: "string" },
              left: { type: "number" },
              top: { type: "number" },
              width: { type: "number" },
              height: { type: "number" },
            },
          },
        },
      },
    },
    example: {
      request: { kind: "getShapes", slideIndex: 0 },
      response: { ok: true, result: { slideIndex: 0, shapes: [{ id: "4", name: "Title 1", type: "GeometricShape", left: 50, top: 30, width: 620, height: 90 }] } },
    },
  },

  {
    name: "setShapeText",
    host: "powerpoint",
    description: "Overwrite the text of one shape. Target by shapeId (stable) or shapeName. Text is replaced whole — no anchor search.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "setShapeText" },
        slideIndex: { type: "integer", minimum: 0, default: 0 },
        shapeId: { type: "string", description: "Preferred. Use getShapes to obtain." },
        shapeName: { type: "string", description: "Fallback when shapeId is not available." },
        text: { type: "string", description: "Replacement text." },
      },
      required: ["kind", "text"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { ok: { type: "boolean" }, slideIndex: { type: "integer" }, shapeName: { type: "string" } },
    },
    example: {
      request: { kind: "setShapeText", slideIndex: 0, shapeName: "Title 1", text: "Introduction (revised)" },
      response: { ok: true, result: { ok: true, slideIndex: 0, shapeName: "Title 1" } },
    },
  },

  {
    name: "findReplaceAll",
    host: "powerpoint",
    description: "Exact-text find and replace across every text-bearing shape in the presentation. Regex-escaped — plain-text only.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "findReplaceAll" },
        find: { type: "string" },
        replace: { type: "string" },
        matchCase: { type: "boolean", default: false },
      },
      required: ["kind", "find", "replace"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { matched: { type: "integer" }, replaced: { type: "integer" } },
    },
    example: {
      request: { kind: "findReplaceAll", find: "DRAFT", replace: "FINAL" },
      response: { ok: true, result: { matched: 4, replaced: 4 } },
    },
  },

  {
    name: "addSlide",
    host: "powerpoint",
    description: "Append a new blank slide using the default layout. Use addSlideWithLayout to control layout or insertion index.",
    inputSchema: {
      type: "object",
      properties: { kind: { const: "addSlide" } },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { ok: { type: "boolean" }, totalSlides: { type: "integer" } },
    },
    example: {
      request: { kind: "addSlide" },
      response: { ok: true, result: { ok: true, totalSlides: 4 } },
    },
  },

  {
    name: "addSlideWithLayout",
    host: "powerpoint",
    description: "Add a new slide with a specific layout and optionally move it to a target index. Use getLayouts to list available layout ids.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "addSlideWithLayout" },
        layoutId: { type: "string" },
        slideMasterId: { type: "string" },
        index: { type: "integer", minimum: 0, description: "Zero-based position to move the new slide to." },
      },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { ok: { type: "boolean" }, totalSlides: { type: "integer" } },
    },
    example: {
      request: { kind: "addSlideWithLayout", layoutId: "2147483649", index: 1 },
      response: { ok: true, result: { ok: true, totalSlides: 4 } },
    },
  },

  {
    name: "deleteSlide",
    host: "powerpoint",
    description: "Delete the slide at the given zero-based index. Not reversible through the bridge.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "deleteSlide" },
        slideIndex: { type: "integer", minimum: 0 },
      },
      required: ["kind", "slideIndex"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { ok: { type: "boolean" }, deleted: { type: "integer" } },
    },
    example: {
      request: { kind: "deleteSlide", slideIndex: 2 },
      response: { ok: true, result: { ok: true, deleted: 2 } },
    },
  },

  {
    name: "duplicateSlide",
    host: "powerpoint",
    description: "Duplicate a slide via export-and-reinsert (base64 round-trip). Optionally insert at a target index; default is end of deck.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "duplicateSlide" },
        slideIndex: { type: "integer", minimum: 0, default: 0 },
        targetIndex: { type: "integer", minimum: 0, description: "Zero-based insertion index. Omit to append." },
      },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { ok: { type: "boolean" }, totalSlides: { type: "integer" } },
    },
    example: {
      request: { kind: "duplicateSlide", slideIndex: 0, targetIndex: 1 },
      response: { ok: true, result: { ok: true, totalSlides: 4 } },
    },
  },

  {
    name: "moveSlide",
    host: "powerpoint",
    description: "Reorder a slide from slideIndex to targetIndex (both zero-based).",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "moveSlide" },
        slideIndex: { type: "integer", minimum: 0 },
        targetIndex: { type: "integer", minimum: 0 },
      },
      required: ["kind", "slideIndex", "targetIndex"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { ok: { type: "boolean" }, from: { type: "integer" }, to: { type: "integer" } },
    },
    example: {
      request: { kind: "moveSlide", slideIndex: 3, targetIndex: 0 },
      response: { ok: true, result: { ok: true, from: 3, to: 0 } },
    },
  },

  {
    name: "insertImage",
    host: "powerpoint",
    description: "Add an image to a slide from a base64 payload (PNG/JPEG). Coordinates are in points.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "insertImage" },
        slideIndex: { type: "integer", minimum: 0, default: 0 },
        base64: { type: "string", description: "Image bytes base64-encoded (no data: prefix)." },
        left: { type: "number", default: 50 },
        top: { type: "number", default: 50 },
        width: { type: "number", default: 400 },
        height: { type: "number", default: 300 },
      },
      required: ["kind", "base64"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { ok: { type: "boolean" }, slideIndex: { type: "integer" } },
    },
    example: {
      request: { kind: "insertImage", slideIndex: 0, base64: "iVBORw0KGgo...", left: 100, top: 100, width: 200, height: 150 },
      response: { ok: true, result: { ok: true, slideIndex: 0 } },
    },
  },

  {
    name: "replaceImage",
    host: "powerpoint",
    description: "Replace the fill of an existing shape with a new base64 image. Target by shapeId or shapeName.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "replaceImage" },
        slideIndex: { type: "integer", minimum: 0, default: 0 },
        shapeId: { type: "string" },
        shapeName: { type: "string" },
        base64: { type: "string" },
      },
      required: ["kind", "base64"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: { ok: { type: "boolean" }, shapeName: { type: "string" }, shapeId: { type: "string" } },
    },
    example: {
      request: { kind: "replaceImage", slideIndex: 0, shapeName: "Picture 3", base64: "iVBORw0KGgo..." },
      response: { ok: true, result: { ok: true, shapeName: "Picture 3", shapeId: "7" } },
    },
  },

  {
    name: "getSlideImage",
    host: "powerpoint",
    description: "Render a slide to a base64 PNG at a requested pixel height (width is derived from aspect ratio). Handy to feed slide contents to a vision model.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { const: "getSlideImage" },
        slideIndex: { type: "integer", minimum: 0, default: 0 },
        height: { type: "integer", minimum: 32, default: 300 },
      },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: {
        slideIndex: { type: "integer" },
        base64: { type: "string", description: "PNG bytes, base64-encoded." },
        height: { type: "integer" },
      },
    },
    example: {
      request: { kind: "getSlideImage", slideIndex: 0, height: 600 },
      response: { ok: true, result: { slideIndex: 0, base64: "iVBORw0KGgo...", height: 600 } },
    },
  },

  {
    name: "getLayouts",
    host: "powerpoint",
    description: "List every slide master and its layouts (id + name). Use the ids with addSlideWithLayout.",
    inputSchema: {
      type: "object",
      properties: { kind: { const: "getLayouts" } },
      required: ["kind"],
      additionalProperties: false,
    },
    outputSchema: {
      type: "object",
      properties: {
        masters: {
          type: "array",
          items: {
            type: "object",
            properties: {
              id: { type: "string" },
              name: { type: "string" },
              layouts: {
                type: "array",
                items: {
                  type: "object",
                  properties: { id: { type: "string" }, name: { type: "string" } },
                },
              },
            },
          },
        },
      },
    },
    example: {
      request: { kind: "getLayouts" },
      response: { ok: true, result: { masters: [{ id: "2147483648", name: "Office Theme", layouts: [{ id: "2147483649", name: "Title Slide" }, { id: "2147483650", name: "Title and Content" }] }] } },
    },
  },
];

export function getToolByName(name) {
  return tools.find((t) => t.name === name) || null;
}

export function getToolsByHost(host) {
  if (!host) return tools;
  return tools.filter((t) => t.host === host || t.host === "any");
}
