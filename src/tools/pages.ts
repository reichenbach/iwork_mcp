import { existsSync } from "node:fs";
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runJXA, OsascriptError, isCreatorStudio, creatorStudioSaveAs, creatorStudioPagesExportPDF, saveTimeout, clickMenuItem, resolveAppName, pagesInsertTableViaUI, isDraftFile, selectForMenu, closeDocument } from "../jxa.js";
import { ANNOTATIONS } from "../annotations.js";

function toolResult(text: string, isError = false) {
  return { content: [{ type: "text" as const, text }], isError };
}

async function handleJXA<T>(fn: () => Promise<T>): Promise<{ content: { type: "text"; text: string }[]; isError?: boolean }> {
  try {
    const result = await fn();
    const text = typeof result === "string" ? result : JSON.stringify(result, null, 2);
    return toolResult(text);
  } catch (err) {
    // Creator Studio auto-save can rename documents mid-operation, causing transient -1728.
    // Retry once so the document name resolution injection picks up the new name.
    if (err instanceof OsascriptError && err.appleScriptErrorCode === -1728 && isCreatorStudio()) {
      await new Promise(r => setTimeout(r, 500));
      try {
        const result = await fn();
        const text = typeof result === "string" ? result : JSON.stringify(result, null, 2);
        return toolResult(text);
      } catch (retryErr) {
        if (retryErr instanceof OsascriptError) return toolResult(retryErr.message, true);
        return toolResult(String(retryErr), true);
      }
    }
    if (err instanceof OsascriptError) {
      return toolResult(err.message, true);
    }
    return toolResult(String(err), true);
  }
}

export function registerPagesTools(server: McpServer): void {
  // The text tools edit the body in place. Setting doc.bodyText (or a whole paragraph)
  // resets the formatting of everything it covers, character styles and paragraph
  // styles alike, and deletes the inline images and tables anchored there (U+FFFC in
  // the text). Setting one character to a string replaces just that character, and the
  // new text takes its formatting; deleting a character range leaves the rest alone.
  // Pages counts characters in code points, so positions index Array.from(bodyText),
  // not the UTF-16 string (an emoji is one character to Pages, two string units to JS).
  // Pages ends a paragraph at "\\n" and at a page break ("\\f").
  const BODY_EDIT_HELPERS = `
    ObjC.import("Foundation");
    const OBJECT_CHAR = "\\uFFFC";
    function openDoc(name) {
      // By id: Creator Studio's auto-save renames the document after an edit
      const app = Application("Pages");
      return app.documents.byId(app.documents.byName(name).id());
    }
    // The body as an array of characters, indexed like doc.bodyText.characters. The edit
    // functions below keep a copy of what the body should become, for checkBody.
    let expectedBody = null;
    function bodyChars(doc) {
      const chars = Array.from(doc.bodyText());
      expectedBody = chars.slice();
      return chars;
    }
    function setBody(doc, s) {
      doc.bodyText = s;
      expectedBody = Array.from(s);
    }
    // Pages sometimes ignores a scripted text change without an error (seen while the
    // document's window was in the middle of an edit), so compare the body with what the
    // edits should have produced instead of reporting success regardless.
    function checkBody(doc) {
      if (doc.bodyText() !== expectedBody.join("")) {
        throw new Error("Pages did not apply the change to the body text. Its window may be in the middle of an edit: click into the document, or close and reopen it, and try again.");
      }
    }
    // One entry per paragraph, matching doc.bodyText.paragraphs: its characters are
    // text.slice(start, end), and sep ("\\n", "\\f" or "" for the last) sits at end.
    function paraBounds(text) {
      const paras = [];
      let start = 0;
      for (let i = 0; i < text.length; i++) {
        if (text[i] === "\\n" || text[i] === "\\f") {
          paras.push({ start: start, end: i, sep: text[i] });
          start = i + 1;
        }
      }
      paras.push({ start: start, end: text.length, sep: "" });
      return paras;
    }
    function setChar(doc, index, s) {
      doc.bodyText.characters[index] = s;
      Array.prototype.splice.apply(expectedBody, [index, 1].concat(Array.from(s)));
    }
    // A character next to which text goes in. An inline object can't be the one: setting
    // it to a string deletes the object.
    function editableChar(body, index) {
      if (body[index] === OBJECT_CHAR) {
        throw new Error("Can't add text here without removing the inline image or table next to it. Insert the text at another paragraph.");
      }
      return body[index];
    }
    function runAS(src) {
      const err = $();
      const result = $.NSAppleScript.alloc.initWithSource(src).executeAndReturnError(err);
      if (!result || (result.isNil && result.isNil())) {
        const info = ObjC.deepUnwrap(err) || {};
        throw new Error(info.NSAppleScriptErrorBriefMessage || info.NSAppleScriptErrorMessage || "AppleScript execution failed");
      }
      return result;
    }
    function charRange(doc, from, to) {  // 0-based, inclusive
      return "characters " + (from + 1) + " thru " + (to + 1) + ' of body text of document id "' + doc.id() + '"';
    }
    function deleteChars(doc, from, to) {
      if (to < from) return;
      runAS('tell application "Pages" to delete (' + charRange(doc, from, to) + ")");
      expectedBody.splice(from, to - from + 1);
    }
  `;
  // ── Document Management ──

  server.tool(
    "pages_list_documents",
    "List all open Pages documents",
    {},
    ANNOTATIONS.readOnly,
    async () => handleJXA(() => runJXA<string[]>(`
      const app = Application("Pages");
      const docs = app.documents();
      return JSON.stringify(docs.map(d => ({ name: d.name(), path: d.file() ? d.file().toString() : null })));
    `)),
  );

  server.tool(
    "pages_list_templates",
    "List all available Pages templates (e.g. Resume, Report, Letter, Flyer)",
    {},
    ANNOTATIONS.readOnly,
    async () => handleJXA(() => runJXA<string[]>(`
      const app = Application("Pages");
      return JSON.stringify(app.templates().map(t => t.name()));
    `)),
  );

  server.tool(
    "pages_create_document",
    "Create a new Pages document (optionally from a template — use pages_list_templates to see available templates)",
    {
      templateName: z.string().optional().describe("Template name (optional)"),
    },
    ANNOTATIONS.readWrite,
    async ({ templateName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      let doc;
      if (params.templateName) {
        doc = app.Document({ documentTemplate: app.templates[params.templateName] });
        app.documents.push(doc);
      } else {
        doc = app.Document();
        app.documents.push(doc);
      }
      return JSON.stringify({ name: doc.name() });
    `, { templateName: templateName ?? null })),
  );

  server.tool(
    "pages_open_document",
    "Open a document from disk: .pages, or Word, RTF and plain text files (imported as a new, unsaved document)",
    {
      filePath: z.string().startsWith("/").describe("Absolute path to the file"),
    },
    ANNOTATIONS.readWrite,
    async ({ filePath }) => handleJXA(async () => {
      // Pages shows a modal alert for missing or unsupported files, which blocks every later call
      if (!existsSync(filePath)) throw new Error(`File not found: ${filePath}`);
      if (!/\.(pages|docx?|dotx|rtfd?|txt|md|markdown)$/i.test(filePath)) {
        throw new Error(`Pages can't open this file type: ${filePath}`);
      }
      return runJXA<string>(`
      const app = Application("Pages");
      const doc = app.open(Path(params.filePath));
      return JSON.stringify({ name: doc.name() });
    `, { filePath });
    }),
  );

  server.tool(
    "pages_save_document",
    "Save a Pages document as .pages (use this to save to disk — use pages_export_document for PDF/Word/EPUB)",
    {
      documentName: z.string().describe("Name of the open document"),
      filePath: z.string().startsWith("/").optional().describe("File path to save to (for Save As)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, filePath }) => handleJXA(async () => {
      if (!filePath) {
        // New and imported documents only have an auto-saved draft; saving one opens a Save sheet that blocks
        const file = await runJXA<string | null>(`
      const f = Application("Pages").documents.byName(params.documentName).file();
      return JSON.stringify(f ? f.toString() : null);
    `, { documentName });
        if (!file || await isDraftFile(file)) {
          throw new Error("The document has never been saved (it is new or was imported from another format); pass filePath to choose where to save it");
        }
      }
      if (isCreatorStudio()) {
        if (filePath) {
          const { name, reopened } = await creatorStudioSaveAs("Pages", documentName, filePath);
          return JSON.stringify(reopened
            ? { saved: true, name }
            : { saved: true, name, reopened: false, note: `Saved, but Pages did not reopen the file; open it with pages_open_document if needed.` });
        }
        return runJXA<string>(`
      ${PAGES_TABLE_HELPERS}
      const doc = Application("Pages").documents.byName(params.documentName);
      runAS('tell application "Pages" to save ' + docRef());
      return JSON.stringify({ saved: true, name: doc.name() });
    `, { documentName }, { timeout: saveTimeout(), label: "pages:save" });
      }
      return runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      if (params.filePath) {
        doc.save({ in: Path(params.filePath) });
        doc.close({ saving: "no" });
        const newDoc = app.open(Path(params.filePath));
        return JSON.stringify({ saved: true, name: newDoc.name() });
      } else {
        doc.save();
      }
      return JSON.stringify({ saved: true, name: doc.name() });
    `, { documentName, filePath: filePath ?? null }, { timeout: saveTimeout(), label: "pages:save" });
    }),
  );

  server.tool(
    "pages_export_document",
    "Export a Pages document to a different format: PDF, Word (.docx), EPUB, or plain text (not .pages — use pages_save_document for that)",
    {
      documentName: z.string().describe("Name of the open document"),
      filePath: z.string().startsWith("/").describe("Absolute path for the exported file"),
      format: z.enum(["PDF", "Word", "EPUB", "Text"]).describe("Export format"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, filePath, format }) => handleJXA(async () => {
      const ext = { PDF: [".pdf"], Word: [".docx"], EPUB: [".epub"], Text: [".txt"] }[format];
      if (!ext.some((e) => filePath.toLowerCase().endsWith(e))) {
        throw new Error(`filePath must end in ${ext.join(" or ")} for ${format} export`);
      }
      if (isCreatorStudio() && format === "PDF") {
        const { method, pageCount, documentPageCount } = await creatorStudioPagesExportPDF(documentName, filePath);
        const result: Record<string, unknown> = { exported: true, path: filePath, format: "PDF", method, pageCount };
        if (method === "Quick Look") result.note = "Exported from the last saved state of the document";
        if (pageCount === null || documentPageCount === null) result.warning = "Could not check the page count of the exported PDF";
        return JSON.stringify(result);
      }
      return runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      const formatMap = {
        "PDF": "PDF",
        "Word": "Microsoft Word",
        "EPUB": "EPUB",
        "Text": "unformatted text",
      };
      const fmt = formatMap[params.format];
      app.export(doc, { to: Path(params.filePath), as: fmt });
      return JSON.stringify({ exported: true, path: params.filePath, format: params.format });
    `, { documentName, filePath, format });
    }),
  );

  server.tool(
    "pages_close_document",
    "Close a Pages document. With saving \"no\", a document that was never saved (new, or imported from another format) is discarded, including its iCloud draft (moved to the Trash). Such a document can't be closed with saving \"yes\": save it with a file path first.",
    {
      documentName: z.string().describe("Name of the open document"),
      saving: z.enum(["yes", "no", "ask"]).optional().describe("Whether to save before closing"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, saving }) => handleJXA(() => closeDocument("Pages", documentName, saving)),
  );

  // ── Text Reading Tools ──

  server.tool(
    "pages_get_body_text",
    "Read all body text from a Pages document",
    {
      documentName: z.string().describe("Name of the open document"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      const text = doc.bodyText();
      return JSON.stringify({ text: text });
    `, { documentName })),
  );

  server.tool(
    "pages_get_paragraphs",
    "Get all paragraphs from a Pages document with their text, font, size, and color",
    {
      documentName: z.string().describe("Name of the open document"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      const paras = doc.bodyText.paragraphs;
      const count = paras.length;
      const result = [];
      for (let i = 0; i < count; i++) {
        const p = paras[i];
        const text = p().replace(/\\n$/, "");
        if (text === "" && i === count - 1) continue;  // skip trailing empty paragraph
        const entry = { index: i, text: text };
        try { entry.font = p.font(); } catch(e) {}
        try { entry.size = p.size(); } catch(e) {}
        try {
          const c = p.color();
          if (Array.isArray(c) && c.length >= 3) {
            const toHex = (v) => Math.round(v * 255).toString(16).padStart(2, "0");
            entry.color = "#" + toHex(c[0]) + toHex(c[1]) + toHex(c[2]);
          }
        } catch(e) {}
        result.push(entry);
      }
      return JSON.stringify(result);
    `, { documentName })),
  );

  server.tool(
    "pages_get_document_info",
    "Get document metadata: word count, character count, page count, modified status",
    {
      documentName: z.string().describe("Name of the open document"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      const result = {
        name: doc.name(),
        modified: doc.modified(),
        passwordProtected: doc.passwordProtected(),
      };
      try { result.wordCount = doc.bodyText.words().length; } catch(e) { result.wordCount = null; }
      try { result.characterCount = doc.bodyText.characters().length; } catch(e) { result.characterCount = null; }
      try { result.pageCount = doc.pages.length; } catch(e) { result.pageCount = null; }
      try { result.paragraphCount = doc.bodyText.paragraphs.length; } catch(e) { result.paragraphCount = null; }
      try { result.filePath = doc.file() ? doc.file().toString() : null; } catch(e) { result.filePath = null; }
      return JSON.stringify(result);
    `, { documentName })),
  );

  server.tool(
    "pages_list_images",
    "List all images in a Pages document with their positions and sizes (index is 0-based)",
    {
      documentName: z.string().describe("Name of the open document"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      const imgs = doc.images();
      var result = [];
      for (var i = 0; i < imgs.length; i++) {
        var img = imgs[i];
        var entry = { index: i };
        try { entry.width = img.width(); } catch(e) {}
        try { entry.height = img.height(); } catch(e) {}
        try { var p = img.position(); entry.position = { x: p.x || p[0] || 0, y: p.y || p[1] || 0 }; } catch(e) {}
        try { entry.fileName = img.fileName(); } catch(e) {}
        try { entry.description = img.objectDescription(); } catch(e) {}
        result.push(entry);
      }
      return JSON.stringify(result);
    `, { documentName })),
  );

  // ── Text Writing Tools ──

  server.tool(
    "pages_add_text",
    "Append text to the end of the document body. It takes the formatting of the text it follows; existing text, images and tables are left as they are. Include a trailing newline to start a new paragraph.",
    {
      documentName: z.string().describe("Name of the open document"),
      text: z.string().describe("Text to append (include trailing newline for a new paragraph)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, text }) => handleJXA(() => runJXA<string>(`
      ${BODY_EDIT_HELPERS}
      const doc = openDoc(params.documentName);
      const body = bodyChars(doc);
      if (!body.length) setBody(doc, params.text);
      else if (params.text) setChar(doc, body.length - 1, editableChar(body, body.length - 1) + params.text);
      checkBody(doc);
      return JSON.stringify({ appended: true, paragraphCount: doc.bodyText.paragraphs.length });
    `, { documentName, text })),
  );

  server.tool(
    "pages_insert_text_at",
    "Insert text as new paragraphs after a paragraph. The new paragraphs get the character and paragraph formatting of the paragraph before them (of paragraph 0 when inserting at the beginning); other text, images and tables are left as they are.",
    {
      documentName: z.string().describe("Name of the open document"),
      text: z.string().describe("Text to insert; newlines separate paragraphs (a trailing newline is ignored)"),
      afterParagraph: z.number().int().min(-1).describe("Insert after this paragraph index (0-based). Use -1 to insert at the beginning."),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, text, afterParagraph }) => handleJXA(() => runJXA<string>(`
      ${BODY_EDIT_HELPERS}
      const doc = openDoc(params.documentName);
      const body = bodyChars(doc);
      const paras = paraBounds(body);
      if (params.afterParagraph >= paras.length) {
        throw new Error("Paragraph index " + params.afterParagraph + " out of range (0-" + (paras.length - 1) + ")");
      }
      const newText = params.text.endsWith("\\n") ? params.text.slice(0, -1) : params.text;
      if (!body.length) {
        setBody(doc, params.afterParagraph < 0 ? newText + "\\n" : "\\n" + newText);
      } else if (params.afterParagraph < 0) {
        setChar(doc, 0, newText + "\\n" + editableChar(body, 0));
      } else {
        const p = paras[params.afterParagraph];
        // Replace the paragraph's separator with itself plus the new paragraphs; after the
        // last paragraph, which has none, extend its last character instead.
        if (p.sep) setChar(doc, p.end, p.sep + newText + "\\n");
        else setChar(doc, body.length - 1, editableChar(body, body.length - 1) + "\\n" + newText);
      }
      checkBody(doc);
      return JSON.stringify({ inserted: true, paragraphCount: doc.bodyText.paragraphs.length });
    `, { documentName, text, afterParagraph })),
  );

  server.tool(
    "pages_delete_text",
    "Delete a paragraph by index, with any inline images or tables in it. Other paragraphs keep their text and formatting.",
    {
      documentName: z.string().describe("Name of the open document"),
      paragraphIndex: z.number().int().min(0).describe("Paragraph index to delete (0-based)"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, paragraphIndex }) => handleJXA(() => runJXA<string>(`
      ${BODY_EDIT_HELPERS}
      const doc = openDoc(params.documentName);
      const paras = paraBounds(bodyChars(doc));
      const i = params.paragraphIndex;
      if (i < 0 || i >= paras.length) {
        throw new Error("Paragraph index " + i + " out of range (0-" + (paras.length - 1) + ")");
      }
      const p = paras[i];
      if (paras.length === 1) deleteChars(doc, p.start, p.end - 1);
      // The last paragraph, or one that ends with a page break: delete it together with the
      // separator before it, so the paragraph before ends the body or takes over the break.
      else if (i > 0 && p.sep !== "\\n") deleteChars(doc, paras[i - 1].end, p.end - 1);
      else deleteChars(doc, p.start, p.end);
      checkBody(doc);
      return JSON.stringify({ deleted: true, paragraphCount: doc.bodyText.paragraphs.length });
    `, { documentName, paragraphIndex })),
  );

  server.tool(
    "pages_replace_text",
    "Find and replace text in a Pages document. The replacement takes the formatting of the first character it replaces; the rest of the paragraph, other text, inline images and tables are left as they are.",
    {
      documentName: z.string().describe("Name of the open document"),
      find: z.string().min(1).describe("Text to find"),
      replace: z.string().describe("Replacement text"),
      all: z.boolean().optional().describe("Replace all occurrences (default: true)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, find, replace, all }) => handleJXA(() => runJXA<string>(`
      ${BODY_EDIT_HELPERS}
      const doc = openDoc(params.documentName);
      const text = bodyChars(doc).join("");
      const findLength = Array.from(params.find).length;
      const replaceLength = Array.from(params.replace).length;
      // Character positions of the matches (code points, like doc.bodyText.characters)
      const hits = [];
      let from = 0, pos = 0, idx;
      while ((idx = text.indexOf(params.find, from)) !== -1) {
        pos += Array.from(text.slice(from, idx)).length;
        hits.push(pos);
        pos += findLength;
        from = idx + params.find.length;
        if (params.all === false) break;
      }
      // From the last match back, so the positions before it stay valid
      for (let h = hits.length - 1; h >= 0; h--) {
        const k = hits[h];
        if (!replaceLength) {
          deleteChars(doc, k, k + findLength - 1);
        } else {
          setChar(doc, k, params.replace);
          // Make sure Pages took the first change before deleting the rest of the match
          if (h === hits.length - 1) checkBody(doc);
          deleteChars(doc, k + replaceLength, k + replaceLength + findLength - 2);
        }
      }
      if (hits.length) checkBody(doc);
      return JSON.stringify({ replacements: hits.length });
    `, { documentName, find, replace, all: all ?? true })),
  );

  server.tool(
    "pages_format_text",
    "Set formatting on a paragraph, or on part of it: font (PostScript name), size, color. For bold use a bold font name like 'HelveticaNeue-Bold', for italic use 'HelveticaNeue-Italic'. Pass text to format only the first occurrence of that text in the paragraph, e.g. a bold 'Figure 1:' at the start of a caption.",
    {
      documentName: z.string().describe("Name of the open document"),
      paragraphIndex: z.number().int().min(0).describe("Paragraph index (0-based)"),
      text: z.string().min(1).optional().describe("Format only the first occurrence of this text in the paragraph (default: the whole paragraph)"),
      format: z.object({
        fontSize: z.number().positive().optional().describe("Font size in points"),
        fontName: z.string().optional().describe("PostScript font name (e.g. 'HelveticaNeue-Bold' for bold, 'Georgia-Italic' for italic)"),
        textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Text color as hex, e.g. '#FF0000'"),
      }).refine((f) => Object.values(f).some((v) => v !== undefined), "Set at least one of fontSize, fontName or textColor").describe("Formatting options"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, paragraphIndex, text, format }) => handleJXA(() => runJXA<string>(`
      ${BODY_EDIT_HELPERS}
      const doc = openDoc(params.documentName);
      const paraCount = doc.bodyText.paragraphs.length;
      if (params.paragraphIndex >= paraCount) {
        throw new Error("Paragraph index " + params.paragraphIndex + " out of range (0-" + (paraCount - 1) + ")");
      }
      const fmt = params.format;
      let color = null;
      if (fmt.textColor !== undefined) {
        const hex = fmt.textColor;
        color = [parseInt(hex.slice(1, 3), 16) * 257, parseInt(hex.slice(3, 5), 16) * 257, parseInt(hex.slice(5, 7), 16) * 257];
      }

      if (params.text === null) {
        const paragraph = doc.bodyText.paragraphs[params.paragraphIndex];
        if (fmt.fontSize !== undefined) paragraph.size = fmt.fontSize;
        if (fmt.fontName !== undefined) paragraph.font = fmt.fontName;
        if (color) paragraph.color = color;
        return JSON.stringify({ formatted: true, paragraphIndex: params.paragraphIndex });
      }

      // Part of the paragraph: set the properties on a character range through AppleScript
      const chars = bodyChars(doc);
      const p = paraBounds(chars)[params.paragraphIndex];
      const paraText = chars.slice(p.start, p.end).join("");
      const idx = paraText.indexOf(params.text);
      if (idx === -1) throw new Error(JSON.stringify(params.text) + " not found in paragraph " + params.paragraphIndex);
      const from = p.start + Array.from(paraText.slice(0, idx)).length;
      const range = "(" + charRange(doc, from, from + Array.from(params.text).length - 1) + ")";
      const lines = [];
      if (fmt.fontSize !== undefined) lines.push("set size of " + range + " to " + fmt.fontSize);
      // A JSON string is a valid AppleScript string for a font name (no control characters)
      if (fmt.fontName !== undefined) lines.push("set font of " + range + " to " + JSON.stringify(fmt.fontName));
      if (color) lines.push("set color of " + range + " to {" + color.join(", ") + "}");
      runAS('tell application "Pages"\\n' + lines.join("\\n") + "\\nend tell");
      return JSON.stringify({ formatted: true, paragraphIndex: params.paragraphIndex, text: params.text });
    `, { documentName, paragraphIndex, text: text ?? null, format })),
  );

  server.tool(
    "pages_add_image",
    "Insert an image into the document",
    {
      documentName: z.string().describe("Name of the open document"),
      filePath: z.string().startsWith("/").describe("Absolute path to the image file"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, filePath }) => handleJXA(async () => {
      if (!existsSync(filePath)) throw new Error(`Image file not found: ${filePath}`);
      // Creating an image at document level fails with -2763 "Don't know how to
      // create TMAScriptImageInfoProxy" (in JXA and AppleScript alike); creating it
      // on a page works. The file must be a Path — a plain string silently no-ops.
      return runJXA<string>(`
        const app = Application("Pages");
        const doc = app.documents.byName(params.documentName);
        const before = doc.images.length;
        const pageIndex = doc.pages.length - 1;
        let pageErr = null;
        try {
          doc.pages[pageIndex].images.push(app.Image({ file: Path(params.filePath) }));
        } catch (e) { pageErr = e; }
        if (doc.images.length === before) {
          try { doc.images.push(app.Image({ file: Path(params.filePath) })); } catch (e) { throw pageErr || e; }
        }
        if (doc.images.length === before) throw new Error("Pages did not insert the image");
        const imgs = doc.images();
        const img = imgs[imgs.length - 1];
        return JSON.stringify({ added: true, path: params.filePath, page: pageIndex + 1, width: img.width(), height: img.height() });
      `, { documentName, filePath });
    }),
  );

  server.tool(
    "pages_add_table",
    "Insert a table into the document",
    {
      documentName: z.string().describe("Name of the open document"),
      rows: z.number().int().positive().optional().describe("Number of rows (omit to keep the app's default table size)"),
      columns: z.number().int().positive().optional().describe("Number of columns (omit to keep the app's default table size)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, rows, columns }) => handleJXA(async () => {
      // Pages 15.x (Creator Studio): doc.tables.push() and AppleScript
      // `make new table` both fail with -2763. Insert via menu UI scripting.
      if (resolveAppName("Pages") !== "Pages") {
        const name = await pagesInsertTableViaUI(documentName, rows, columns);
        return JSON.stringify({ added: true, name });
      }
      return runJXA<string>(`
        const app = Application("Pages");
        const doc = app.documents.byName(params.documentName);
        const props = {};
        if (params.rows) props.rowCount = params.rows;
        if (params.columns) props.columnCount = params.columns;
        const table = app.Table(props);
        doc.tables.push(table);
        return JSON.stringify({ added: true, name: table.name() });
      `, { documentName, rows: rows ?? null, columns: columns ?? null });
    }),
  );

  // ── Table Cell Tools (AppleScript bridge) ──
  // Pages tables are unreachable via JXA (-2763 TMAScriptTableInfoProxy), but the
  // AppleScript dictionary exposes them fully — same workaround as the Keynote
  // master-slide tools: run AppleScript through NSAppleScript from within JXA.

  const PAGES_TABLE_HELPERS = `
    ObjC.import("Foundation");
    function osType(s) { let n = 0; for (let i = 0; i < 4; i++) n = n * 256 + s.charCodeAt(i); return n; }
    function runAS(src) {
      const script = $.NSAppleScript.alloc.initWithSource(src);
      const err = $();
      const result = script.executeAndReturnError(err);
      if (!result || (result.isNil && result.isNil())) {
        const errInfo = ObjC.deepUnwrap(err) || {};
        throw new Error(errInfo.NSAppleScriptErrorBriefMessage || errInfo.NSAppleScriptErrorMessage || "AppleScript execution failed");
      }
      return result;
    }
    function descToJS(d) {
      const t = d.descriptorType;
      if (t === osType("msng")) return null;
      if (t === osType("doub")) return d.doubleValue;
      if (t === osType("long")) return d.int32Value;
      if (t === osType("bool")) return d.booleanValue;
      if (t === osType("true")) return true;
      if (t === osType("fals")) return false;
      if (t === osType("list")) {
        const arr = [];
        for (let i = 1; i <= d.numberOfItems; i++) arr.push(descToJS(d.descriptorAtIndex(i)));
        return arr;
      }
      const s = d.stringValue;
      return (s && !s.isNil()) ? s.js : null;
    }
    function asQuote(s) {
      return '"' + String(s).replace(/\\\\/g, "\\\\\\\\").replace(/"/g, '\\\\"').replace(/\\n/g, "\\\\n").replace(/\\r/g, "\\\\r").replace(/\\t/g, "\\\\t") + '"';
    }
    function asValue(v) {
      if (v === null || v === undefined) return '""';
      if (typeof v === "number") return String(v);
      if (typeof v === "boolean") return v ? "true" : "false";
      return asQuote(v);
    }
    // Creator Studio auto-save transiently renames docs ("X" -> "X.pages" and
    // back) after edits, so name-based AppleScript references race the flip
    // even when a name probe just succeeded. Resolve the document's stable id
    // once (polling both names across the rename window) and reference
    // 'document id "UUID"' everywhere — immune to renames.
    (function resolveDocAS() {
      const extName = params.documentName + ".pages";
      for (let attempt = 0; attempt < 8; attempt++) {
        for (const nm of [params.documentName, extName]) {
          try {
            const docId = descToJS(runAS('tell application "Pages" to id of document ' + asQuote(nm)));
            if (docId) { params.__docRef = 'document id ' + asQuote(docId); return; }
          } catch (e) {}
        }
        delay(0.3);
      }
      // unreachable doc: fall back to the name; the original error will surface
      params.__docRef = "document " + asQuote(params.documentName);
    })();
    function docRef() { return params.__docRef; }
    function tableRef() {
      return "table " + params.tableIndex + " of " + docRef();
    }
  `;

  server.tool(
    "pages_list_tables",
    "List all tables in a Pages document with their name, size, and header configuration",
    {
      documentName: z.string().describe("Name of the open document"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName }) => handleJXA(() => runJXA<string>(`
      ${PAGES_TABLE_HELPERS}
      const src = 'tell application "Pages"\\n' +
        'tell ' + docRef() + '\\n' +
        'set out to {}\\n' +
        'repeat with t in tables\\n' +
        'set end of out to {name of t, row count of t, column count of t, header row count of t, header column count of t}\\n' +
        'end repeat\\n' +
        'return out\\n' +
        'end tell\\n' +
        'end tell';
      const list = descToJS(runAS(src)) || [];
      return JSON.stringify(list.map(function (t, i) {
        return { index: i + 1, name: t[0], rows: t[1], columns: t[2], headerRows: t[3], headerColumns: t[4] };
      }));
    `, { documentName })),
  );

  server.tool(
    "pages_read_table",
    "Read all cell values and formulas from a table in a Pages document (1-based table index; use pages_list_tables to see available tables)",
    {
      documentName: z.string().describe("Name of the open document"),
      tableIndex: z.number().int().min(1).optional().describe("Table index, 1-based (default: 1)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, tableIndex }) => handleJXA(() => runJXA<string>(`
      ${PAGES_TABLE_HELPERS}
      const src = 'tell application "Pages"\\n' +
        'tell ' + tableRef() + '\\n' +
        'set vals to {}\\n' +
        'set fmls to {}\\n' +
        'repeat with r from 1 to count of rows\\n' +
        'set end of vals to (value of every cell of row r)\\n' +
        'set end of fmls to (formula of every cell of row r)\\n' +
        'end repeat\\n' +
        'return {name, row count, column count, header row count, vals, fmls}\\n' +
        'end tell\\n' +
        'end tell';
      const r = descToJS(runAS(src));
      return JSON.stringify({ name: r[0], rows: r[1], columns: r[2], headerRows: r[3], values: r[4], formulas: r[5] });
    `, { documentName, tableIndex: tableIndex ?? 1 })),
  );

  server.tool(
    "pages_write_table_cells",
    "Write values into table cells in a Pages document (1-based row/column). Strings starting with '=' become formulas (e.g. '=SUM(D2:D5)'); formulas read back with the app's localized function names. Strings that would lose digits as numbers (e.g. '0049') are stored as text; booleans are written as =TRUE()/=FALSE() so they read back as true/false. Existing formulas in other cells recalculate automatically. By default the table grows to fit out-of-range writes.",
    {
      documentName: z.string().describe("Name of the open document"),
      tableIndex: z.number().int().min(1).optional().describe("Table index, 1-based (default: 1)"),
      cells: z.array(z.object({
        row: z.number().int().min(1).describe("Row number (1-based, header row is 1)"),
        column: z.number().int().min(1).describe("Column number (1-based)"),
        value: z.union([z.string(), z.number(), z.boolean(), z.null()]).describe("Cell value; string starting with '=' sets a formula; null clears the cell"),
      })).min(1).describe("Cells to write"),
      autoGrow: z.boolean().optional().describe("Grow the table if a write is beyond current bounds (default: true)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, tableIndex, cells, autoGrow }) => handleJXA(() => runJXA<string>(`
      ${PAGES_TABLE_HELPERS}
      let maxRow = 0, maxCol = 0;
      for (const c of params.cells) {
        if (c.row > maxRow) maxRow = c.row;
        if (c.column > maxCol) maxCol = c.column;
      }
      let body = '';
      if (params.autoGrow) {
        body += 'if row count < ' + maxRow + ' then set row count to ' + maxRow + '\\n';
        body += 'if column count < ' + maxCol + ' then set column count to ' + maxCol + '\\n';
      }
      for (const c of params.cells) {
        const cellAS = 'cell ' + c.column + ' of row ' + c.row;
        // Keep strings like "0049" or "1.50" as text instead of numbers. Booleans would be stored
        // as 1/0 (Pages can't set the checkbox format by script), but TRUE()/FALSE() keep the type.
        const v = c.value;
        const lossy = typeof v === "string" && v.trim() !== "" && v[0] !== "=" && !isNaN(Number(v)) && String(Number(v)) !== v;
        // A cell holding a boolean coerces whatever is written next (7 or "0049" -> true) and
        // refuses the text format; emptying it first resets the type, the format stays
        if (v !== null) body += 'set value of ' + cellAS + ' to ""\\n';
        if (lossy) body += 'try\\nset format of ' + cellAS + ' to text\\nend try\\n';
        // In a text-formatted cell a formula would be stored as plain text
        if (typeof v === "boolean" || (typeof v === "string" && v[0] === "=")) body += 'try\\nset format of ' + cellAS + ' to automatic\\nend try\\n';
        const asV = typeof v === "boolean" ? asQuote(v ? "=TRUE()" : "=FALSE()") : asValue(v);
        body += 'set value of ' + cellAS + ' to ' + asV + '\\n';
      }
      const src = 'tell application "Pages"\\n' +
        'tell ' + tableRef() + '\\n' +
        body +
        'return {row count, column count}\\n' +
        'end tell\\n' +
        'end tell';
      const dims = descToJS(runAS(src));
      return JSON.stringify({ written: params.cells.length, rows: dims[0], columns: dims[1] });
    `, { documentName, tableIndex: tableIndex ?? 1, cells, autoGrow: autoGrow ?? true })),
  );

  server.tool(
    "pages_resize_table",
    "Change the row and/or column count of a table in a Pages document. Shrinking deletes the content of removed rows/columns.",
    {
      documentName: z.string().describe("Name of the open document"),
      tableIndex: z.number().int().min(1).optional().describe("Table index, 1-based (default: 1)"),
      rows: z.number().int().min(1).optional().describe("New row count"),
      columns: z.number().int().min(1).optional().describe("New column count"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, tableIndex, rows, columns }) => handleJXA(() => runJXA<string>(`
      ${PAGES_TABLE_HELPERS}
      let body = '';
      if (params.rows) body += 'set row count to ' + params.rows + '\\n';
      if (params.columns) body += 'set column count to ' + params.columns + '\\n';
      if (!body) throw new Error("Provide rows and/or columns");
      const src = 'tell application "Pages"\\n' +
        'tell ' + tableRef() + '\\n' +
        body +
        'return {row count, column count}\\n' +
        'end tell\\n' +
        'end tell';
      const dims = descToJS(runAS(src));
      return JSON.stringify({ rows: dims[0], columns: dims[1] });
    `, { documentName, tableIndex: tableIndex ?? 1, rows: rows ?? null, columns: columns ?? null })),
  );

  // ── Page Break Tool ──

  server.tool(
    "pages_insert_page_break",
    "Insert a page break after a specific paragraph. Paragraph indices stay the same, except after the last paragraph, where an empty paragraph starts the new page.",
    {
      documentName: z.string().describe("Name of the open document"),
      afterParagraph: z.number().int().min(0).describe("Insert page break after this paragraph index (0-based)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, afterParagraph }) => handleJXA(() => runJXA<string>(`
      ${BODY_EDIT_HELPERS}
      const doc = openDoc(params.documentName);
      const body = bodyChars(doc);
      const paras = paraBounds(body);
      if (params.afterParagraph >= paras.length) {
        throw new Error("Paragraph index " + params.afterParagraph + " out of range (0-" + (paras.length - 1) + ")");
      }

      // A form feed ends a paragraph with a page break. Replace the paragraph's separator
      // with it; adding one before the "\\n" would leave an empty paragraph on the new page.
      // The last paragraph has no separator, so the break goes after its last character
      // and an empty paragraph starts the new page.
      const p = paras[params.afterParagraph];
      if (p.sep === "\\f") throw new Error("Paragraph " + params.afterParagraph + " is already followed by a page break");
      if (p.sep) setChar(doc, p.end, "\\f");
      else if (!body.length) setBody(doc, "\\f");
      else setChar(doc, body.length - 1, editableChar(body, body.length - 1) + "\\f");
      checkBody(doc);

      return JSON.stringify({ pageBreakInserted: true, afterParagraph: params.afterParagraph, paragraphCount: doc.bodyText.paragraphs.length });
    `, { documentName, afterParagraph })),
  );

  // ── Compound Tools ──

  server.tool(
    "pages_create_document_with_content",
    "Create a Pages document with multiple formatted paragraphs in one call (much faster than adding paragraphs individually). For bold/italic, use PostScript font names like 'HelveticaNeue-Bold' or 'Georgia-Italic'.",
    {
      paragraphs: z.array(z.object({
        text: z.string().describe("Paragraph text (no trailing newline needed; a newline inside it starts another paragraph with the same formatting)"),
        fontSize: z.number().positive().optional().describe("Font size in points (default: the template's body size)"),
        fontName: z.string().optional().describe("PostScript font name, e.g. 'HelveticaNeue-Bold', 'Georgia-Italic'"),
        textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Text color as hex, e.g. '#FF0000'"),
      })).describe("Array of paragraphs with optional formatting"),
      filePath: z.string().startsWith("/").optional().describe("Absolute path to save as .pages file"),
    },
    ANNOTATIONS.readWrite,
    async ({ paragraphs, filePath }) => handleJXA(() => runJXA<string>(`
      ObjC.import("AppKit");
      // A wrong font name fails halfway with an opaque -1700; check them before creating anything
      const badFonts = [];
      for (const p of params.paragraphs) {
        if (p.fontName && $.NSFont.fontWithNameSize(p.fontName, 12).isNil() && badFonts.indexOf(p.fontName) < 0) badFonts.push(p.fontName);
      }
      if (badFonts.length) throw new Error("Unknown font name(s): " + badFonts.join(", ") + ". Use PostScript names like 'HelveticaNeue-Bold'.");

      const app = Application("Pages");
      let doc = app.Document();
      app.documents.push(doc);
      try {
        // Build full text with newlines between paragraphs
        let fullText = "";
        for (let i = 0; i < params.paragraphs.length; i++) {
          fullText += params.paragraphs[i].text + "\\n";
        }
        doc.bodyText = fullText;

        // Format each paragraph via bodyText.paragraphs
        // A single input paragraph with \\n creates multiple Pages paragraphs,
        // so we track the real index offset and apply formatting to all sub-paragraphs.
        let paraIdx = 0;
        for (let i = 0; i < params.paragraphs.length; i++) {
          const p = params.paragraphs[i];
          const nlCount = (p.text.match(/\\n/g) || []).length;
          let r, g, b;
          if (p.textColor !== undefined) {
            const hex = p.textColor;
            r = parseInt(hex.slice(1, 3), 16) * 257;
            g = parseInt(hex.slice(3, 5), 16) * 257;
            b = parseInt(hex.slice(5, 7), 16) * 257;
          }
          for (let j = 0; j <= nlCount; j++) {
            const para = doc.bodyText.paragraphs[paraIdx + j];
            if (p.fontSize !== undefined) para.size = p.fontSize;
            if (p.fontName !== undefined) para.font = p.fontName;
            if (r !== undefined) para.color = [r, g, b];
          }
          paraIdx += nlCount + 1;
        }
      } catch (e) {
        // Don't leave a half-filled untitled document (or its iCloud auto-save) behind
        let autoSave = null;
        try { autoSave = doc.file() ? doc.file().toString() : null; } catch (e2) {}
        try { doc.close({ saving: "no" }); } catch (e2) {}
        if (autoSave && autoSave.indexOf("/Library/Mobile Documents/") >= 0) {
          delay(1);
          $.NSFileManager.defaultManager.trashItemAtURLResultingItemURLError($.NSURL.fileURLWithPath(autoSave), null, null);
        }
        throw e;
      }

      if (params.filePath) {
        doc.save({ in: Path(params.filePath) });
        doc.close({ saving: "no" });
        const newDoc = app.open(Path(params.filePath));
        return JSON.stringify({ name: newDoc.name(), paragraphCount: params.paragraphs.length });
      }

      return JSON.stringify({ name: doc.name(), paragraphCount: params.paragraphs.length });
    `, {
      paragraphs,
      filePath: filePath ?? null,
    })),
  );

  // ── Creator Studio Features (require subscription) ──

  server.tool(
    "pages_super_resolution",
    "Upscale an image in a document using AI Super Resolution (Creator Studio only). Increases resolution while preserving quality.",
    {
      documentName: z.string().describe("Name of the open document"),
      imageIndex: z.number().int().min(1).describe("Image index, 1-based (pages_list_images index + 1)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, imageIndex }) => {
      if (!isCreatorStudio()) {
        return toolResult("Super Resolution requires Apple Creator Studio (iWork 15.1+).", true);
      }
      return handleJXA(async () => {
        const docName = await runJXA<string>(`
          const doc = Application("Pages").documents.byName(params.documentName);
          const images = doc.images();
          if (params.imageIndex > images.length) throw new Error("Image index " + params.imageIndex + " out of range (document has " + images.length + " images)");
          return JSON.stringify(doc.name());
        `, { documentName, imageIndex },
        { label: "super_resolution:check" });
        await selectForMenu("Pages", docName, "image " + imageIndex);

        await clickMenuItem("Pages", ["Format", "Image", "Super Resolution"], { postdelay: 2 });

        return JSON.stringify({ success: true, message: "Super Resolution started on image " + imageIndex + "." });
      });
    },
  );

  server.tool(
    "pages_remove_background",
    "Remove the background from an image in a document using AI (Creator Studio only).",
    {
      documentName: z.string().describe("Name of the open document"),
      imageIndex: z.number().int().min(1).describe("Image index, 1-based (pages_list_images index + 1)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, imageIndex }) => {
      if (!isCreatorStudio()) {
        return toolResult("Remove Background requires Apple Creator Studio (iWork 15.1+).", true);
      }
      return handleJXA(async () => {
        const docName = await runJXA<string>(`
          const doc = Application("Pages").documents.byName(params.documentName);
          const images = doc.images();
          if (params.imageIndex > images.length) throw new Error("Image index " + params.imageIndex + " out of range (document has " + images.length + " images)");
          return JSON.stringify(doc.name());
        `, { documentName, imageIndex },
        { label: "remove_background:check" });
        await selectForMenu("Pages", docName, "image " + imageIndex);

        await clickMenuItem("Pages", ["Format", "Image", "Remove Background"], { postdelay: 2 });

        return JSON.stringify({ success: true, message: "Background removal started on image " + imageIndex + "." });
      });
    },
  );
}
