import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runJXA, OsascriptError, isCreatorStudio, creatorStudioSaveAs, creatorStudioExportPDF, clickMenuItem, resolveAppName } from "../jxa.js";
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
    "Open a .pages file from disk",
    {
      filePath: z.string().startsWith("/").describe("Absolute path to the .pages file"),
    },
    ANNOTATIONS.readWrite,
    async ({ filePath }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.open(Path(params.filePath));
      return JSON.stringify({ name: doc.name() });
    `, { filePath })),
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
      if (isCreatorStudio()) {
        if (filePath) {
          const newName = await creatorStudioSaveAs("Pages", documentName, filePath);
          return JSON.stringify({ saved: true, name: newName });
        }
        return JSON.stringify({ saved: true, name: documentName });
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
    `, { documentName, filePath: filePath ?? null });
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
      if (isCreatorStudio() && format === "PDF") {
        await creatorStudioExportPDF("Pages", documentName, filePath);
        return JSON.stringify({ exported: true, path: filePath, format: "PDF" });
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
    "Close a Pages document",
    {
      documentName: z.string().describe("Name of the open document"),
      saving: z.enum(["yes", "no", "ask"]).optional().describe("Whether to save before closing"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, saving }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      const saveOpts = { yes: "yes", no: "no", ask: "ask" };
      if (params.saving) {
        doc.close({ saving: saveOpts[params.saving] });
      } else {
        doc.close();
      }
      return JSON.stringify({ closed: true });
    `, { documentName, saving: saving ?? null })),
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
        const text = p();
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
    "List all images in a Pages document with their positions and sizes",
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
    "Append text to the end of the document body (preserves existing formatting). Include a trailing newline to start a new paragraph.",
    {
      documentName: z.string().describe("Name of the open document"),
      text: z.string().describe("Text to append (include trailing newline for a new paragraph)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, text }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      // Save formatting on all existing paragraphs
      const paras = doc.bodyText.paragraphs;
      const count = paras.length;
      const formats = [];
      for (let i = 0; i < count; i++) {
        try { formats.push({ font: paras[i].font(), size: paras[i].size(), color: paras[i].color() }); }
        catch(e) { formats.push(null); }
      }
      // Append
      doc.bodyText = doc.bodyText() + params.text;
      // Restore formatting on pre-existing paragraphs
      for (let i = 0; i < formats.length; i++) {
        if (!formats[i]) continue;
        try {
          const p = doc.bodyText.paragraphs[i];
          p.font = formats[i].font;
          p.size = formats[i].size;
          p.color = formats[i].color;
        } catch(e) {}
      }
      return JSON.stringify({ appended: true, paragraphCount: doc.bodyText.paragraphs.length });
    `, { documentName, text })),
  );

  server.tool(
    "pages_insert_text_at",
    "Insert text at a specific paragraph index (preserves formatting on other paragraphs)",
    {
      documentName: z.string().describe("Name of the open document"),
      text: z.string().describe("Text to insert (include trailing newline)"),
      afterParagraph: z.number().int().min(-1).describe("Insert after this paragraph index (0-based). Use -1 to insert at the beginning."),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, text, afterParagraph }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      // Save formatting
      const paras = doc.bodyText.paragraphs;
      const count = paras.length;
      const formats = [];
      for (let i = 0; i < count; i++) {
        try { formats.push({ font: paras[i].font(), size: paras[i].size(), color: paras[i].color() }); }
        catch(e) { formats.push(null); }
      }
      // Insert via string manipulation
      const bt = doc.bodyText();
      const lines = bt.split("\\n");
      const insertIdx = params.afterParagraph < 0 ? 0 : params.afterParagraph + 1;
      const newText = params.text.endsWith("\\n") ? params.text.slice(0, -1) : params.text;
      const nlCount = (newText.match(/\\n/g) || []).length;
      lines.splice(insertIdx, 0, newText);
      doc.bodyText = lines.join("\\n");
      // Restore formatting (shift indices after insert point by number of new paragraphs)
      for (let i = 0; i < formats.length; i++) {
        if (!formats[i]) continue;
        const newIdx = i < insertIdx ? i : i + nlCount + 1;
        try {
          const p = doc.bodyText.paragraphs[newIdx];
          p.font = formats[i].font;
          p.size = formats[i].size;
          p.color = formats[i].color;
        } catch(e) {}
      }
      return JSON.stringify({ inserted: true, paragraphCount: doc.bodyText.paragraphs.length });
    `, { documentName, text, afterParagraph })),
  );

  server.tool(
    "pages_delete_text",
    "Delete a paragraph by index (preserves formatting on other paragraphs)",
    {
      documentName: z.string().describe("Name of the open document"),
      paragraphIndex: z.number().int().min(0).describe("Paragraph index to delete (0-based)"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, paragraphIndex }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      // Save formatting
      const paras = doc.bodyText.paragraphs;
      const count = paras.length;
      const formats = [];
      for (let i = 0; i < count; i++) {
        try { formats.push({ font: paras[i].font(), size: paras[i].size(), color: paras[i].color() }); }
        catch(e) { formats.push(null); }
      }
      // Delete via string manipulation
      const lines = doc.bodyText().split("\\n");
      if (params.paragraphIndex < 0 || params.paragraphIndex >= lines.length) {
        throw new Error("Paragraph index " + params.paragraphIndex + " out of range (0-" + (lines.length - 1) + ")");
      }
      lines.splice(params.paragraphIndex, 1);
      formats.splice(params.paragraphIndex, 1);
      doc.bodyText = lines.join("\\n");
      // Restore formatting
      for (let i = 0; i < formats.length; i++) {
        if (!formats[i]) continue;
        try {
          const p = doc.bodyText.paragraphs[i];
          p.font = formats[i].font;
          p.size = formats[i].size;
          p.color = formats[i].color;
        } catch(e) {}
      }
      return JSON.stringify({ deleted: true, paragraphCount: doc.bodyText.paragraphs.length });
    `, { documentName, paragraphIndex })),
  );

  server.tool(
    "pages_replace_text",
    "Find and replace text in a Pages document (preserves formatting)",
    {
      documentName: z.string().describe("Name of the open document"),
      find: z.string().describe("Text to find"),
      replace: z.string().describe("Replacement text"),
      all: z.boolean().optional().describe("Replace all occurrences (default: true)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, find, replace, all }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      const paras = doc.bodyText.paragraphs;
      const count = paras.length;
      let total = 0;
      for (let i = 0; i < count; i++) {
        const text = paras[i]();
        if (text.indexOf(params.find) === -1) continue;
        if (params.all !== false) {
          const parts = text.split(params.find);
          total += parts.length - 1;
          doc.bodyText.paragraphs[i] = parts.join(params.replace);
        } else if (total === 0) {
          const idx = text.indexOf(params.find);
          doc.bodyText.paragraphs[i] = text.substring(0, idx) + params.replace + text.substring(idx + params.find.length);
          total = 1;
          break;
        }
      }
      return JSON.stringify({ replacements: total });
    `, { documentName, find, replace, all: all ?? true })),
  );

  server.tool(
    "pages_format_text",
    "Set formatting on a paragraph: font (PostScript name), size, color. For bold use a bold font name like 'HelveticaNeue-Bold', for italic use 'HelveticaNeue-Italic'.",
    {
      documentName: z.string().describe("Name of the open document"),
      paragraphIndex: z.number().int().min(0).describe("Paragraph index (0-based)"),
      format: z.object({
        fontSize: z.number().positive().optional().describe("Font size in points"),
        fontName: z.string().optional().describe("PostScript font name (e.g. 'HelveticaNeue-Bold' for bold, 'Georgia-Italic' for italic)"),
        textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Text color as hex, e.g. '#FF0000'"),
      }).describe("Formatting options"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, paragraphIndex, format }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      const paragraph = doc.bodyText.paragraphs[params.paragraphIndex];
      const fmt = params.format;

      if (fmt.fontSize !== undefined) paragraph.size = fmt.fontSize;
      if (fmt.fontName !== undefined) paragraph.font = fmt.fontName;
      if (fmt.textColor !== undefined) {
        const hex = fmt.textColor;
        const r = parseInt(hex.slice(1, 3), 16) * 257;
        const g = parseInt(hex.slice(3, 5), 16) * 257;
        const b = parseInt(hex.slice(5, 7), 16) * 257;
        paragraph.color = [r, g, b];
      }

      return JSON.stringify({ formatted: true, paragraphIndex: params.paragraphIndex });
    `, { documentName, paragraphIndex, format })),
  );

  server.tool(
    "pages_add_image",
    "Insert an image into the document",
    {
      documentName: z.string().describe("Name of the open document"),
      filePath: z.string().startsWith("/").describe("Absolute path to the image file"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, filePath }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      const image = app.Image({ file: Path(params.filePath) });
      doc.images.push(image);
      return JSON.stringify({ added: true, path: params.filePath });
    `, { documentName, filePath })),
  );

  server.tool(
    "pages_add_table",
    "Insert a table into the document",
    {
      documentName: z.string().describe("Name of the open document"),
      rows: z.number().int().positive().optional().describe("Number of rows (default: 3)"),
      columns: z.number().int().positive().optional().describe("Number of columns (default: 3)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, rows, columns }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      const props = {};
      if (params.rows) props.rowCount = params.rows;
      if (params.columns) props.columnCount = params.columns;
      const table = app.Table(props);
      doc.tables.push(table);
      return JSON.stringify({ added: true, name: table.name() });
    `, { documentName, rows: rows ?? null, columns: columns ?? null })),
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
    function tableRef() {
      return "table " + params.tableIndex + " of document " + asQuote(params.documentName);
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
        'tell document ' + asQuote(params.documentName) + '\\n' +
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
    "Write values into table cells in a Pages document (1-based row/column). Strings starting with '=' become formulas (e.g. '=SUM(D2:D5)'). Existing formulas in other cells recalculate automatically. By default the table grows to fit out-of-range writes.",
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
        body += 'set value of cell ' + c.column + ' of row ' + c.row + ' to ' + asValue(c.value) + '\\n';
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
    "Insert a page break after a specific paragraph. Creates a visual page break in the document.",
    {
      documentName: z.string().describe("Name of the open document"),
      afterParagraph: z.number().int().min(0).describe("Insert page break after this paragraph index (0-based)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, afterParagraph }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      const doc = app.documents.byName(params.documentName);
      const paras = doc.bodyText.paragraphs;
      const count = paras.length;
      if (params.afterParagraph >= count) {
        throw new Error("Paragraph index " + params.afterParagraph + " out of range (0-" + (count - 1) + ")");
      }

      // Save formatting on all paragraphs
      const formats = [];
      for (let i = 0; i < count; i++) {
        try { formats.push({ font: paras[i].font(), size: paras[i].size(), color: paras[i].color() }); }
        catch(e) { formats.push(null); }
      }

      // Insert form feed character at the end of the target paragraph's text.
      // \f creates a visual page break in Pages.
      const lines = doc.bodyText().split("\\n");
      lines[params.afterParagraph] = lines[params.afterParagraph] + "\\f";
      doc.bodyText = lines.join("\\n");

      // Restore formatting — the \f creates one extra paragraph, so shift indices after the break
      for (let i = 0; i < formats.length; i++) {
        if (!formats[i]) continue;
        const newIdx = i <= params.afterParagraph ? i : i + 1;
        try {
          const p = doc.bodyText.paragraphs[newIdx];
          p.font = formats[i].font;
          p.size = formats[i].size;
          p.color = formats[i].color;
        } catch(e) {}
      }

      return JSON.stringify({ pageBreakInserted: true, afterParagraph: params.afterParagraph });
    `, { documentName, afterParagraph })),
  );

  // ── Compound Tools ──

  server.tool(
    "pages_create_document_with_content",
    "Create a Pages document with multiple formatted paragraphs in one call (much faster than adding paragraphs individually). For bold/italic, use PostScript font names like 'HelveticaNeue-Bold' or 'Georgia-Italic'.",
    {
      paragraphs: z.array(z.object({
        text: z.string().describe("Paragraph text (no trailing newline needed)"),
        fontSize: z.number().positive().optional().describe("Font size in points (default: 12)"),
        fontName: z.string().optional().describe("PostScript font name, e.g. 'HelveticaNeue-Bold', 'Georgia-Italic'"),
        textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Text color as hex, e.g. '#FF0000'"),
      })).describe("Array of paragraphs with optional formatting"),
      filePath: z.string().startsWith("/").optional().describe("Absolute path to save as .pages file"),
    },
    ANNOTATIONS.readWrite,
    async ({ paragraphs, filePath }) => handleJXA(() => runJXA<string>(`
      const app = Application("Pages");
      let doc = app.Document();
      app.documents.push(doc);

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
      imageIndex: z.number().int().min(1).describe("Image index (1-based) in the document"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, imageIndex }) => {
      if (!isCreatorStudio()) {
        return toolResult("Super Resolution requires Apple Creator Studio (iWork 15.1+).", true);
      }
      return handleJXA(async () => {
        await runJXA<void>(`
          const app = Application("Pages");
          app.activate();
          const doc = app.documents.byName(params.documentName);
          const images = doc.images();
          if (params.imageIndex > images.length) throw new Error("Image index " + params.imageIndex + " out of range (document has " + images.length + " images)");
          app.selection = [images[params.imageIndex - 1]];
        `, { documentName, imageIndex },
        { label: "super_resolution:select" });

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
      imageIndex: z.number().int().min(1).describe("Image index (1-based) in the document"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, imageIndex }) => {
      if (!isCreatorStudio()) {
        return toolResult("Remove Background requires Apple Creator Studio (iWork 15.1+).", true);
      }
      return handleJXA(async () => {
        await runJXA<void>(`
          const app = Application("Pages");
          app.activate();
          const doc = app.documents.byName(params.documentName);
          const images = doc.images();
          if (params.imageIndex > images.length) throw new Error("Image index " + params.imageIndex + " out of range (document has " + images.length + " images)");
          app.selection = [images[params.imageIndex - 1]];
        `, { documentName, imageIndex },
        { label: "remove_background:select" });

        await clickMenuItem("Pages", ["Format", "Image", "Remove Background"], { postdelay: 2 });

        return JSON.stringify({ success: true, message: "Background removal started on image " + imageIndex + "." });
      });
    },
  );
}
