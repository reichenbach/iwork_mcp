import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runJXA, OsascriptError, isCreatorStudio, creatorStudioSaveAs, creatorStudioExportPDF, saveTimeout, clickMenuItem, resolveAppName } from "../jxa.js";
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

/** Numbers rejects tables smaller than 2x2 ("Invalid row/column count", -10000). */
export const MIN_TABLE_DIMENSION = 2;

/**
 * Table size for a block of data: at least 2x2, wide enough for the longest row,
 * and with at least one body row below the header rows (Numbers rejects
 * headerRowCount >= rowCount with "Invalid header row count").
 */
export function tableDimensions(
  data: readonly (readonly unknown[])[],
  headerRowCount?: number | null,
): { rowCount: number; columnCount: number } {
  const dataCols = data.reduce((max, row) => Math.max(max, row.length), 0);
  return {
    rowCount: Math.max(MIN_TABLE_DIMENSION, data.length, (headerRowCount ?? 0) + 1),
    columnCount: Math.max(MIN_TABLE_DIMENSION, dataCols),
  };
}

export function registerNumbersTools(server: McpServer): void {
  // ── Document Management ──

  server.tool(
    "numbers_list_documents",
    "List all open Numbers documents",
    {},
    ANNOTATIONS.readOnly,
    async () => handleJXA(() => runJXA<string[]>(`
      const app = Application("Numbers");
      const docs = app.documents();
      return JSON.stringify(docs.map(d => ({ name: d.name(), path: d.file() ? d.file().toString() : null })));
    `)),
  );

  server.tool(
    "numbers_list_templates",
    "List all available Numbers templates (e.g. Personal Budget, Invoice, Checklist)",
    {},
    ANNOTATIONS.readOnly,
    async () => handleJXA(() => runJXA<string[]>(`
      const app = Application("Numbers");
      return JSON.stringify(app.templates().map(t => t.name()));
    `)),
  );

  server.tool(
    "numbers_create_document",
    "Create a new Numbers document (optionally from a template — use numbers_list_templates to see available templates)",
    {
      templateName: z.string().optional().describe("Template name (optional)"),
    },
    ANNOTATIONS.readWrite,
    async ({ templateName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      let doc;
      if (params.templateName) {
        doc = app.Document({ documentTemplate: app.templates[params.templateName] });
        app.documents.push(doc);
      } else {
        doc = app.Document();
        app.documents.push(doc);
      }
      return JSON.stringify({ name: doc.name(), sheets: doc.sheets().map(s => s.name()) });
    `, { templateName: templateName ?? null })),
  );

  server.tool(
    "numbers_open_document",
    "Open a .numbers file from disk",
    {
      filePath: z.string().startsWith("/").describe("Absolute path to the .numbers file"),
    },
    ANNOTATIONS.readWrite,
    async ({ filePath }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.open(Path(params.filePath));
      return JSON.stringify({ name: doc.name(), sheets: doc.sheets().map(s => s.name()) });
    `, { filePath })),
  );

  server.tool(
    "numbers_save_document",
    "Save a Numbers document as .numbers (use this to save to disk — use numbers_export_document for PDF/Excel/CSV)",
    {
      documentName: z.string().describe("Name of the open document"),
      filePath: z.string().startsWith("/").optional().describe("File path to save to (for Save As)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, filePath }) => handleJXA(async () => {
      if (isCreatorStudio()) {
        if (filePath) {
          const { name, reopened } = await creatorStudioSaveAs("Numbers", documentName, filePath);
          return JSON.stringify(reopened
            ? { saved: true, name }
            : { saved: true, name, reopened: false, note: `Saved, but Numbers did not reopen the file; open it with numbers_open_document if needed.` });
        }
        return JSON.stringify({ saved: true, name: documentName });
      }
      return runJXA<string>(`
      const app = Application("Numbers");
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
    `, { documentName, filePath: filePath ?? null }, { timeout: saveTimeout(), label: "numbers:save" });
    }),
  );

  server.tool(
    "numbers_export_document",
    "Export a Numbers document to a different format: PDF, Excel (.xlsx), or CSV (not .numbers — use numbers_save_document for that)",
    {
      documentName: z.string().describe("Name of the open document"),
      filePath: z.string().startsWith("/").describe("Absolute path for the exported file"),
      format: z.enum(["PDF", "Excel", "CSV"]).describe("Export format"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, filePath, format }) => handleJXA(async () => {
      if (isCreatorStudio() && format === "PDF") {
        await creatorStudioExportPDF("Numbers", documentName, filePath);
        return JSON.stringify({ exported: true, path: filePath, format: "PDF" });
      }
      return runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const formatMap = {
        "PDF": "PDF",
        "Excel": "Microsoft Excel",
        "CSV": "CSV",
      };
      const fmt = formatMap[params.format];
      app.export(doc, { to: Path(params.filePath), as: fmt });
      return JSON.stringify({ exported: true, path: params.filePath, format: params.format });
    `, { documentName, filePath, format });
    }),
  );

  server.tool(
    "numbers_close_document",
    "Close a Numbers document",
    {
      documentName: z.string().describe("Name of the open document"),
      saving: z.enum(["yes", "no", "ask"]).optional().describe("Whether to save before closing"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, saving }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
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

  // ── Structure Tools ──

  server.tool(
    "numbers_list_sheets",
    "List all sheets in a Numbers document",
    {
      documentName: z.string().describe("Name of the open document"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheets = doc.sheets();
      return JSON.stringify(sheets.map((s, i) => ({
        index: i,
        name: s.name(),
        tableCount: s.tables.length,
      })));
    `, { documentName })),
  );

  server.tool(
    "numbers_add_sheet",
    "Add a new sheet to a Numbers document. By default, deletes the empty 'Table 1' that Numbers auto-creates, so you start with a clean sheet.",
    {
      documentName: z.string().describe("Name of the open document"),
      sheetName: z.string().optional().describe("Name for the new sheet"),
      deleteDefaultTable: z.boolean().optional().describe("Delete the default 'Table 1' that Numbers auto-creates on new sheets (default: true)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, sheetName, deleteDefaultTable }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      // Clean up partial previous attempt (Creator Studio retry safety)
      if (params.sheetName) {
        try { app.delete(doc.sheets.byName(params.sheetName)); } catch(e) {}
      }
      const sheet = app.Sheet();
      doc.sheets.push(sheet);
      if (params.sheetName) {
        sheet.name = params.sheetName;
      }
      if (params.deleteDefaultTable !== false) {
        try { app.delete(sheet.tables[0]); } catch(e) {}
      }
      return JSON.stringify({ name: sheet.name(), index: doc.sheets.length - 1 });
    `, { documentName, sheetName: sheetName ?? null, deleteDefaultTable: deleteDefaultTable ?? true })),
  );

  server.tool(
    "numbers_list_tables",
    "List all tables in a sheet with their dimensions",
    {
      documentName: z.string().describe("Name of the open document"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, sheetName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const tables = sheet.tables();
      return JSON.stringify(tables.map((t, i) => ({
        index: i,
        name: t.name(),
        rowCount: t.rowCount(),
        columnCount: t.columnCount(),
        headerRowCount: t.headerRowCount(),
        headerColumnCount: t.headerColumnCount(),
      })));
    `, { documentName, sheetName: sheetName ?? null })),
  );

  server.tool(
    "numbers_add_table",
    "Create a new table in a sheet",
    {
      documentName: z.string().describe("Name of the open document"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Name for the new table"),
      rows: z.number().int().positive().optional().describe("Number of rows (default: 4)"),
      columns: z.number().int().positive().optional().describe("Number of columns (default: 4)"),
      headerRowCount: z.number().int().min(0).optional().describe("Number of header rows (default: 1, set to 0 to remove header row styling)"),
      headerColumnCount: z.number().int().min(0).optional().describe("Number of header columns (default: 0)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, sheetName, tableName, rows, columns, headerRowCount, headerColumnCount }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const props = {};
      if (params.rows) props.rowCount = Math.max(${MIN_TABLE_DIMENSION}, params.rows);
      if (params.columns) props.columnCount = Math.max(${MIN_TABLE_DIMENSION}, params.columns);
      const table = app.Table(props);
      sheet.tables.push(table);
      if (params.tableName) {
        table.name = params.tableName;
      }
      if (params.headerRowCount !== null && params.headerRowCount !== undefined) {
        table.headerRowCount = params.headerRowCount;
      }
      if (params.headerColumnCount !== null && params.headerColumnCount !== undefined) {
        table.headerColumnCount = params.headerColumnCount;
      }
      return JSON.stringify({
        name: table.name(),
        rowCount: table.rowCount(),
        columnCount: table.columnCount(),
        headerRowCount: table.headerRowCount(),
        headerColumnCount: table.headerColumnCount(),
      });
    `, { documentName, sheetName: sheetName ?? null, tableName: tableName ?? null, rows: rows ?? null, columns: columns ?? null, headerRowCount: headerRowCount ?? null, headerColumnCount: headerColumnCount ?? null })),
  );

  server.tool(
    "numbers_rename_sheet",
    "Rename a sheet in a Numbers document",
    {
      documentName: z.string().describe("Name of the open document"),
      sheetName: z.string().describe("Current sheet name"),
      newName: z.string().describe("New name for the sheet"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, sheetName, newName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = doc.sheets.byName(params.sheetName);
      sheet.name = params.newName;
      return JSON.stringify({ renamed: true, oldName: params.sheetName, newName: params.newName });
    `, { documentName, sheetName, newName })),
  );

  server.tool(
    "numbers_delete_sheet",
    "Delete a sheet from a Numbers document",
    {
      documentName: z.string().describe("Name of the open document"),
      sheetName: z.string().describe("Name of the sheet to delete"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, sheetName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = doc.sheets.byName(params.sheetName);
      app.delete(sheet);
      return JSON.stringify({ deleted: true, sheetName: params.sheetName });
    `, { documentName, sheetName })),
  );

  server.tool(
    "numbers_rename_table",
    "Rename a table in a Numbers document",
    {
      documentName: z.string().describe("Name of the open document"),
      tableName: z.string().describe("Current table name"),
      newName: z.string().describe("New name for the table"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, tableName, newName, sheetName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = sheet.tables.byName(params.tableName);
      table.name = params.newName;
      return JSON.stringify({ renamed: true, oldName: params.tableName, newName: params.newName });
    `, { documentName, tableName, newName, sheetName: sheetName ?? null })),
  );

  server.tool(
    "numbers_delete_table",
    "Delete a table from a sheet",
    {
      documentName: z.string().describe("Name of the open document"),
      tableName: z.string().describe("Name of the table to delete"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, tableName, sheetName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = sheet.tables.byName(params.tableName);
      app.delete(table);
      return JSON.stringify({ deleted: true, tableName: params.tableName });
    `, { documentName, tableName, sheetName: sheetName ?? null })),
  );

  server.tool(
    "numbers_delete_row",
    "Delete one or more rows from a table",
    {
      documentName: z.string().describe("Name of the open document"),
      rowIndex: z.number().int().min(1).describe("Row number to delete (1-based)"),
      count: z.number().int().positive().optional().describe("Number of rows to delete (default: 1)"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, rowIndex, count, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      const n = params.count || 1;
      for (let i = 0; i < n; i++) {
        app.delete(table.rows[params.rowIndex - 1]);
      }
      return JSON.stringify({ deleted: n, newRowCount: table.rowCount() });
    `, { documentName, rowIndex, count: count ?? null, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_delete_column",
    "Delete one or more columns from a table",
    {
      documentName: z.string().describe("Name of the open document"),
      column: z.string().regex(/^[A-Z]{1,3}$/).describe("Column letter to delete, e.g. 'A', 'B'"),
      count: z.number().int().positive().optional().describe("Number of columns to delete (default: 1)"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, column, count, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      const colStr = params.column.toUpperCase();
      let colIndex = 0;
      for (let i = 0; i < colStr.length; i++) {
        colIndex = colIndex * 26 + (colStr.charCodeAt(i) - 64);
      }
      colIndex -= 1;
      const n = params.count || 1;
      for (let i = 0; i < n; i++) {
        app.delete(table.columns[colIndex]);
      }
      return JSON.stringify({ deleted: n, newColumnCount: table.columnCount() });
    `, { documentName, column, count: count ?? null, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  // ── Data Reading Tools ──

  server.tool(
    "numbers_read_table",
    "Read all data from a table as a 2D array. Returns rows of cell values; empty rows/columns at the end are left out unless includeEmpty is set.",
    {
      documentName: z.string().describe("Name of the open document"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
      includeEmpty: z.boolean().optional().describe("Keep empty rows/columns at the end of the table (default: false)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, sheetName, tableName, includeEmpty }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      const rowCount = table.rowCount();
      const colCount = table.columnCount();
      // Batch read all cell values in one IPC call
      const allValues = table.cells.value();
      let data = [];
      for (let r = 0; r < rowCount; r++) {
        data.push(allValues.slice(r * colCount, (r + 1) * colCount));
      }
      if (!params.includeEmpty) {
        const isEmpty = (v) => v === null || v === undefined || v === "";
        while (data.length > 0 && data[data.length - 1].every(isEmpty)) data.pop();
        let cols = colCount;
        while (cols > 0 && data.every(row => isEmpty(row[cols - 1]))) cols--;
        data = data.map(row => row.slice(0, cols));
      }
      return JSON.stringify(data);
    `, { documentName, sheetName: sheetName ?? null, tableName: tableName ?? null, includeEmpty: includeEmpty ?? false })),
  );

  server.tool(
    "numbers_read_cell",
    "Read a single cell's value",
    {
      documentName: z.string().describe("Name of the open document"),
      cellRef: z.string().regex(/^[A-Z]{1,3}\d+$/).describe("Cell reference, e.g. 'A1', 'B3'"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, cellRef, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      const cell = table.cells[params.cellRef];
      return JSON.stringify({ cellRef: params.cellRef, value: cell.value(), formattedValue: cell.formattedValue() });
    `, { documentName, cellRef, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_get_table_info",
    "Get table metadata: row/column counts, header info, table name, and position",
    {
      documentName: z.string().describe("Name of the open document"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      var result = {
        name: table.name(),
        rowCount: table.rowCount(),
        columnCount: table.columnCount(),
        headerRowCount: table.headerRowCount(),
        headerColumnCount: table.headerColumnCount(),
        footerRowCount: table.footerRowCount(),
      };
      try { result.position = table.position(); } catch(e) {}
      return JSON.stringify(result);
    `, { documentName, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_get_cell_format",
    "Read the number format and alignment of a cell (e.g. percent, currency, number, automatic)",
    {
      documentName: z.string().describe("Name of the open document"),
      cellRef: z.string().regex(/^[A-Z]{1,3}\d+$/).describe("Cell reference, e.g. 'A1', 'B3'"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, cellRef, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      const cell = table.cells[params.cellRef];
      var result = { cellRef: params.cellRef };
      try { result.format = cell.format() || "automatic"; } catch(e) { result.format = "automatic"; }
      try { result.alignment = cell.alignment(); } catch(e) {}
      try { result.verticalAlignment = cell.verticalAlignment(); } catch(e) {}
      return JSON.stringify(result);
    `, { documentName, cellRef, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_move_table",
    "Move a table to a new position on the sheet",
    {
      documentName: z.string().describe("Name of the open document"),
      x: z.number().describe("X position in points"),
      y: z.number().describe("Y position in points"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, x, y, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      table.position = { x: params.x, y: params.y };
      const pos = table.position();
      return JSON.stringify({ moved: true, name: table.name(), position: pos });
    `, { documentName, x, y, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  // ── Data Writing Tools ──

  server.tool(
    "numbers_write_cell",
    "Write a value to a cell",
    {
      documentName: z.string().describe("Name of the open document"),
      cellRef: z.string().regex(/^[A-Z]{1,3}\d+$/).describe("Cell reference, e.g. 'A1', 'B3'"),
      value: z.union([z.string(), z.number(), z.boolean()]).describe("Value to write"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, cellRef, value, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      table.cells[params.cellRef].value = params.value;
      return JSON.stringify({ cellRef: params.cellRef, written: true });
    `, { documentName, cellRef, value, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_write_cells",
    "Batch write multiple cell values in a single operation",
    {
      documentName: z.string().describe("Name of the open document"),
      writes: z.array(z.object({
        cellRef: z.string().regex(/^[A-Z]{1,3}\d+$/).describe("Cell reference, e.g. 'A1'"),
        value: z.union([z.string(), z.number(), z.boolean()]).describe("Value to write"),
      })).describe("Array of cell writes"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, writes, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      const results = [];
      for (const w of params.writes) {
        table.cells[w.cellRef].value = w.value;
        results.push({ cellRef: w.cellRef, written: true });
      }
      return JSON.stringify(results);
    `, { documentName, writes, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_write_table",
    "Bulk write an entire table of data in a single operation. Much faster than writing cells individually. Resizes the table to fit the data automatically, including removing empty rows/columns past the data.",
    {
      documentName: z.string().describe("Name of the open document"),
      data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).min(1)
        .describe("2D array of data — first row can be headers. Null cells are skipped."),
      startCell: z.string().regex(/^[A-Z]{1,3}\d+$/).optional().describe("Top-left cell to start writing from (default: 'A1')"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
      resizeToFit: z.boolean().optional().describe("Resize table to fit data: add missing rows/columns and remove empty ones past the data (default: true)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, data, startCell, sheetName, tableName, resizeToFit }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];

      // Parse start cell (e.g. "B3" -> col=1, row=2)
      const cellRef = (params.startCell || "A1").toUpperCase();
      const colMatch = cellRef.match(/^([A-Z]+)/);
      const rowMatch = cellRef.match(/(\\d+)$/);
      let startCol = 0;
      if (colMatch) {
        for (let i = 0; i < colMatch[1].length; i++) {
          startCol = startCol * 26 + (colMatch[1].charCodeAt(i) - 64);
        }
        startCol -= 1;
      }
      const startRow = rowMatch ? parseInt(rowMatch[1]) - 1 : 0;

      const dataRows = params.data.length;
      const dataCols = Math.max(...params.data.map(r => r.length));
      const needRows = startRow + dataRows;
      const needCols = startCol + dataCols;

      // Resize table if needed
      if (params.resizeToFit !== false) {
        while (table.rowCount() < needRows) {
          table.rows.push(app.Row());
        }
        while (table.columnCount() < needCols) {
          table.columns.push(app.Column());
        }
      }

      const colCount = table.columnCount();
      let cellsWritten = 0;
      for (let r = 0; r < dataRows; r++) {
        for (let c = 0; c < params.data[r].length; c++) {
          const val = params.data[r][c];
          if (val !== null) {
            table.cells[(startRow + r) * colCount + (startCol + c)].value = val;
            cellsWritten++;
          }
        }
      }

      // Also drop trailing rows/columns past the data if they are empty (e.g. the 22x7 default table)
      if (params.resizeToFit !== false) {
        const values = table.cells.value();
        const cols = table.columnCount();
        const isEmpty = (v) => v === null || v === undefined || v === "";
        const rowEmpty = (r) => { for (let c = 0; c < cols; c++) if (!isEmpty(values[r * cols + c])) return false; return true; };
        const colEmpty = (c, rows) => { for (let r = 0; r < rows; r++) if (!isEmpty(values[r * cols + c])) return false; return true; };
        try {
          for (let r = table.rowCount() - 1; r >= Math.max(needRows, 2) && rowEmpty(r); r--) app.delete(table.rows[r]);
          const rows = table.rowCount();
          for (let c = table.columnCount() - 1; c >= Math.max(needCols, 2) && colEmpty(c, rows); c--) app.delete(table.columns[c]);
        } catch (e) {}
      }
      return JSON.stringify({ cellsWritten: cellsWritten, rows: dataRows, columns: dataCols, tableRows: table.rowCount(), tableColumns: table.columnCount() });
    `, { documentName, data, startCell: startCell ?? null, sheetName: sheetName ?? null, tableName: tableName ?? null, resizeToFit: resizeToFit ?? true })),
  );

  server.tool(
    "numbers_set_formula",
    "Set a formula on a cell (e.g. '=SUM(A1:A10)')",
    {
      documentName: z.string().describe("Name of the open document"),
      cellRef: z.string().regex(/^[A-Z]{1,3}\d+$/).describe("Cell reference, e.g. 'A1'"),
      formula: z.string().describe("Formula string (e.g. '=SUM(A1:A10)')"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, cellRef, formula, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      table.cells[params.cellRef].value = params.formula;
      return JSON.stringify({ cellRef: params.cellRef, formula: params.formula, set: true });
    `, { documentName, cellRef, formula, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_add_row",
    "Add one or more rows at the end of a table, optionally with data",
    {
      documentName: z.string().describe("Name of the open document"),
      data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).optional()
        .describe("2D array of row data to populate (each inner array is one row)"),
      position: z.enum(["end", "beginning"]).optional().describe("Where to add the row (default: end)"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, data, position, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];

      const rowsToAdd = params.data ? params.data.length : 1;
      const colCount = table.columnCount();
      const existingRows = table.rowCount();

      for (let i = 0; i < rowsToAdd; i++) {
        table.rows.push(app.Row());
      }

      if (params.position === "beginning" && existingRows > 0) {
        // Shift existing data down to make room at the top
        for (let r = existingRows - 1; r >= 0; r--) {
          for (let c = 0; c < colCount; c++) {
            const val = table.cells[r * colCount + c].value();
            table.cells[(r + rowsToAdd) * colCount + c].value = val;
            if (val !== null && val !== undefined && val !== "") {
              table.cells[r * colCount + c].value = null;
            }
          }
        }
      }

      if (params.data) {
        const startRow = params.position === "beginning" ? 0 : existingRows;
        for (let r = 0; r < params.data.length; r++) {
          for (let c = 0; c < Math.min(params.data[r].length, colCount); c++) {
            if (params.data[r][c] !== null) {
              table.cells[(startRow + r) * colCount + c].value = params.data[r][c];
            }
          }
        }
      }

      return JSON.stringify({ rowsAdded: rowsToAdd, newRowCount: table.rowCount() });
    `, { documentName, data: data ?? null, position: position ?? null, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_insert_row_at",
    "Insert one or more rows at a specific position in a table, shifting existing rows down",
    {
      documentName: z.string().describe("Name of the open document"),
      rowIndex: z.number().int().min(1).describe("Insert before this row (1-based)"),
      data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).optional()
        .describe("2D array of row data to populate (each inner array is one row)"),
      count: z.number().int().positive().optional().describe("Number of rows to insert (default: 1, ignored if data provided)"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, rowIndex, data, count, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];

      const rowsToAdd = params.data ? params.data.length : (params.count || 1);
      const colCount = table.columnCount();
      const existingRows = table.rowCount();
      const insertAt = params.rowIndex - 1; // 0-based

      // Append empty rows
      for (let i = 0; i < rowsToAdd; i++) {
        table.rows.push(app.Row());
      }

      // Shift existing rows down from insertAt
      for (let r = existingRows - 1; r >= insertAt; r--) {
        for (let c = 0; c < colCount; c++) {
          const val = table.cells[r * colCount + c].value();
          table.cells[(r + rowsToAdd) * colCount + c].value = val;
          table.cells[r * colCount + c].value = null;
        }
      }

      // Write data at insertAt if provided
      if (params.data) {
        for (let r = 0; r < params.data.length; r++) {
          for (let c = 0; c < Math.min(params.data[r].length, colCount); c++) {
            if (params.data[r][c] !== null) {
              table.cells[(insertAt + r) * colCount + c].value = params.data[r][c];
            }
          }
        }
      }

      return JSON.stringify({ rowsInserted: rowsToAdd, newRowCount: table.rowCount() });
    `, { documentName, rowIndex, data: data ?? null, count: count ?? null, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_insert_column_at",
    "Insert one or more columns at a specific position in a table, shifting existing columns right",
    {
      documentName: z.string().describe("Name of the open document"),
      column: z.string().regex(/^[A-Z]{1,3}$/).describe("Insert before this column letter, e.g. 'B'"),
      data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).optional()
        .describe("Column-major 2D array: data[i] = values for column i, top-to-bottom"),
      count: z.number().int().positive().optional().describe("Number of columns to insert (default: 1, ignored if data provided)"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, column, data, count, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];

      // Convert column letter to 0-based index
      const colStr = params.column.toUpperCase();
      let insertAt = 0;
      for (let i = 0; i < colStr.length; i++) {
        insertAt = insertAt * 26 + (colStr.charCodeAt(i) - 64);
      }
      insertAt -= 1;

      const colsToAdd = params.data ? params.data.length : (params.count || 1);
      const existingCols = table.columnCount();
      const rowCount = table.rowCount();

      // Append empty columns
      for (let i = 0; i < colsToAdd; i++) {
        table.columns.push(app.Column());
      }
      const newTotalCols = table.columnCount();

      // Shift existing columns right: for each row, from rightmost existing col down to insertAt
      for (let r = 0; r < rowCount; r++) {
        for (let c = existingCols - 1; c >= insertAt; c--) {
          const val = table.cells[r * newTotalCols + c].value();
          table.cells[r * newTotalCols + (c + colsToAdd)].value = val;
          table.cells[r * newTotalCols + c].value = null;
        }
      }

      // Write data if provided (column-major: data[i][r] -> row r, column insertAt+i)
      if (params.data) {
        for (let i = 0; i < params.data.length; i++) {
          for (let r = 0; r < params.data[i].length; r++) {
            if (params.data[i][r] !== null) {
              table.cells[r * newTotalCols + (insertAt + i)].value = params.data[i][r];
            }
          }
        }
      }

      return JSON.stringify({ columnsInserted: colsToAdd, newColumnCount: table.columnCount() });
    `, { documentName, column, data: data ?? null, count: count ?? null, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_add_column",
    "Add a column to a table",
    {
      documentName: z.string().describe("Name of the open document"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      table.columns.push(app.Column());
      return JSON.stringify({ newColumnCount: table.columnCount() });
    `, { documentName, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_read_range",
    "Read a specific cell range (e.g. 'B2:D10') instead of the entire table. Faster for large tables.",
    {
      documentName: z.string().describe("Name of the open document"),
      cellRange: z.string().regex(/^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/).describe("Cell range, e.g. 'A1:C10'"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, cellRange, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];

      // Parse cell range "B2:C3" into row/col indices
      function parseRef(ref) {
        const m = ref.match(/^([A-Z]+)(\\d+)$/);
        let col = 0;
        for (let i = 0; i < m[1].length; i++) col = col * 26 + m[1].charCodeAt(i) - 64;
        return { row: parseInt(m[2], 10) - 1, col: col - 1 };
      }
      const parts = params.cellRange.split(":");
      const start = parseRef(parts[0]);
      const end = parseRef(parts[1]);
      const totalCols = table.columnCount();
      const data = [];
      for (let r = start.row; r <= end.row; r++) {
        const row = [];
        for (let c = start.col; c <= end.col; c++) {
          row.push(table.cells[r * totalCols + c].value());
        }
        data.push(row);
      }
      return JSON.stringify(data);
    `, { documentName, cellRange, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_merge_cells",
    "Merge a range of cells",
    {
      documentName: z.string().describe("Name of the open document"),
      cellRange: z.string().regex(/^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/).describe("Cell range to merge, e.g. 'A1:C1'"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, cellRange, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      // Merging fails if range spans header/non-header boundaries — clear headers first
      if (table.headerRowCount() > 0) table.headerRowCount = 0;
      if (table.headerColumnCount() > 0) table.headerColumnCount = 0;
      const range = table.ranges[params.cellRange];
      range.merge();
      return JSON.stringify({ merged: true, cellRange: params.cellRange });
    `, { documentName, cellRange, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_unmerge_cells",
    "Unmerge a previously merged cell range",
    {
      documentName: z.string().describe("Name of the open document"),
      cellRange: z.string().regex(/^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/).describe("Cell range to unmerge, e.g. 'A1:C1'"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, cellRange, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      const range = table.ranges[params.cellRange];
      range.unmerge();
      return JSON.stringify({ unmerged: true, cellRange: params.cellRange });
    `, { documentName, cellRange, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_clear_cells",
    "Clear the contents of a cell or range",
    {
      documentName: z.string().describe("Name of the open document"),
      cellRange: z.string().regex(/^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$/).describe("Cell or range to clear, e.g. 'A1' or 'A1:C10'"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, cellRange, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      if (params.cellRange.includes(":")) {
        const range = table.ranges[params.cellRange];
        const cells = range.cells();
        for (const cell of cells) { cell.value = null; }
      } else {
        table.cells[params.cellRange].value = null;
      }
      return JSON.stringify({ cleared: true, cellRange: params.cellRange });
    `, { documentName, cellRange, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_sort_rows",
    "Sort table rows by a column",
    {
      documentName: z.string().describe("Name of the open document"),
      column: z.string().regex(/^[A-Z]{1,3}$/).describe("Column letter to sort by, e.g. 'A'"),
      order: z.enum(["ascending", "descending"]).optional().describe("Sort order (default: ascending)"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, column, order, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      const colStr = params.column.toUpperCase();
      let colIndex = 0;
      for (let i = 0; i < colStr.length; i++) {
        colIndex = colIndex * 26 + (colStr.charCodeAt(i) - 64);
      }
      colIndex -= 1;
      const col = table.columns[colIndex];
      const direction = params.order === "descending" ? "descending" : "ascending";
      table.sort({ by: col, direction: direction });
      return JSON.stringify({ sorted: true, column: params.column, order: direction });
    `, { documentName, column, order: order ?? null, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  // ── Table Configuration Tools ──

  server.tool(
    "numbers_set_header_rows",
    "Set the number of header rows on a table. Header rows get special styling (bold text, grey background). Set to 0 to remove header styling entirely.",
    {
      documentName: z.string().describe("Name of the open document"),
      headerRowCount: z.number().int().min(0).describe("Number of header rows (0 removes all header styling, 1 is default)"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, headerRowCount, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      table.headerRowCount = params.headerRowCount;
      return JSON.stringify({ headerRowCount: table.headerRowCount(), tableName: table.name() });
    `, { documentName, headerRowCount, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_set_header_columns",
    "Set the number of header columns on a table. Header columns get special styling. Set to 0 to remove header column styling.",
    {
      documentName: z.string().describe("Name of the open document"),
      headerColumnCount: z.number().int().min(0).describe("Number of header columns (0 removes header column styling)"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, headerColumnCount, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      table.headerColumnCount = params.headerColumnCount;
      return JSON.stringify({ headerColumnCount: table.headerColumnCount(), tableName: table.name() });
    `, { documentName, headerColumnCount, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  // ── Formatting Tools ──

  server.tool(
    "numbers_format_cells",
    "Set formatting on a cell or range: font, size, color, alignment, background color, bold, italic",
    {
      documentName: z.string().describe("Name of the open document"),
      cellRange: z.string().regex(/^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$/).describe("Cell or range reference, e.g. 'A1' or 'A1:C3'"),
      format: z.object({
        bold: z.boolean().optional().describe("Set bold (switches to bold variant of current font)"),
        italic: z.boolean().optional().describe("Set italic (switches to italic variant of current font)"),
        fontSize: z.number().positive().optional().describe("Font size in points"),
        fontName: z.string().optional().describe("Font name"),
        textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Text color as hex, e.g. '#FF0000'"),
        backgroundColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Background color as hex, e.g. '#0000FF'"),
        alignment: z.enum(["left", "center", "right", "auto"]).optional().describe("Text alignment"),
        verticalAlignment: z.enum(["top", "center", "bottom"]).optional().describe("Vertical alignment"),
        textWrap: z.boolean().optional().describe("Enable text wrapping"),
        numberFormat: z.enum(["automatic", "number", "currency", "percent", "fraction", "scientific", "text", "checkbox", "star rating"]).optional().describe("Number format for the cell"),
      }).describe("Formatting options"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, cellRange, format, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      const fmt = params.format;

      const rangeStr = params.cellRange;
      let cells = [];
      if (rangeStr.includes(":")) {
        const range = table.ranges[rangeStr];
        cells = range.cells();
      } else {
        cells = [table.cells[rangeStr]];
      }

      function hexToRGB(hex) {
        const r = parseInt(hex.slice(1, 3), 16) * 257;
        const g = parseInt(hex.slice(3, 5), 16) * 257;
        const b = parseInt(hex.slice(5, 7), 16) * 257;
        return [r, g, b];
      }

      for (const cell of cells) {
        if (fmt.fontSize !== undefined) cell.fontSize = fmt.fontSize;
        if (fmt.fontName !== undefined) cell.fontName = fmt.fontName;
        if (fmt.textColor !== undefined) {
          cell.textColor = hexToRGB(fmt.textColor);
        }
        if (fmt.backgroundColor !== undefined) {
          cell.backgroundColor = hexToRGB(fmt.backgroundColor);
        }
        if (fmt.alignment !== undefined) {
          cell.alignment = fmt.alignment;
        }
        if (fmt.verticalAlignment !== undefined) {
          cell.verticalAlignment = fmt.verticalAlignment;
        }
        if (fmt.textWrap !== undefined) {
          cell.textWrap = fmt.textWrap;
        }
        if (fmt.numberFormat !== undefined) {
          cell.format = fmt.numberFormat;
        }
        // Bold/italic: switch font to bold/italic variant (PostScript names)
        if (fmt.bold !== undefined || fmt.italic !== undefined) {
          let fontName = fmt.fontName || cell.fontName();
          const baseName = fontName.replace(/[- ]?(Bold ?Italic|BoldItalic|Bold|Italic)$/i, "").trim();
          const wantBold = fmt.bold !== undefined ? fmt.bold : /Bold/i.test(fontName);
          const wantItalic = fmt.italic !== undefined ? fmt.italic : /Italic/i.test(fontName);
          let suffix = "";
          if (wantBold && wantItalic) suffix = "-BoldItalic";
          else if (wantBold) suffix = "-Bold";
          else if (wantItalic) suffix = "-Italic";
          cell.fontName = baseName + suffix;
        }
      }

      return JSON.stringify({ formatted: true, cellRange: params.cellRange });
    `, { documentName, cellRange, format, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_set_column_width",
    "Set the width of a column",
    {
      documentName: z.string().describe("Name of the open document"),
      column: z.string().regex(/^[A-Z]{1,3}$/).describe("Column letter, e.g. 'A', 'B'"),
      width: z.number().positive().describe("Width in points"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, column, width, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];

      // Convert column letter to index (A=0, B=1, ...)
      const colStr = params.column.toUpperCase();
      let colIndex = 0;
      for (let i = 0; i < colStr.length; i++) {
        colIndex = colIndex * 26 + (colStr.charCodeAt(i) - 64);
      }
      colIndex -= 1;

      table.columns[colIndex].width = params.width;
      return JSON.stringify({ column: params.column, width: params.width, set: true });
    `, { documentName, column, width, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  server.tool(
    "numbers_set_row_height",
    "Set the height of a row",
    {
      documentName: z.string().describe("Name of the open document"),
      row: z.number().int().min(1).describe("Row number (1-based)"),
      height: z.number().positive().describe("Height in points"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, row, height, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
      table.rows[params.row - 1].height = params.height;
      return JSON.stringify({ row: params.row, height: params.height, set: true });
    `, { documentName, row, height, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  // ── Compound / Batch Tools ──

  server.tool(
    "numbers_create_sheet_with_table",
    "Create a complete sheet with a named table, data, and formatting in ONE operation. Much faster than calling individual tools. Use this for bulk setup like calendars, reports, dashboards.",
    {
      documentName: z.string().describe("Name of the open document"),
      sheetName: z.string().describe("Name for the new sheet"),
      tableName: z.string().optional().describe("Name for the table (default: same as sheet name)"),
      data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).min(1)
        .describe("2D array of table data. First row can be headers."),
      headerRowCount: z.number().int().min(0).optional().describe("Number of header rows (default: 0 for no header styling)"),
      columnWidths: z.array(z.number().positive()).optional().describe("Width in points for each column"),
      rowHeights: z.array(z.number().positive()).optional().describe("Height in points for each row"),
      formatting: z.array(z.object({
        cellRange: z.string().regex(/^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$/).describe("Cell or range, e.g. 'A1' or 'A1:G1'"),
        bold: z.boolean().optional(),
        italic: z.boolean().optional(),
        fontSize: z.number().positive().optional(),
        fontName: z.string().optional(),
        textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Hex color, e.g. '#FF0000'"),
        backgroundColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Hex color, e.g. '#0000FF'"),
        alignment: z.enum(["left", "center", "right", "auto"]).optional(),
        verticalAlignment: z.enum(["top", "center", "bottom"]).optional(),
        textWrap: z.boolean().optional(),
      })).optional().describe("Array of formatting rules to apply"),
      deleteDefaultTable: z.boolean().optional().describe("Delete the default 'Table 1' that Numbers auto-creates (default: true)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, sheetName, tableName, data, headerRowCount, columnWidths, rowHeights, formatting, deleteDefaultTable }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);

      // Clean up partial previous attempt (Creator Studio retry safety)
      try { app.delete(doc.sheets.byName(params.sheetName)); } catch(e) {}

      // Create the sheet
      const sheet = app.Sheet();
      doc.sheets.push(sheet);
      sheet.name = params.sheetName;

      // Delete the default "Table 1" that Numbers auto-creates
      if (params.deleteDefaultTable !== false) {
        try { app.delete(sheet.tables[0]); } catch(e) {}
      }

      // Create our table (dimensions pre-clamped to the 2x2 minimum in tableDimensions())
      const dataRows = params.data.length;
      const table = app.Table({ rowCount: params.rowCount, columnCount: params.columnCount });
      sheet.tables.push(table);
      table.name = params.tableName || params.sheetName;

      // Set header counts (default 0 = no header styling, allows merging across all cells)
      table.headerRowCount = (params.headerRowCount !== null && params.headerRowCount !== undefined) ? params.headerRowCount : 0;
      table.headerColumnCount = 0;

      // Write all data
      const colCount = table.columnCount();
      for (let r = 0; r < dataRows; r++) {
        for (let c = 0; c < params.data[r].length; c++) {
          const val = params.data[r][c];
          if (val !== null) {
            table.cells[r * colCount + c].value = val;
          }
        }
      }

      // Set column widths
      if (params.columnWidths) {
        for (let i = 0; i < Math.min(params.columnWidths.length, colCount); i++) {
          table.columns[i].width = params.columnWidths[i];
        }
      }

      // Set row heights
      if (params.rowHeights) {
        for (let i = 0; i < Math.min(params.rowHeights.length, dataRows); i++) {
          table.rows[i].height = params.rowHeights[i];
        }
      }

      // Apply formatting
      function hexToRGB(hex) {
        return [parseInt(hex.slice(1,3),16)*257, parseInt(hex.slice(3,5),16)*257, parseInt(hex.slice(5,7),16)*257];
      }

      if (params.formatting) {
        for (const fmt of params.formatting) {
          let cells = [];
          if (fmt.cellRange.includes(":")) {
            cells = table.ranges[fmt.cellRange].cells();
          } else {
            cells = [table.cells[fmt.cellRange]];
          }
          for (const cell of cells) {
            if (fmt.fontSize !== undefined) cell.fontSize = fmt.fontSize;
            if (fmt.fontName !== undefined) cell.fontName = fmt.fontName;
            if (fmt.textColor !== undefined) cell.textColor = hexToRGB(fmt.textColor);
            if (fmt.backgroundColor !== undefined) cell.backgroundColor = hexToRGB(fmt.backgroundColor);
            if (fmt.alignment !== undefined) cell.alignment = fmt.alignment;
            if (fmt.verticalAlignment !== undefined) cell.verticalAlignment = fmt.verticalAlignment;
            if (fmt.textWrap !== undefined) cell.textWrap = fmt.textWrap;
            if (fmt.bold !== undefined || fmt.italic !== undefined) {
              let fontName = fmt.fontName || cell.fontName();
              const baseName = fontName.replace(/[- ]?(Bold ?Italic|BoldItalic|Bold|Italic)$/i, "").trim();
              const wantBold = fmt.bold !== undefined ? fmt.bold : /Bold/i.test(fontName);
              const wantItalic = fmt.italic !== undefined ? fmt.italic : /Italic/i.test(fontName);
              let suffix = "";
              if (wantBold && wantItalic) suffix = "-BoldItalic";
              else if (wantBold) suffix = "-Bold";
              else if (wantItalic) suffix = "-Italic";
              cell.fontName = baseName + suffix;
            }
          }
        }
      }

      return JSON.stringify({
        sheetName: sheet.name(),
        tableName: table.name(),
        rows: table.rowCount(),
        columns: table.columnCount(),
        headerRowCount: table.headerRowCount(),
      });
    `, {
      documentName, sheetName,
      tableName: tableName ?? null,
      data,
      ...tableDimensions(data, headerRowCount),
      headerRowCount: headerRowCount ?? null,
      columnWidths: columnWidths ?? null,
      rowHeights: rowHeights ?? null,
      formatting: formatting ?? null,
      deleteDefaultTable: deleteDefaultTable ?? true,
    })),
  );

  // ── Batch Format Tool ──

  server.tool(
    "numbers_format_range",
    "Apply multiple formatting rules to different cell ranges in one call. More efficient than calling numbers_format_cells repeatedly.",
    {
      documentName: z.string().describe("Name of the open document"),
      rules: z.array(z.object({
        cellRange: z.string().regex(/^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$/).describe("Cell or range reference, e.g. 'A1' or 'A1:C3'"),
        bold: z.boolean().optional().describe("Set bold"),
        italic: z.boolean().optional().describe("Set italic"),
        fontSize: z.number().positive().optional().describe("Font size in points"),
        fontName: z.string().optional().describe("Font name (PostScript name)"),
        textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Text color as hex, e.g. '#FF0000'"),
        backgroundColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Background color as hex"),
        alignment: z.enum(["left", "center", "right", "auto"]).optional().describe("Text alignment"),
        verticalAlignment: z.enum(["top", "center", "bottom"]).optional().describe("Vertical alignment"),
        textWrap: z.boolean().optional().describe("Enable text wrapping"),
        numberFormat: z.enum(["automatic", "number", "currency", "percent", "fraction", "scientific", "text", "checkbox", "star rating"]).optional().describe("Number format"),
      })).min(1).describe("Array of formatting rules, each targeting a cell range"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, rules, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];

      function hexToRGB(hex) {
        const r = parseInt(hex.slice(1, 3), 16) * 257;
        const g = parseInt(hex.slice(3, 5), 16) * 257;
        const b = parseInt(hex.slice(5, 7), 16) * 257;
        return [r, g, b];
      }

      let rulesApplied = 0;
      let cellsFormatted = 0;

      for (const fmt of params.rules) {
        const rangeStr = fmt.cellRange;
        let cells = [];
        if (rangeStr.includes(":")) {
          const range = table.ranges[rangeStr];
          cells = range.cells();
        } else {
          cells = [table.cells[rangeStr]];
        }

        for (const cell of cells) {
          if (fmt.fontSize !== undefined) cell.fontSize = fmt.fontSize;
          if (fmt.fontName !== undefined) cell.fontName = fmt.fontName;
          if (fmt.textColor !== undefined) cell.textColor = hexToRGB(fmt.textColor);
          if (fmt.backgroundColor !== undefined) cell.backgroundColor = hexToRGB(fmt.backgroundColor);
          if (fmt.alignment !== undefined) cell.alignment = fmt.alignment;
          if (fmt.verticalAlignment !== undefined) cell.verticalAlignment = fmt.verticalAlignment;
          if (fmt.textWrap !== undefined) cell.textWrap = fmt.textWrap;
          if (fmt.numberFormat !== undefined) cell.format = fmt.numberFormat;
          if (fmt.bold !== undefined || fmt.italic !== undefined) {
            let fontName = fmt.fontName || cell.fontName();
            const baseName = fontName.replace(/[- ]?(Bold ?Italic|BoldItalic|Bold|Italic)$/i, "").trim();
            const wantBold = fmt.bold !== undefined ? fmt.bold : /Bold/i.test(fontName);
            const wantItalic = fmt.italic !== undefined ? fmt.italic : /Italic/i.test(fontName);
            let suffix = "";
            if (wantBold && wantItalic) suffix = "-BoldItalic";
            else if (wantBold) suffix = "-Bold";
            else if (wantItalic) suffix = "-Italic";
            cell.fontName = baseName + suffix;
          }
          cellsFormatted++;
        }
        rulesApplied++;
      }

      return JSON.stringify({ rulesApplied, cellsFormatted });
    `, {
      documentName,
      rules,
      sheetName: sheetName ?? null,
      tableName: tableName ?? null,
    })),
  );

  // ── Auto-Format Tool ──

  server.tool(
    "numbers_auto_format",
    "Scan a cell range, detect formats from string values (currency, percent, numbers with commas), convert to numeric values and apply native formatting. For example: '$1,234.56' becomes 1234.56 with currency format, '45%' becomes 0.45 with percent format.",
    {
      documentName: z.string().describe("Name of the open document"),
      cellRange: z.string().regex(/^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/).describe("Cell range to auto-format, e.g. 'A1:C10'"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, cellRange, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];

      function parseRef(ref) {
        const m = ref.match(/^([A-Z]+)(\\d+)$/);
        let col = 0;
        for (let i = 0; i < m[1].length; i++) col = col * 26 + m[1].charCodeAt(i) - 64;
        return { row: parseInt(m[2], 10) - 1, col: col - 1 };
      }
      const parts = params.cellRange.split(":");
      const start = parseRef(parts[0]);
      const end = parseRef(parts[1]);
      const totalCols = table.columnCount();

      const stats = { currency: 0, percent: 0, number: 0, text: 0, empty: 0 };

      for (let r = start.row; r <= end.row; r++) {
        for (let c = start.col; c <= end.col; c++) {
          const cell = table.cells[r * totalCols + c];
          const val = cell.value();
          if (val === null || val === undefined || val === "") {
            stats.empty++;
            continue;
          }
          if (typeof val !== "string") {
            stats.text++;
            continue;
          }
          const s = val.trim();
          // Currency: $1,234.56 or ($1,234.56) for negative
          const currMatch = s.match(/^\\(?\\$([\\d,]+\\.?\\d*)\\)?$/);
          if (currMatch) {
            let num = parseFloat(currMatch[1].replace(/,/g, ""));
            if (s.startsWith("(")) num = -num;
            cell.value = num;
            cell.format = "currency";
            stats.currency++;
            continue;
          }
          // Percent: 45% or 45.5%
          const pctMatch = s.match(/^(-?[\\d.]+)%$/);
          if (pctMatch) {
            cell.value = parseFloat(pctMatch[1]) / 100;
            cell.format = "percent";
            stats.percent++;
            continue;
          }
          // Number with commas: 1,234 or 1,234.56
          const numMatch = s.match(/^-?[\\d,]+\\.?\\d*$/);
          if (numMatch && s.includes(",")) {
            cell.value = parseFloat(s.replace(/,/g, ""));
            cell.format = "number";
            stats.number++;
            continue;
          }
          stats.text++;
        }
      }

      return JSON.stringify({ autoFormatted: true, stats });
    `, { documentName, cellRange, sheetName: sheetName ?? null, tableName: tableName ?? null })),
  );

  // ── Cross-Document Copy Tool ──

  server.tool(
    "numbers_copy_range",
    "Copy values from one document/sheet/table range to another. Both documents must be open. Copies values only (not formulas or formatting).",
    {
      sourceDoc: z.string().describe("Name of the source document"),
      sourceRange: z.string().regex(/^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/).describe("Source cell range, e.g. 'A1:C10'"),
      destDoc: z.string().describe("Name of the destination document"),
      destCell: z.string().regex(/^[A-Z]{1,3}\d+$/).describe("Top-left cell in the destination, e.g. 'A1'"),
      sourceSheet: z.string().optional().describe("Source sheet name (defaults to first sheet)"),
      sourceTable: z.string().optional().describe("Source table name (defaults to first table)"),
      destSheet: z.string().optional().describe("Destination sheet name (defaults to first sheet)"),
      destTable: z.string().optional().describe("Destination table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ sourceDoc, sourceRange, destDoc, destCell, sourceSheet, sourceTable, destSheet, destTable }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");

      function parseRef(ref) {
        const m = ref.match(/^([A-Z]+)(\\d+)$/);
        let col = 0;
        for (let i = 0; i < m[1].length; i++) col = col * 26 + m[1].charCodeAt(i) - 64;
        return { row: parseInt(m[2], 10) - 1, col: col - 1 };
      }

      // Source
      const srcDoc = app.documents.byName(params.sourceDoc);
      const srcSheet = params.sourceSheet ? srcDoc.sheets.byName(params.sourceSheet) : srcDoc.sheets[0];
      const srcTable = params.sourceTable ? srcSheet.tables.byName(params.sourceTable) : srcSheet.tables[0];
      const rangeParts = params.sourceRange.split(":");
      const srcStart = parseRef(rangeParts[0]);
      const srcEnd = parseRef(rangeParts[1]);
      const srcCols = srcTable.columnCount();
      const copyRows = srcEnd.row - srcStart.row + 1;
      const copyCols = srcEnd.col - srcStart.col + 1;

      // Read source values
      const values = [];
      for (let r = srcStart.row; r <= srcEnd.row; r++) {
        const row = [];
        for (let c = srcStart.col; c <= srcEnd.col; c++) {
          row.push(srcTable.cells[r * srcCols + c].value());
        }
        values.push(row);
      }

      // Destination
      const dstDoc = app.documents.byName(params.destDoc);
      const dstSheet = params.destSheet ? dstDoc.sheets.byName(params.destSheet) : dstDoc.sheets[0];
      const dstTable = params.destTable ? dstSheet.tables.byName(params.destTable) : dstSheet.tables[0];
      const dstStart = parseRef(params.destCell);
      const dstCols = dstTable.columnCount();

      // Auto-resize destination if needed
      const needRows = dstStart.row + copyRows;
      const needCols = dstStart.col + copyCols;
      while (dstTable.rowCount() < needRows) dstTable.rows.push(app.Row());
      while (dstTable.columnCount() < needCols) dstTable.columns.push(app.Column());
      const newDstCols = dstTable.columnCount();

      // Write values
      let cellsCopied = 0;
      for (let r = 0; r < copyRows; r++) {
        for (let c = 0; c < copyCols; c++) {
          const val = values[r][c];
          if (val !== null && val !== undefined) {
            dstTable.cells[(dstStart.row + r) * newDstCols + (dstStart.col + c)].value = val;
            cellsCopied++;
          }
        }
      }

      return JSON.stringify({ copied: true, cellsCopied, rows: copyRows, columns: copyCols });
    `, {
      sourceDoc, sourceRange, destDoc, destCell,
      sourceSheet: sourceSheet ?? null, sourceTable: sourceTable ?? null,
      destSheet: destSheet ?? null, destTable: destTable ?? null,
    })),
  );

  // ── Image Tools ──

  server.tool(
    "numbers_add_image",
    "Add an image to a sheet from a file path",
    {
      documentName: z.string().describe("Name of the open document"),
      filePath: z.string().startsWith("/").describe("Absolute path to the image file"),
      x: z.number().optional().describe("X position in points"),
      y: z.number().optional().describe("Y position in points"),
      width: z.number().positive().optional().describe("Width in points"),
      height: z.number().positive().optional().describe("Height in points"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, filePath, x, y, width, height, sheetName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const img = app.Image({ file: Path(params.filePath) });
      sheet.images.push(img);
      const added = sheet.images[sheet.images.length - 1];
      if (params.x != null || params.y != null) {
        var pos = added.position();
        added.position = { x: params.x != null ? params.x : pos.x, y: params.y != null ? params.y : pos.y };
      }
      if (params.width != null) added.width = params.width;
      if (params.height != null) added.height = params.height;
      return JSON.stringify({
        added: true,
        imageCount: sheet.images.length,
        width: added.width(),
        height: added.height(),
        position: added.position(),
      });
    `, { documentName, filePath, x: x ?? null, y: y ?? null, width: width ?? null, height: height ?? null, sheetName: sheetName ?? null })),
  );

  server.tool(
    "numbers_list_images",
    "List all images on a sheet with their positions and sizes",
    {
      documentName: z.string().describe("Name of the open document"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, sheetName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const imgs = sheet.images();
      const result = [];
      for (var i = 0; i < imgs.length; i++) {
        var img = imgs[i];
        var entry = { index: i, width: img.width(), height: img.height(), position: img.position() };
        try { entry.fileName = img.fileName(); } catch(e) {}
        try { entry.description = img.description(); } catch(e) {}
        result.push(entry);
      }
      return JSON.stringify(result);
    `, { documentName, sheetName: sheetName ?? null })),
  );

  // ── Chart Tool ──

  server.tool(
    "numbers_add_chart",
    "Create a chart from table data. Selects the specified data range and creates a chart on the same sheet. Chart type can be set after creation.",
    {
      documentName: z.string().describe("Name of the open document"),
      dataRange: z.string().regex(/^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/).describe("Data range to chart, e.g. 'A1:B10' (include headers)"),
      chartType: z.enum(["bar_2d", "line_2d", "pie_2d", "area_2d", "column_2d", "stacked_bar_2d", "stacked_column_2d"]).optional()
        .describe("Chart type (default: column_2d)"),
      sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
      tableName: z.string().optional().describe("Table name (defaults to first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, dataRange, chartType, sheetName, tableName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Numbers");
      const doc = app.documents.byName(params.documentName);
      const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
      const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];

      // Select the data range
      table.selectionRange = table.ranges[params.dataRange];

      // Find sheet index by name (JXA indexOf doesn't work on scriptable objects)
      const sheets = doc.sheets();
      const sheetName = sheet.name();
      let sheetIdx = 1;
      for (let si = 0; si < sheets.length; si++) {
        if (sheets[si].name() === sheetName) { sheetIdx = si + 1; break; }
      }

      // Create chart via AppleScript bridge (JXA chart creation doesn't bind to selection)
      ObjC.import("Foundation");
      const docName = doc.name();
      const makeChartAS = 'tell application "Numbers" to tell sheet ' + sheetIdx + ' of document "' + docName + '" to make new chart';
      const makeScript = $.NSAppleScript.alloc.initWithSource(makeChartAS);
      const errDict = Ref();
      const makeResult = makeScript.executeAndReturnError(errDict);
      if (!makeResult) {
        const errInfo = ObjC.deepUnwrap(errDict[0]);
        throw new Error("Failed to create chart: " + (errInfo.NSAppleScriptErrorBriefMessage || "unknown error"));
      }

      // Set chart type via AppleScript if requested
      const chartType = params.chartType || "column_2d";
      const setTypeAS = 'tell application "Numbers" to tell sheet ' + sheetIdx + ' of document "' + docName + '" to set chart type of chart 1 to ' + chartType;
      const setScript = $.NSAppleScript.alloc.initWithSource(setTypeAS);
      const errDict2 = Ref();
      setScript.executeAndReturnError(errDict2);

      // Read chart info via JXA
      const chartCount = sheet.charts.length;
      const chart = sheet.charts[chartCount - 1];
      return JSON.stringify({
        chartCreated: true,
        chartCount: chartCount,
        width: chart.width(),
        height: chart.height(),
      });
    `, {
      documentName, dataRange,
      chartType: chartType ?? null,
      sheetName: sheetName ?? null,
      tableName: tableName ?? null,
    })),
  );

  // ── Creator Studio Features (require subscription) ──

  server.tool(
    "numbers_magic_fill",
    "Use AI to fill cells based on a pattern (Creator Studio only). Provide example values in the first rows of a column, then specify the range of empty cells to fill. Magic Fill detects the pattern and fills the remaining cells automatically.",
    {
      documentName: z.string().describe("Name of the open document"),
      range: z.string().describe("Cell range to fill, e.g. 'C3:C20' (must follow rows that contain example values showing the pattern)"),
      sheetName: z.string().optional().describe("Sheet name (default: first sheet)"),
      tableName: z.string().optional().describe("Table name (default: first table)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, range, sheetName, tableName }) => {
      if (!isCreatorStudio()) {
        return toolResult("Magic Fill requires Apple Creator Studio (iWork 15.1+). Install from the App Store.", true);
      }
      return handleJXA(async () => {
        const appName = resolveAppName("Numbers");

        // 1. Enter the table and set the selection range via JXA
        await runJXA<void>(`
          const app = Application("Numbers");
          app.activate();
          const doc = app.documents.byName(params.documentName);
          const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
          const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
          table.selectionRange = table.ranges[params.range];
        `, { documentName, range, sheetName: sheetName ?? null, tableName: tableName ?? null },
        { label: "magic_fill:select" });

        // 2. Enter the table in the UI (needed for menu to activate)
        await runJXA<void>(`
          const app = Application("Numbers");
          app.activate();
          delay(0.3);
          const se = Application("System Events");
          se.processes.byName("Numbers").frontmost();
          se.keystroke("\\t");
          delay(0.3);
          // Re-set selection after entering table
          const doc = app.documents.byName(params.documentName);
          const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
          const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
          table.selectionRange = table.ranges[params.range];
        `, { documentName, range, sheetName: sheetName ?? null, tableName: tableName ?? null },
        { label: "magic_fill:enter_table" });

        // 3. Click "Magic Fill Cells" (or "Magic Fill Cell" for single cell)
        // The title depends on the version and the selection ("Magic Fill" in 15.4)
        let filled = false;
        for (const title of ["Magic Fill Cells", "Magic Fill Cell", "Magic Fill"]) {
          try {
            await clickMenuItem("Numbers", ["Table", title]);
            filled = true;
            break;
          } catch {}
        }
        if (!filled) {
          throw new Error("Magic Fill is not available. Ensure the selected range is adjacent to rows with example values that show a recognizable pattern.");
        }

        // 4. Wait for AI to process suggestions and accept with Return
        await runJXA<void>(`
          delay(3);
          const se = Application("System Events");
          se.keystroke("\\r");
          delay(1);
        `, undefined, { label: "magic_fill:accept", timeout: 15_000 });

        // 5. Read back the filled values
        return runJXA<string>(`
          const app = Application("Numbers");
          const doc = app.documents.byName(params.documentName);
          const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
          const table = params.tableName ? sheet.tables.byName(params.tableName) : sheet.tables[0];
          const range = table.ranges[params.range];
          const cells = range.cells();
          const values = cells.map(c => ({ name: c.name(), value: c.value() }));
          const filled = values.filter(v => v.value !== null).length;
          const total = values.length;
          return JSON.stringify({ filled, total, values });
        `, { documentName, range, sheetName: sheetName ?? null, tableName: tableName ?? null },
        { label: "magic_fill:read" });
      });
    },
  );

  server.tool(
    "numbers_super_resolution",
    "Upscale an image using AI Super Resolution (Creator Studio only). Increases resolution while preserving quality.",
    {
      documentName: z.string().describe("Name of the open document"),
      imageIndex: z.number().int().min(1).describe("Image index (1-based) on the sheet"),
      sheetName: z.string().optional().describe("Sheet name (default: first sheet)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, imageIndex, sheetName }) => {
      if (!isCreatorStudio()) {
        return toolResult("Super Resolution requires Apple Creator Studio (iWork 15.1+).", true);
      }
      return handleJXA(async () => {
        // Select the image via JXA
        await runJXA<void>(`
          const app = Application("Numbers");
          app.activate();
          const doc = app.documents.byName(params.documentName);
          const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
          const images = sheet.images();
          if (params.imageIndex > images.length) throw new Error("Image index " + params.imageIndex + " out of range (sheet has " + images.length + " images)");
          app.selection = [images[params.imageIndex - 1]];
        `, { documentName, imageIndex, sheetName: sheetName ?? null },
        { label: "super_resolution:select" });

        // Click Format > Image > Super Resolution
        await clickMenuItem("Numbers", ["Format", "Image", "Super Resolution"], { postdelay: 2 });

        return JSON.stringify({ success: true, message: "Super Resolution started. The image is being upscaled in the background." });
      });
    },
  );

  server.tool(
    "numbers_remove_background",
    "Remove the background from an image using AI (Creator Studio only). Opens background removal mode.",
    {
      documentName: z.string().describe("Name of the open document"),
      imageIndex: z.number().int().min(1).describe("Image index (1-based) on the sheet"),
      sheetName: z.string().optional().describe("Sheet name (default: first sheet)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, imageIndex, sheetName }) => {
      if (!isCreatorStudio()) {
        return toolResult("Remove Background requires Apple Creator Studio (iWork 15.1+).", true);
      }
      return handleJXA(async () => {
        // Select the image via JXA
        await runJXA<void>(`
          const app = Application("Numbers");
          app.activate();
          const doc = app.documents.byName(params.documentName);
          const sheet = params.sheetName ? doc.sheets.byName(params.sheetName) : doc.sheets[0];
          const images = sheet.images();
          if (params.imageIndex > images.length) throw new Error("Image index " + params.imageIndex + " out of range (sheet has " + images.length + " images)");
          app.selection = [images[params.imageIndex - 1]];
        `, { documentName, imageIndex, sheetName: sheetName ?? null },
        { label: "remove_background:select" });

        // Click Format > Image > Remove Background
        await clickMenuItem("Numbers", ["Format", "Image", "Remove Background"], { postdelay: 2 });

        return JSON.stringify({ success: true, message: "Background removal started. The background is being analyzed and removed." });
      });
    },
  );
}
