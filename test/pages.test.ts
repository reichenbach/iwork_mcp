import { describe, it, before, after } from "node:test";
import assert from "node:assert/strict";
import { unlinkSync } from "node:fs";
import { createTestServer, type TestContext } from "./helpers/server.js";
import { isAppAvailable } from "./helpers/app-check.js";

async function call(ctx: TestContext, name: string, args: Record<string, unknown> = {}) {
  const result = await ctx.client.callTool({ name, arguments: args });
  const text = (result.content as Array<{ type: string; text: string }>)[0].text;
  let json: unknown;
  try { json = JSON.parse(text); } catch { json = null; }
  return { text, json, isError: result.isError };
}

describe("Pages Integration", async () => {
  const available = await isAppAvailable("Pages");
  if (!available) {
    console.log("⏭  Skipping Pages tests — app not available");
    return;
  }

  let ctx: TestContext;
  let docName: string;
  const suffix = Date.now();
  const tmpFiles: string[] = [];

  before(async () => {
    ctx = await createTestServer();
  });

  after(async () => {
    if (ctx) {
      // Close ALL open Pages documents (catches leaked test docs)
      try {
        const { json: docs } = await call(ctx, "pages_list_documents");
        for (const doc of docs as Array<{ name: string }>) {
          try { await call(ctx, "pages_close_document", { documentName: doc.name, saving: "no" }); } catch {}
        }
      } catch {}
      // Clean up temp files
      for (const f of tmpFiles) {
        try { unlinkSync(f); } catch {}
      }
      await ctx.cleanup();
    }
  });

  it("creates a document with formatted content", async () => {
    const { json } = await call(ctx, "pages_create_document_with_content", {
      paragraphs: [
        { text: "Test Title", fontSize: 24, fontName: "HelveticaNeue-Bold" },
        { text: "This is body text for testing.", fontSize: 12 },
        { text: "A third paragraph.", fontSize: 12 },
      ],
    });
    docName = json.name;
    assert.ok(docName, "Document should have a name");
    assert.equal(json.paragraphCount, 3);
  });

  it("reads body text", async () => {
    const { json } = await call(ctx, "pages_get_body_text", { documentName: docName });
    assert.ok(json.text.includes("Test Title"));
    assert.ok(json.text.includes("body text for testing"));
    assert.ok(json.text.includes("third paragraph"));
  });

  it("reads paragraphs with formatting info", async () => {
    const { json: paragraphs } = await call(ctx, "pages_get_paragraphs", {
      documentName: docName,
    });
    assert.ok(Array.isArray(paragraphs));
    assert.ok(paragraphs.length >= 3);
    assert.ok(paragraphs[0].text.includes("Test Title"));
  });

  it("gets document info", async () => {
    const { json } = await call(ctx, "pages_get_document_info", { documentName: docName });
    assert.ok(json.name);
    assert.ok(json.wordCount > 0, "Should have words");
    assert.ok(json.characterCount > 0, "Should have characters");
    assert.ok(json.pageCount >= 1, "Should have at least one page");
    assert.ok(json.paragraphCount >= 3, "Should have at least 3 paragraphs");
    assert.equal(typeof json.modified, "boolean");
  });

  it("lists images (empty document)", async () => {
    const { json } = await call(ctx, "pages_list_images", { documentName: docName });
    assert.ok(Array.isArray(json));
    assert.equal(json.length, 0, "New document should have no images");
  });

  it("appends text", async () => {
    const { json } = await call(ctx, "pages_add_text", {
      documentName: docName,
      text: "Appended paragraph.\n",
    });
    assert.ok(json.appended);

    const { json: body } = await call(ctx, "pages_get_body_text", { documentName: docName });
    assert.ok(body.text.includes("Appended paragraph"));
  });

  it("replaces text", async () => {
    const { json } = await call(ctx, "pages_replace_text", {
      documentName: docName,
      find: "Appended paragraph.",
      replace: "Replaced paragraph.",
    });
    assert.ok(json.replacements >= 1);

    const { json: body } = await call(ctx, "pages_get_body_text", { documentName: docName });
    assert.ok(body.text.includes("Replaced paragraph"));
    assert.ok(!body.text.includes("Appended paragraph"));
  });

  // ── Format text ──

  it("formats a paragraph", async () => {
    const { json } = await call(ctx, "pages_format_text", {
      documentName: docName,
      paragraphIndex: 0,
      format: { fontName: "Georgia-Bold", fontSize: 28, textColor: "#FF0000" },
    });
    assert.ok(json.formatted);

    const { json: paragraphs } = await call(ctx, "pages_get_paragraphs", { documentName: docName });
    const first = paragraphs[0];
    assert.equal(first.font, "Georgia-Bold");
    assert.equal(first.size, 28);
    // Color roundtrip is lossy due to JXA color space conversion — verify it's approximately red
    assert.ok(first.color, "color should be present");
    const r = parseInt(first.color.slice(1, 3), 16);
    const g = parseInt(first.color.slice(3, 5), 16);
    const b = parseInt(first.color.slice(5, 7), 16);
    assert.ok(r > 200, `red channel ${r} should be high`);
    assert.ok(g < 20, `green channel ${g} should be low`);
    assert.ok(b < 20, `blue channel ${b} should be low`);
  });

  // ── Insert text at position ──

  it("inserts text at beginning and middle", async () => {
    // Insert at beginning (afterParagraph = -1)
    const { json: r1 } = await call(ctx, "pages_insert_text_at", {
      documentName: docName,
      text: "Inserted at beginning.\n",
      afterParagraph: -1,
    });
    assert.ok(r1.inserted);

    const { json: body1 } = await call(ctx, "pages_get_body_text", { documentName: docName });
    const lines1 = body1.text.split("\n").filter((l: string) => l.length > 0);
    assert.equal(lines1[0], "Inserted at beginning.");

    // Insert after paragraph 1 (middle)
    const { json: r2 } = await call(ctx, "pages_insert_text_at", {
      documentName: docName,
      text: "Inserted in middle.\n",
      afterParagraph: 1,
    });
    assert.ok(r2.inserted);

    const { json: body2 } = await call(ctx, "pages_get_body_text", { documentName: docName });
    const lines2 = body2.text.split("\n").filter((l: string) => l.length > 0);
    assert.equal(lines2[2], "Inserted in middle.");
  });

  // ── Delete text ──

  it("deletes a paragraph by index", async () => {
    // Get current paragraphs to know what we're deleting
    const { json: before } = await call(ctx, "pages_get_paragraphs", { documentName: docName });
    const countBefore = before.length;
    const deletedText = before[0].text;

    const { json } = await call(ctx, "pages_delete_text", {
      documentName: docName,
      paragraphIndex: 0,
    });
    assert.ok(json.deleted);

    const { json: body } = await call(ctx, "pages_get_body_text", { documentName: docName });
    assert.ok(!body.text.startsWith(deletedText));

    const { json: afterParas } = await call(ctx, "pages_get_paragraphs", { documentName: docName });
    assert.ok(afterParas.length < countBefore);
  });

  // ── Page break ──

  it("inserts a page break between paragraphs", async () => {
    const { json } = await call(ctx, "pages_insert_page_break", {
      documentName: docName,
      afterParagraph: 0,
    });
    assert.ok(json.pageBreakInserted);

    // Verify form feed is present in body text
    const { json: body } = await call(ctx, "pages_get_body_text", { documentName: docName });
    assert.ok(body.text.includes("\f"), "Body text should contain form feed character");
  });

  // ── Export to PDF ──

  it("exports to PDF", async () => {
    const tmpPath = `/tmp/iwork_test_${suffix}.pdf`;
    tmpFiles.push(tmpPath);
    const { json } = await call(ctx, "pages_export_document", {
      documentName: docName,
      filePath: tmpPath,
      format: "PDF",
    }) as { json: { exported: boolean; format: string } };
    assert.ok(json.exported);
    assert.equal(json.format, "PDF");
  });

  // ── Save/close/reopen cycle ──

  it("saves, closes, and reopens a document", async () => {
    const tmpPath = `/tmp/iwork_test_reopen_${suffix}.pages`;
    tmpFiles.push(tmpPath);

    const { json: saveResult } = await call(ctx, "pages_save_document", {
      documentName: docName,
      filePath: tmpPath,
    }) as { json: { saved: boolean; name: string } };
    assert.ok(saveResult.saved);
    docName = saveResult.name;

    await call(ctx, "pages_close_document", { documentName: docName, saving: "no" });

    const { json: openResult } = await call(ctx, "pages_open_document", { filePath: tmpPath }) as { json: { name: string } };
    docName = openResult.name;

    const { json: body } = await call(ctx, "pages_get_body_text", { documentName: docName }) as { json: { text: string } };
    assert.ok(body.text.includes("Test Title"));
  });

  // ── Table tools (AppleScript bridge) ──
  // Table creation is broken on Creator Studio 15.x (-2763 TMAScriptTableInfoProxy,
  // both JXA and AppleScript), so these tests use the built-in Invoice template,
  // which ships with a table. Read/write/resize work on both 14.x and 15.x.

  let tableDocName = "";

  it("creates a document from an invoice template", async (t) => {
    const { json: templates } = await call(ctx, "pages_list_templates");
    const invoice = (templates as string[]).find((n) => n.toLowerCase().includes("invoice"));
    if (!invoice) return t.skip("no invoice template available");
    const { json } = await call(ctx, "pages_create_document", { templateName: invoice });
    tableDocName = json.name;
    assert.ok(tableDocName);
  });

  it("lists tables in the template document", async (t) => {
    if (!tableDocName) return t.skip("no table document");
    const { json } = await call(ctx, "pages_list_tables", { documentName: tableDocName });
    assert.ok(Array.isArray(json));
    assert.ok(json.length >= 1, "Invoice template should contain a table");
    assert.ok(json[0].rows >= 2);
    assert.ok(json[0].columns >= 2);
  });

  it("reads table values and formulas", async (t) => {
    if (!tableDocName) return t.skip("no table document");
    const { json } = await call(ctx, "pages_read_table", { documentName: tableDocName });
    assert.ok(json.rows >= 2);
    assert.ok(Array.isArray(json.values));
    assert.equal(json.values.length, json.rows);
    assert.ok(Array.isArray(json.formulas));
  });

  it("writes cells with escaping and typed round-trip", async (t) => {
    if (!tableDocName) return t.skip("no table document");
    const tricky = 'He said "hi" \\ there';
    const { json } = await call(ctx, "pages_write_table_cells", {
      documentName: tableDocName,
      autoGrow: false,
      cells: [
        { row: 2, column: 1, value: tricky },
        { row: 2, column: 2, value: 3 },
      ],
    });
    assert.equal(json.written, 2);

    const { json: table } = await call(ctx, "pages_read_table", { documentName: tableDocName });
    assert.equal(table.values[1][0], tricky, "quotes/backslashes should round-trip");
    assert.equal(table.values[1][1], 3, "numbers should round-trip typed");
  });

  it("writes a formula that evaluates", async (t) => {
    if (!tableDocName) return t.skip("no table document");
    const { isError } = await call(ctx, "pages_write_table_cells", {
      documentName: tableDocName,
      autoGrow: false,
      cells: [{ row: 2, column: 3, value: "=B2*2" }],
    });
    assert.ok(!isError);

    const { json: table } = await call(ctx, "pages_read_table", { documentName: tableDocName });
    assert.equal(table.values[1][2], 6, "=B2*2 should evaluate to 6");
    assert.ok(typeof table.formulas[1][2] === "string", "formula text should be readable");
  });

  it("resizes the table and returns errors for bad input", async (t) => {
    if (!tableDocName) return t.skip("no table document");
    const { json: before } = await call(ctx, "pages_read_table", { documentName: tableDocName });
    const { json: grown } = await call(ctx, "pages_resize_table", {
      documentName: tableDocName,
      rows: before.rows + 2,
    });
    assert.equal(grown.rows, before.rows + 2);
    const { json: shrunk } = await call(ctx, "pages_resize_table", {
      documentName: tableDocName,
      rows: before.rows,
    });
    assert.equal(shrunk.rows, before.rows);

    const bad = await call(ctx, "pages_read_table", { documentName: tableDocName, tableIndex: 99 });
    assert.equal(bad.isError, true, "bad tableIndex should return isError");
  });

  it("closes the table document without saving", async (t) => {
    if (!tableDocName) return t.skip("no table document");
    const { json } = await call(ctx, "pages_close_document", {
      documentName: tableDocName,
      saving: "no",
    });
    assert.ok(json.closed);
    tableDocName = "";
  });

  // ── Error on invalid document ──

  it("returns isError for a nonexistent document", async () => {
    const result = await ctx.client.callTool({
      name: "pages_get_body_text",
      arguments: { documentName: "NoSuchDocument_999" },
    });
    assert.equal(result.isError, true);
  });

  it("closes the document without saving", async () => {
    const { json } = await call(ctx, "pages_close_document", {
      documentName: docName,
      saving: "no",
    });
    assert.ok(json.closed);
    docName = "";
  });
});
