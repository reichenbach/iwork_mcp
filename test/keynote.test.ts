import { describe, it, before, after } from "node:test";
import assert from "node:assert/strict";
import { unlinkSync } from "node:fs";
import { createTestServer, type TestContext } from "./helpers/server.js";
import { isAppAvailable } from "./helpers/app-check.js";
import { runJXA } from "../src/jxa.js";

async function call(ctx: TestContext, name: string, args: Record<string, unknown> = {}) {
  const result = await ctx.client.callTool({ name, arguments: args });
  const text = (result.content as Array<{ type: string; text: string }>)[0].text;
  let json: unknown;
  try { json = JSON.parse(text); } catch { json = null; }
  return { text, json, isError: result.isError };
}

describe("Keynote Integration", async () => {
  const available = await isAppAvailable("Keynote");
  if (!available) {
    console.log("⏭  Skipping Keynote tests — app not available");
    return;
  }

  let ctx: TestContext;
  let docName: string;
  const suffix = Date.now();
  const tmpFiles: string[] = [];

  before(async () => {
    ctx = await createTestServer();
  });

  // Layout and theme names are localized (e.g. "Leer" instead of "Blank" in German)
  async function layoutNames(): Promise<string[]> {
    const { json } = await call(ctx, "keynote_list_master_slides", { documentName: docName });
    return (json as Array<{ name: string }>).map(l => l.name);
  }
  async function blankLayout(): Promise<string | undefined> {
    const names = await layoutNames();
    return names.find(n => ["Blank", "Leer", "Vide", "Vacío", "Vuoto", "Leeg", "Em branco", "空白"].includes(n));
  }

  after(async () => {
    if (ctx) {
      // Close ALL open Keynote presentations (catches leaked test docs)
      try {
        const { json: docs } = await call(ctx, "keynote_list_presentations");
        for (const doc of docs as Array<{ name: string }>) {
          try { await call(ctx, "keynote_close_presentation", { documentName: doc.name, saving: "no" }); } catch {}
        }
      } catch {}
      // Clean up temp files
      for (const f of tmpFiles) {
        try { unlinkSync(f); } catch {}
      }
      await ctx.cleanup();
    }
  });

  it("creates a new presentation", async () => {
    const { json } = await call(ctx, "keynote_create_presentation");
    docName = json.name;
    assert.ok(docName, "Presentation should have a name");
    assert.ok(json.slideCount >= 1, "Should have at least one slide");
  });

  it("lists slides", async () => {
    const { json: slides } = await call(ctx, "keynote_list_slides", {
      documentName: docName,
    });
    assert.ok(Array.isArray(slides));
    assert.ok(slides.length >= 1);
    assert.equal(slides[0].slideNumber, 1);
  });

  it("adds a new slide", async () => {
    const { json } = await call(ctx, "keynote_add_slide", { documentName: docName });
    assert.ok(json.added);

    const { json: slides } = await call(ctx, "keynote_list_slides", {
      documentName: docName,
    });
    assert.ok(slides.length >= 2);
  });

  it("sets slide title and reads it back", async () => {
    const { json: setResult } = await call(ctx, "keynote_set_slide_title", {
      documentName: docName,
      slideNumber: 1,
      title: "Integration Test Title",
    });
    assert.ok(setResult.set);

    const { json: content } = await call(ctx, "keynote_get_slide_content", {
      documentName: docName,
      slideNumber: 1,
    });
    assert.equal(content.title, "Integration Test Title");
  });

  it("sets slide body and reads it back", async () => {
    const { json: setResult } = await call(ctx, "keynote_set_slide_body", {
      documentName: docName,
      slideNumber: 1,
      body: "Bullet one\nBullet two",
    });
    assert.ok(setResult.set);

    const { json: content } = await call(ctx, "keynote_get_slide_content", {
      documentName: docName,
      slideNumber: 1,
    });
    assert.ok(content.body.includes("Bullet one"));
    assert.ok(content.body.includes("Bullet two"));
  });

  it("sets presenter notes and reads them back", async () => {
    const { json: setResult } = await call(ctx, "keynote_set_presenter_notes", {
      documentName: docName,
      slideNumber: 1,
      notes: "Remember to pause here.",
    });
    assert.ok(setResult.set);

    const { json: content } = await call(ctx, "keynote_get_slide_content", {
      documentName: docName,
      slideNumber: 1,
    });
    assert.equal(content.presenterNotes, "Remember to pause here.");
  });

  // ── Delete slide ──

  it("deletes a slide", async () => {
    // Add an extra slide, then delete it
    await call(ctx, "keynote_add_slide", { documentName: docName });
    const { json: before } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const countBefore = before.length;

    const { json } = await call(ctx, "keynote_delete_slide", {
      documentName: docName,
      slideNumber: countBefore,
    });
    assert.ok(json.deleted);
    assert.equal(json.remainingSlides, countBefore - 1);
  });

  // ── Duplicate slide ──

  it("duplicates a slide", async () => {
    const { json: before } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const countBefore = before.length;

    const { json } = await call(ctx, "keynote_duplicate_slide", {
      documentName: docName,
      slideNumber: 1,
    });
    assert.ok(json.duplicated);
    assert.equal(json.totalSlides, countBefore + 1);
    assert.equal(json.slideNumber, json.totalSlides);

    // Clean up: delete the duplicate
    await call(ctx, "keynote_delete_slide", {
      documentName: docName,
      slideNumber: json.totalSlides,
    });
  });

  // ── Reorder slide ──

  it("reorders slides", async () => {
    // Ensure at least 3 slides
    await call(ctx, "keynote_add_slide", { documentName: docName });
    await call(ctx, "keynote_add_slide", { documentName: docName });

    // Set distinct titles to track
    await call(ctx, "keynote_set_slide_title", { documentName: docName, slideNumber: 1, title: "Slide A" });
    await call(ctx, "keynote_set_slide_title", { documentName: docName, slideNumber: 2, title: "Slide B" });

    // Move slide 2 to position 1
    const { json } = await call(ctx, "keynote_reorder_slide", {
      documentName: docName,
      fromSlideNumber: 2,
      toSlideNumber: 1,
    });
    assert.ok(json.moved);

    // Verify: the first slide should now be "Slide B"
    const { json: content } = await call(ctx, "keynote_get_slide_content", {
      documentName: docName,
      slideNumber: 1,
    });
    assert.equal(content.title, "Slide B");
  });

  it("reorders a slide forward without losing the target slide", async () => {
    const { json: before } = await call(ctx, "keynote_list_slides", { documentName: docName });
    assert.ok(before.length >= 3);
    for (let i = 1; i <= 3; i++) {
      await call(ctx, "keynote_set_slide_title", { documentName: docName, slideNumber: i, title: `Order ${i}` });
    }

    const { json, isError } = await call(ctx, "keynote_reorder_slide", {
      documentName: docName,
      fromSlideNumber: 1,
      toSlideNumber: 3,
    });
    assert.ok(!isError && json.moved);

    const { json: after } = await call(ctx, "keynote_list_slides", { documentName: docName });
    assert.equal(after.length, before.length);
    assert.deepEqual(after.slice(0, 3).map((s: { title: string }) => s.title), ["Order 2", "Order 3", "Order 1"]);
  });

  it("adds a slide after a given slide", async () => {
    const { json: before } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const { json, isError } = await call(ctx, "keynote_add_slide", { documentName: docName, afterSlide: 1 });
    assert.ok(!isError && json.added);
    assert.equal(json.slideNumber, 2);
    assert.equal(json.totalSlides, before.length + 1);

    const { json: after } = await call(ctx, "keynote_list_slides", { documentName: docName });
    assert.equal(after.length, before.length + 1);
    assert.equal(after[2].title, before[1].title, "Former slide 2 should now be slide 3");
    await call(ctx, "keynote_delete_slide", { documentName: docName, slideNumber: 2 });
  });

  it("adds a slide with a master layout after a given slide", async () => {
    const { json: before } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const layout = (await layoutNames()).pop();
    const { json, isError } = await call(ctx, "keynote_add_slide", {
      documentName: docName, masterSlideName: layout, afterSlide: 1,
    });
    assert.ok(!isError && json.added);
    assert.equal(json.slideNumber, 2);

    const { json: after } = await call(ctx, "keynote_list_slides", { documentName: docName });
    assert.equal(after.length, before.length + 1);
    await call(ctx, "keynote_delete_slide", { documentName: docName, slideNumber: 2 });
  });

  it("reports an error for an unknown master layout or document", async () => {
    const { json: before } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const badMaster = await call(ctx, "keynote_add_slide", {
      documentName: docName, masterSlideName: "NoSuchLayout", afterSlide: 1,
    });
    assert.ok(badMaster.isError);
    const badDoc = await call(ctx, "keynote_add_slide", {
      documentName: `NoSuchDocument-${suffix}`, masterSlideName: "Blank",
    });
    assert.ok(badDoc.isError);

    const { json: after } = await call(ctx, "keynote_list_slides", { documentName: docName });
    assert.equal(after.length, before.length);
  });

  it("duplicates a slide to a given position", async () => {
    await call(ctx, "keynote_set_slide_title", { documentName: docName, slideNumber: 2, title: "Dup source" });
    const { json: before } = await call(ctx, "keynote_list_slides", { documentName: docName });

    const { json, isError } = await call(ctx, "keynote_duplicate_slide", {
      documentName: docName, slideNumber: 2, toSlideNumber: 1,
    });
    assert.ok(!isError && json.duplicated);
    assert.equal(json.slideNumber, 1);

    const { json: after } = await call(ctx, "keynote_list_slides", { documentName: docName });
    assert.equal(after.length, before.length + 1);
    assert.equal(after[0].title, "Dup source");
    await call(ctx, "keynote_delete_slide", { documentName: docName, slideNumber: 1 });
  });

  it("creates a presentation with a custom slide size", async () => {
    const { json, isError } = await call(ctx, "keynote_create_presentation", { width: 1920, height: 1080 });
    assert.ok(!isError, "create with size should succeed");
    assert.equal(json.width, 1920);
    assert.equal(json.height, 1080);
    await call(ctx, "keynote_close_presentation", { documentName: json.name, saving: "no" });
  });

  it("lists each text box on a slide only once", async () => {
    await call(ctx, "keynote_set_slide_title", { documentName: docName, slideNumber: 1, title: "Unique title" });
    const { json: content } = await call(ctx, "keynote_get_slide_content", { documentName: docName, slideNumber: 1 });
    const matches = [...content.textItems, ...content.shapes].filter((t: { text: string }) => t.text === "Unique title");
    assert.ok(matches.length <= 1, `Title listed ${matches.length} times`);
  });

  // ── Layout tools ──

  it("positions a shape on a slide", async () => {
    // Add a fresh slide with shapes for layout tests
    await call(ctx, "keynote_add_slide", { documentName: docName });
    const { json: slides } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const layoutSlide = slides.length;

    // Add 3 shapes with known positions
    const { json: added } = await call(ctx, "keynote_add_shape", {
      documentName: docName, slideNumber: layoutSlide, text: "Shape A",
      x: 50, y: 50, width: 100, height: 80,
    });
    assert.ok(Math.abs(added.x - 50) < 2, `Expected x~50, got ${added.x}`);
    assert.ok(Math.abs(added.y - 50) < 2, `Expected y~50, got ${added.y}`);

    // Position the shape we just added (index 0 on this fresh slide)
    const { json } = await call(ctx, "keynote_position_item", {
      documentName: docName,
      slideNumber: layoutSlide,
      itemType: "shape",
      itemIndex: 0,
      x: 200,
      y: 300,
      width: 150,
    });
    assert.ok(json.positioned);
    assert.ok(Math.abs(json.x - 200) < 2, `Expected x~200, got ${json.x}`);
    assert.ok(Math.abs(json.y - 300) < 2, `Expected y~300, got ${json.y}`);
    assert.equal(json.width, 150);

    // Add two more shapes for align/distribute tests
    await call(ctx, "keynote_add_shape", {
      documentName: docName, slideNumber: layoutSlide, text: "Shape B",
      x: 100, y: 100, width: 80, height: 60,
    });
    await call(ctx, "keynote_add_shape", {
      documentName: docName, slideNumber: layoutSlide, text: "Shape C",
      x: 300, y: 200, width: 80, height: 60,
    });

    // Test align
    const { json: alignResult } = await call(ctx, "keynote_align_items", {
      documentName: docName,
      slideNumber: layoutSlide,
      items: [
        { itemType: "shape", itemIndex: 0 },
        { itemType: "shape", itemIndex: 1 },
        { itemType: "shape", itemIndex: 2 },
      ],
      alignment: "left",
    });
    assert.ok(alignResult.aligned);
    assert.equal(alignResult.itemCount, 3);

    // Test distribute
    const { json: distResult } = await call(ctx, "keynote_distribute_items", {
      documentName: docName,
      slideNumber: layoutSlide,
      items: [
        { itemType: "shape", itemIndex: 0 },
        { itemType: "shape", itemIndex: 1 },
        { itemType: "shape", itemIndex: 2 },
      ],
      direction: "vertical",
    });
    assert.ok(distResult.distributed);
    assert.equal(distResult.itemCount, 3);
  });

  // ── Shape formatting tools ──

  it("reads and formats a shape", async () => {
    // Add a slide with a shape for formatting tests
    await call(ctx, "keynote_add_slide", { documentName: docName });
    const { json: slides } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const fmtSlide = slides.length;

    await call(ctx, "keynote_add_shape", {
      documentName: docName, slideNumber: fmtSlide, text: "Format me",
      x: 100, y: 100, width: 200, height: 100,
    });

    const { json: info } = await call(ctx, "keynote_get_shape_info", {
      documentName: docName,
      slideNumber: fmtSlide,
      shapeIndex: 0,
    });
    assert.ok(info.width > 0);
    assert.ok(info.opacity !== undefined);
    assert.ok(info.text !== undefined);

    const { json: fmtResult } = await call(ctx, "keynote_format_shape", {
      documentName: docName,
      slideNumber: fmtSlide,
      shapeIndex: 0,
      format: { opacity: 75, rotation: 15, fontName: "HelveticaNeue-Bold", fontSize: 18, textColor: "#FFFFFF" },
    });
    assert.ok(fmtResult.formatted);
    assert.deepEqual(fmtResult.applied, ["opacity", "rotation", "fontName", "fontSize", "textColor"]);

    // Verify
    const { json: after } = await call(ctx, "keynote_get_shape_info", {
      documentName: docName,
      slideNumber: fmtSlide,
      shapeIndex: 0,
    });
    assert.equal(after.opacity, 75);
    assert.equal(after.rotation, 15);
    assert.equal(after.textFormat.color, "#ffffff", "White text must not come out near-black");

    // Text alignment isn't scriptable: reject it instead of half-applying the format
    const { isError } = await call(ctx, "keynote_format_shape", {
      documentName: docName, slideNumber: fmtSlide, shapeIndex: 0,
      format: { opacity: 50, textAlignment: "center" },
    });
    assert.ok(isError);
  });

  // ── Theme tools ──

  it("gets and sets the presentation theme", async () => {
    // Read current theme
    const { json: before } = await call(ctx, "keynote_get_theme", { documentName: docName });
    assert.ok(before.theme, "Should return a theme name");

    // Change to a different theme
    const { json: themes } = await call(ctx, "keynote_list_themes");
    const targetTheme = (themes as string[]).find(t => t !== before.theme);
    const { json: setResult } = await call(ctx, "keynote_set_theme", {
      documentName: docName,
      themeName: targetTheme,
    });
    assert.ok(setResult.themeSet);
    assert.equal(setResult.theme, targetTheme);

    // Verify it stuck
    const { json: after } = await call(ctx, "keynote_get_theme", { documentName: docName });
    assert.equal(after.theme, targetTheme);
  });

  // ── Slide layout ──

  it("changes a slide layout", async () => {
    const layout = (await blankLayout()) ?? (await layoutNames()).pop();
    const { json } = await call(ctx, "keynote_set_slide_layout", {
      documentName: docName,
      slideNumber: 1,
      masterSlideName: layout,
    });
    assert.ok(json.layoutSet);
    assert.equal(json.masterSlide, layout);
  });

  // ── Slide tables ──

  it("adds a table to a slide and reads it back", async () => {
    // Add a slide for table tests
    await call(ctx, "keynote_add_slide", { documentName: docName });
    const { json: slides } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const tblSlide = slides.length;

    const { json: addResult } = await call(ctx, "keynote_add_table_to_slide", {
      documentName: docName,
      slideNumber: tblSlide,
      data: [
        ["Product", "Q1", "Q2"],
        ["Widget", 100, 200],
        ["Gadget", 300, 400],
      ],
    });
    assert.ok(addResult.added);
    assert.equal(addResult.rows, 3);
    assert.equal(addResult.columns, 3);

    // Read it back
    const { json: readResult } = await call(ctx, "keynote_read_slide_table", {
      documentName: docName,
      slideNumber: tblSlide,
      tableIndex: 0,
    });
    assert.equal(readResult.rows, 3);
    assert.equal(readResult.columns, 3);
    assert.equal(readResult.data[0][0], "Product");
    assert.equal(readResult.data[1][1], 100);
    assert.equal(readResult.data[2][2], 400);
  });

  // ── Write to slide table ──

  it("writes data to an existing slide table", async () => {
    // Add a slide with a table
    await call(ctx, "keynote_add_slide", { documentName: docName });
    const { json: slides } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const writeSlide = slides.length;

    await call(ctx, "keynote_add_table_to_slide", {
      documentName: docName,
      slideNumber: writeSlide,
      rows: 2,
      columns: 2,
    });

    // Write data to the table
    const { json: writeResult } = await call(ctx, "keynote_write_slide_table", {
      documentName: docName,
      slideNumber: writeSlide,
      tableIndex: 0,
      data: [["Name", "Score"], ["Alice", 95]],
    });
    assert.ok(writeResult.written);
    assert.equal(writeResult.cellsWritten, 4);

    // Read back and verify
    const { json: readResult } = await call(ctx, "keynote_read_slide_table", {
      documentName: docName,
      slideNumber: writeSlide,
      tableIndex: 0,
    });
    assert.equal(readResult.data[0][0], "Name");
    assert.equal(readResult.data[0][1], "Score");
    assert.equal(readResult.data[1][0], "Alice");
    assert.equal(readResult.data[1][1], 95);
  });

  // ── Format slide table ──

  it("formats cells in a slide table", async () => {
    // Add a slide with a table for formatting
    await call(ctx, "keynote_add_slide", { documentName: docName });
    const { json: slides } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const fmtSlide = slides.length;

    await call(ctx, "keynote_add_table_to_slide", {
      documentName: docName,
      slideNumber: fmtSlide,
      data: [["Header", "Value"], ["Item", 42]],
    });

    const { json } = await call(ctx, "keynote_format_slide_table", {
      documentName: docName,
      slideNumber: fmtSlide,
      tableIndex: 0,
      cellRange: "A1:B1",
      format: { textColor: "#FF0000" },
    });
    assert.ok(json.formatted);
    assert.equal(json.cellsFormatted, 2);
  });

  // ── Add line ──

  it("adds a line to a slide", async () => {
    await call(ctx, "keynote_add_slide", { documentName: docName });
    const { json: slides } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const lineSlide = slides.length;

    const { json } = await call(ctx, "keynote_add_line", {
      documentName: docName,
      slideNumber: lineSlide,
      startX: 100,
      startY: 200,
      endX: 500,
      endY: 200,
    });
    assert.ok(json.added);
    assert.ok(json.startPoint);
    assert.ok(json.endPoint);
  });

  // ── List slide items ──

  it("lists all items on a slide", async () => {
    // Use a slide that has content (slide 1 has title, body, shapes from earlier tests)
    const { json } = await call(ctx, "keynote_list_slide_items", {
      documentName: docName,
      slideNumber: 1,
    });
    assert.equal(json.slideNumber, 1);
    assert.ok(typeof json.itemCount === "number");
    assert.ok(Array.isArray(json.items));
    assert.equal(json.itemCount, json.items.length);
    // Each item should have a type
    for (const item of json.items) {
      assert.ok(["shape", "textItem", "image", "table", "line"].includes(item.type), `Unknown item type: ${item.type}`);
    }
  });

  // ── Export to PDF ──

  it("exports to PDF", async () => {
    const tmpPath = `/tmp/iwork_test_${suffix}.pdf`;
    tmpFiles.push(tmpPath);
    const { json } = await call(ctx, "keynote_export_presentation", {
      documentName: docName,
      filePath: tmpPath,
      format: "PDF",
    });
    assert.ok(json.exported);
    assert.equal(json.format, "PDF");

    // One PDF page per (non-skipped) slide — Creator Studio used to produce a single blank page
    const { json: slides } = await call(ctx, "keynote_list_slides", { documentName: docName });
    const expected = (slides as Array<{ skipped: boolean }>).filter(s => !s.skipped).length;
    const pageCount = await runJXA<number>(`
      ObjC.import("Quartz");
      const pdf = $.PDFDocument.alloc.initWithURL($.NSURL.fileURLWithPath(params.path));
      return JSON.stringify(Number(pdf.pageCount));
    `, { path: tmpPath });
    assert.equal(pageCount, expected);
  });

  // ── Error on nonexistent presentation ──

  it("returns isError when reading a missing slide or presentation", async () => {
    for (const tool of ["keynote_get_slide_content", "keynote_list_slide_items"]) {
      const missingSlide = await call(ctx, tool, { documentName: docName, slideNumber: 99 });
      assert.ok(missingSlide.isError, `${tool} should fail for slide 99`);
      const missingDoc = await call(ctx, tool, { documentName: "NoSuchPresentation_999", slideNumber: 1 });
      assert.ok(missingDoc.isError, `${tool} should fail for a missing presentation`);
    }
  });

  it("returns isError when saving a nonexistent presentation", async () => {
    const { isError } = await call(ctx, "keynote_save_presentation", { documentName: "NoSuchPresentation_999" });
    assert.ok(isError);
  });

  it("lists each box once in list_slide_items", async () => {
    const { json } = await call(ctx, "keynote_list_slide_items", { documentName: docName, slideNumber: 1 });
    const keys = (json.items as Array<{ type: string; text?: string; width: number; height: number; position: { x: number; y: number } | null }>)
      .filter(i => i.type === "shape" || i.type === "textItem")
      .map(i => [i.text, Math.round(i.width), Math.round(i.height), i.position && Math.round(i.position.x), i.position && Math.round(i.position.y)].join("|"));
    assert.equal(new Set(keys).size, keys.length, "No box should be listed twice");
  });

  it("makes a hidden title placeholder visible when setting a title", async (t) => {
    const blank = await blankLayout();
    if (!blank) return t.skip("no blank layout in this theme");
    const { json: added } = await call(ctx, "keynote_add_slide", { documentName: docName, masterSlideName: blank });
    const { json, isError, text } = await call(ctx, "keynote_set_slide_title", {
      documentName: docName, slideNumber: added.slideNumber, title: "Visible title",
    });
    if (!isError) {
      const { json: slides } = await call(ctx, "keynote_list_slides", { documentName: docName });
      assert.notEqual(slides[added.slideNumber - 1].titleShowing, false, "Title was written but is hidden");
      assert.equal(slides[added.slideNumber - 1].title, "Visible title");
    } else {
      assert.match(text, /no visible title placeholder/);
    }
    assert.ok(isError || json.set);
    await call(ctx, "keynote_delete_slide", { documentName: docName, slideNumber: added.slideNumber });
  });

  it("accepts 'none' as a transition and explains invalid effect names", async () => {
    const { json, isError } = await call(ctx, "keynote_set_transition", { documentName: docName, slideNumber: 1, effect: "none" });
    assert.ok(!isError && json.set);
    const bad = await call(ctx, "keynote_set_transition", { documentName: docName, slideNumber: 1, effect: "NoSuchEffect" });
    assert.ok(bad.isError);
    assert.match(bad.text, /Valid effects: .*dissolve/);
  });

  it("returns isError for a nonexistent presentation", async () => {
    const result = await ctx.client.callTool({
      name: "keynote_list_slides",
      arguments: { documentName: "NoSuchPresentation_999" },
    });
    assert.equal(result.isError, true);
  });

  // ── Compound tool ──

  it("creates a presentation with multiple slides in one call", async () => {
    const { json } = await call(ctx, "keynote_create_presentation_with_slides", {
      slides: [
        {
          title: "Welcome",
          body: "First slide body",
          presenterNotes: "Greet the audience",
          transition: { effect: "dissolve", duration: 1.5 },
        },
        {
          title: "Agenda",
          body: "Item 1\nItem 2\nItem 3",
          presenterNotes: "Walk through agenda",
        },
      ],
    });
    const compoundDocName = json.name;
    assert.ok(compoundDocName);
    assert.equal(json.slideCount, 2);
    assert.equal(json.slides.length, 2);
    assert.equal(json.slides[0].slideNumber, 1);
    assert.equal(json.slides[0].title, "Welcome");
    assert.ok(json.slides[0].hasNotes);
    assert.equal(json.slides[1].slideNumber, 2);
    assert.equal(json.slides[1].title, "Agenda");
    assert.ok(json.slides[1].hasBody);
    assert.ok(json.slides[1].hasNotes);

    // Verify via get_slide_content
    const { json: slide1 } = await call(ctx, "keynote_get_slide_content", {
      documentName: compoundDocName,
      slideNumber: 1,
    });
    assert.equal(slide1.title, "Welcome");
    assert.equal(slide1.presenterNotes, "Greet the audience");

    // Clean up
    await call(ctx, "keynote_close_presentation", { documentName: compoundDocName, saving: "no" });
  });

  it("closes the presentation without saving", async () => {
    const { json } = await call(ctx, "keynote_close_presentation", {
      documentName: docName,
      saving: "no",
    });
    assert.ok(json.closed);
    docName = "";
  });
});
