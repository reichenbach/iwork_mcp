import { describe, it, before, after } from "node:test";
import assert from "node:assert/strict";
import { createTestServer, type TestContext } from "./helpers/server.js";

describe("Tool Registration", () => {
  let ctx: TestContext;

  before(async () => {
    ctx = await createTestServer();
  });

  after(async () => {
    await ctx.cleanup();
  });

  it("sends instructions to the client", async () => {
    const instructions = ctx.client.getInstructions();
    assert.ok(instructions, "Server should send instructions");
    assert.ok(instructions.length > 100, "Instructions should be substantial");
    assert.ok(instructions.includes("iWork"), "Instructions should mention iWork");
  });

  it("registers all 117 tools", async () => {
    const { tools } = await ctx.client.listTools();
    assert.equal(tools.length, 117);
  });

  it("registers 50 Numbers tools", async () => {
    const { tools } = await ctx.client.listTools();
    const numbers = tools.filter((t) => t.name.startsWith("numbers_"));
    assert.equal(numbers.length, 50);
  });

  it("registers 26 Pages tools", async () => {
    const { tools } = await ctx.client.listTools();
    const pages = tools.filter((t) => t.name.startsWith("pages_"));
    assert.equal(pages.length, 26);
  });

  it("registers 41 Keynote tools", async () => {
    const { tools } = await ctx.client.listTools();
    const keynote = tools.filter((t) => t.name.startsWith("keynote_"));
    assert.equal(keynote.length, 41);
  });

  it("every tool has a description", async () => {
    const { tools } = await ctx.client.listTools();
    for (const tool of tools) {
      assert.ok(
        tool.description && tool.description.length > 0,
        `${tool.name} is missing a description`,
      );
    }
  });

  it("every tool has a valid inputSchema", async () => {
    const { tools } = await ctx.client.listTools();
    for (const tool of tools) {
      assert.equal(
        tool.inputSchema.type,
        "object",
        `${tool.name} inputSchema.type should be "object"`,
      );
    }
  });

  it("contains expected Numbers tool names", async () => {
    const { tools } = await ctx.client.listTools();
    const names = new Set(tools.map((t) => t.name));
    const expected = [
      "numbers_list_templates",
      "numbers_create_document",
      "numbers_read_cell",
      "numbers_write_cell",
      "numbers_write_table",
      "numbers_set_formula",
      "numbers_format_cells",
      "numbers_list_sheets",
      "numbers_add_sheet",
      "numbers_close_document",
      "numbers_create_sheet_with_table",
      "numbers_insert_row_at",
      "numbers_insert_column_at",
      "numbers_format_range",
      "numbers_auto_format",
      "numbers_copy_range",
      "numbers_get_cell_format",
      "numbers_move_table",
      "numbers_add_image",
      "numbers_list_images",
      "numbers_add_chart",
      "numbers_magic_fill",
      "numbers_super_resolution",
      "numbers_remove_background",
    ];
    for (const name of expected) {
      assert.ok(names.has(name), `Missing tool: ${name}`);
    }
  });

  it("contains expected Pages tool names", async () => {
    const { tools } = await ctx.client.listTools();
    const names = new Set(tools.map((t) => t.name));
    const expected = [
      "pages_list_templates",
      "pages_create_document",
      "pages_get_body_text",
      "pages_get_paragraphs",
      "pages_get_document_info",
      "pages_list_images",
      "pages_add_text",
      "pages_replace_text",
      "pages_close_document",
      "pages_create_document_with_content",
      "pages_insert_page_break",
      "pages_super_resolution",
      "pages_remove_background",
    ];
    for (const name of expected) {
      assert.ok(names.has(name), `Missing tool: ${name}`);
    }
  });

  it("contains expected Keynote tool names", async () => {
    const { tools } = await ctx.client.listTools();
    const names = new Set(tools.map((t) => t.name));
    const expected = [
      "keynote_list_themes",
      "keynote_get_theme",
      "keynote_set_theme",
      "keynote_create_presentation",
      "keynote_list_slides",
      "keynote_add_slide",
      "keynote_set_slide_title",
      "keynote_set_slide_body",
      "keynote_set_presenter_notes",
      "keynote_close_presentation",
      "keynote_position_item",
      "keynote_align_items",
      "keynote_distribute_items",
      "keynote_get_shape_info",
      "keynote_format_shape",
      "keynote_set_slide_layout",
      "keynote_add_table_to_slide",
      "keynote_read_slide_table",
      "keynote_write_slide_table",
      "keynote_format_slide_table",
      "keynote_add_line",
      "keynote_list_slide_items",
      "keynote_create_presentation_with_slides",
      "keynote_clean_up_slide",
      "keynote_super_resolution",
      "keynote_remove_background",
    ];
    for (const name of expected) {
      assert.ok(names.has(name), `Missing tool: ${name}`);
    }
  });

  // ── Annotation Tests ──

  it("every tool has annotations with openWorldHint: false", async () => {
    const { tools } = await ctx.client.listTools();
    for (const tool of tools) {
      assert.ok(
        tool.annotations,
        `${tool.name} is missing annotations`,
      );
      assert.equal(
        tool.annotations!.openWorldHint,
        false,
        `${tool.name} should have openWorldHint: false`,
      );
    }
  });

  it("read-only tools have readOnlyHint: true", async () => {
    const readOnlyTools = [
      "numbers_list_documents", "numbers_list_templates", "numbers_list_sheets", "numbers_list_tables",
      "numbers_read_table", "numbers_read_cell", "numbers_read_range", "numbers_get_table_info", "numbers_get_cell_format",
      "numbers_list_images",
      "pages_list_documents", "pages_list_templates", "pages_get_body_text", "pages_get_paragraphs", "pages_get_document_info", "pages_list_images",
      "keynote_list_presentations", "keynote_list_themes", "keynote_get_theme", "keynote_list_slides", "keynote_get_slide_content", "keynote_list_master_slides", "keynote_get_shape_info", "keynote_read_slide_table", "keynote_list_slide_items",
    ];
    const { tools } = await ctx.client.listTools();
    const byName = new Map(tools.map((t) => [t.name, t]));
    for (const name of readOnlyTools) {
      const tool = byName.get(name);
      assert.ok(tool, `Missing tool: ${name}`);
      assert.equal(
        tool!.annotations?.readOnlyHint,
        true,
        `${name} should have readOnlyHint: true`,
      );
    }
  });

  it("destructive tools have destructiveHint: true", async () => {
    const destructiveTools = [
      "numbers_close_document", "numbers_delete_sheet", "numbers_delete_table",
      "numbers_delete_row", "numbers_delete_column",
      "pages_close_document", "pages_delete_text",
      "keynote_close_presentation", "keynote_delete_slide",
    ];
    const { tools } = await ctx.client.listTools();
    const byName = new Map(tools.map((t) => [t.name, t]));
    for (const name of destructiveTools) {
      const tool = byName.get(name);
      assert.ok(tool, `Missing tool: ${name}`);
      assert.equal(
        tool!.annotations?.destructiveHint,
        true,
        `${name} should have destructiveHint: true`,
      );
    }
  });
});
