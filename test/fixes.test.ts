import { describe, it, before, after, afterEach } from "node:test";
import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import {
  envTimeout, jxaTimeout, saveTimeout, runJXA,
  DEFAULT_JXA_TIMEOUT_MS, DEFAULT_SAVE_TIMEOUT_MS,
  buildDocumentFileScript, buildSaveInPlaceScript, neverSavedError,
} from "../src/jxa.js";
import { tableDimensions } from "../src/tools/numbers.js";
import { createTestServer, type TestContext } from "./helpers/server.js";

describe("tableDimensions (numbers_create_sheet_with_table)", () => {
  it("passes through data that is already at least 2x2", () => {
    assert.deepEqual(tableDimensions([["a", "b", "c"], [1, 2, 3], [4, 5, 6]]), { rowCount: 3, columnCount: 3 });
  });

  it("pads a single row up to the 2-row minimum", () => {
    assert.deepEqual(tableDimensions([["Q1", "Q2", "Q3"]]), { rowCount: 2, columnCount: 3 });
  });

  it("pads a single column up to the 2-column minimum", () => {
    assert.deepEqual(tableDimensions([["a"], ["b"], ["c"]]), { rowCount: 3, columnCount: 2 });
  });

  it("handles a 1x1 and an empty row", () => {
    assert.deepEqual(tableDimensions([["only"]]), { rowCount: 2, columnCount: 2 });
    assert.deepEqual(tableDimensions([[]]), { rowCount: 2, columnCount: 2 });
  });

  it("uses the longest row for ragged data", () => {
    assert.deepEqual(tableDimensions([["a"], ["b", "c", "d", "e"]]), { rowCount: 2, columnCount: 4 });
  });

  it("keeps at least one body row below the header rows", () => {
    assert.deepEqual(tableDimensions([["h1", "h2"], [1, 2]], 2), { rowCount: 3, columnCount: 2 });
    assert.deepEqual(tableDimensions([["h1", "h2"]], 1), { rowCount: 2, columnCount: 2 });
    assert.deepEqual(tableDimensions([["h1", "h2"]], 4), { rowCount: 5, columnCount: 2 });
  });

  it("treats null/undefined header count as 0", () => {
    assert.deepEqual(tableDimensions([["a", "b"], [1, 2]], null), { rowCount: 2, columnCount: 2 });
    assert.deepEqual(tableDimensions([["a", "b"], [1, 2]], undefined), { rowCount: 2, columnCount: 2 });
  });
});

describe("osascript timeouts", () => {
  const saved = { t: process.env.IWORK_MCP_TIMEOUT_MS, s: process.env.IWORK_MCP_SAVE_TIMEOUT_MS };
  afterEach(() => {
    for (const [k, v] of [["IWORK_MCP_TIMEOUT_MS", saved.t], ["IWORK_MCP_SAVE_TIMEOUT_MS", saved.s]] as const) {
      if (v === undefined) delete process.env[k]; else process.env[k] = v;
    }
  });

  it("envTimeout falls back when unset or invalid", () => {
    delete process.env.IWORK_MCP_TEST_TIMEOUT;
    assert.equal(envTimeout("IWORK_MCP_TEST_TIMEOUT", 123), 123);
    for (const bad of ["", "abc", "0", "-5", "NaN"]) {
      process.env.IWORK_MCP_TEST_TIMEOUT = bad;
      assert.equal(envTimeout("IWORK_MCP_TEST_TIMEOUT", 123), 123, `value ${JSON.stringify(bad)}`);
    }
    delete process.env.IWORK_MCP_TEST_TIMEOUT;
  });

  it("envTimeout reads a positive value", () => {
    process.env.IWORK_MCP_TEST_TIMEOUT = "4500.7";
    assert.equal(envTimeout("IWORK_MCP_TEST_TIMEOUT", 123), 4500);
    delete process.env.IWORK_MCP_TEST_TIMEOUT;
  });

  it("save timeout defaults well above the ordinary JXA timeout", () => {
    delete process.env.IWORK_MCP_TIMEOUT_MS;
    delete process.env.IWORK_MCP_SAVE_TIMEOUT_MS;
    assert.equal(jxaTimeout(), DEFAULT_JXA_TIMEOUT_MS);
    assert.equal(saveTimeout(), DEFAULT_SAVE_TIMEOUT_MS);
    assert.ok(DEFAULT_SAVE_TIMEOUT_MS >= 4 * DEFAULT_JXA_TIMEOUT_MS);
  });

  it("timeouts are configurable via env vars", () => {
    process.env.IWORK_MCP_TIMEOUT_MS = "45000";
    process.env.IWORK_MCP_SAVE_TIMEOUT_MS = "600000";
    assert.equal(jxaTimeout(), 45_000);
    assert.equal(saveTimeout(), 600_000);
  });

  it("reports a killed osascript as a timeout, not a bare 'exit null'", { skip: !existsSync("/usr/bin/osascript") }, async () => {
    await assert.rejects(
      runJXA("delay(10); return 1;", undefined, { timeout: 300, label: "slowScript" }),
      (err: Error) => {
        assert.match(err.message, /slowScript timed out after 300ms/);
        assert.match(err.message, /IWORK_MCP_SAVE_TIMEOUT_MS/);
        return true;
      },
    );
  });
});

describe("pages_add_image input validation", () => {
  let ctx: TestContext;
  before(async () => { ctx = await createTestServer(); });
  after(async () => { await ctx.cleanup(); });

  it("rejects a missing image file before calling Pages", async () => {
    const result = await ctx.client.callTool({
      name: "pages_add_image",
      arguments: { documentName: "Nope", filePath: "/nonexistent/iwork-mcp-test.png" },
    });
    const text = (result.content as Array<{ type: string; text: string }>)[0].text;
    assert.equal(result.isError, true);
    assert.match(text, /Image file not found/);
  });
});

// Runs a save script against stubbed JXA globals, recording what it would save.
function runSaveScript(script: string, doc: { name: string; file: string | null }) {
  const calls = { jxaSave: 0, appleScript: [] as string[] };
  const fakeDoc = {
    name: () => doc.name,
    file: () => (doc.file ? { toString: () => doc.file } : null),
    save: () => { calls.jxaSave++; },
  };
  const Application = () => ({ documents: { byName: () => fakeDoc } });
  const $ = Object.assign(() => ({}), {
    NSAppleScript: { alloc: { initWithSource: (src: string) => {
      calls.appleScript.push(src);
      return { executeAndReturnError: () => ({ isNil: () => false }) };
    } } },
  });
  const ObjC = { import: () => {}, deepUnwrap: () => ({}) };
  const out = new Function("params", "Application", "$", "ObjC", script)({ documentName: doc.name }, Application, $, ObjC);
  return { result: out === undefined ? undefined : JSON.parse(out), calls };
}

describe("save without a path (numbers/pages/keynote save tools)", () => {
  const apps = ["Numbers", "Pages", "Keynote"] as const;

  it("never-saved error names the document and asks for filePath", () => {
    const cases = { Numbers: ["numbers_save_document", ".numbers"], Pages: ["pages_save_document", ".pages"], Keynote: ["keynote_save_presentation", ".key"] };
    for (const app of apps) {
      const msg = neverSavedError(app, "Untitled 3");
      assert.match(msg, /"Untitled 3" has never been saved/);
      assert.ok(msg.includes(cases[app][0]), msg);
      assert.ok(msg.includes("filePath"), msg);
      assert.ok(msg.includes(cases[app][1]), msg);
    }
  });

  for (const app of apps) {
    it(`${app}: file lookup reports the document's path, or null if it has never been saved`, () => {
      const saved = runSaveScript(buildDocumentFileScript(app), { name: "Budget", file: "/tmp/x/Budget" });
      assert.deepEqual(saved.result, { name: "Budget", path: "/tmp/x/Budget" });
      const unsaved = runSaveScript(buildDocumentFileScript(app), { name: "Untitled 4", file: null });
      assert.deepEqual(unsaved.result, { name: "Untitled 4", path: null });
      for (const r of [saved, unsaved]) {
        assert.equal(r.calls.jxaSave, 0, "the lookup must not save");
        assert.equal(r.calls.appleScript.length, 0, "the lookup must not save");
      }
    });

    it(`${app} (Creator Studio): saves via AppleScript save document, not JXA save`, () => {
      const { calls } = runSaveScript(buildSaveInPlaceScript(app, true), { name: "Budget", file: "/tmp/x/Budget" });
      assert.equal(calls.jxaSave, 0);
      assert.deepEqual(calls.appleScript, [`tell application "${app}" to save document "Budget"`]);
    });

    it(`${app} (standard): saves via JXA doc.save()`, () => {
      const { calls } = runSaveScript(buildSaveInPlaceScript(app, false), { name: "Budget", file: "/tmp/x/Budget" });
      assert.equal(calls.jxaSave, 1);
      assert.equal(calls.appleScript.length, 0);
    });

    it(`${app}: quotes document names in the AppleScript save`, () => {
      const { calls } = runSaveScript(buildSaveInPlaceScript(app, true), { name: 'Q1 "final" \\ v2', file: "/tmp/q" });
      assert.deepEqual(calls.appleScript, [`tell application "${app}" to save document "Q1 \\"final\\" \\\\ v2"`]);
    });
  }
});
