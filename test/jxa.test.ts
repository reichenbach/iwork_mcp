import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { OsascriptError, injectDocumentNameResolution, isCreatorStudio } from "../src/jxa.js";

describe("OsascriptError", () => {
  it("is an instance of Error", () => {
    const err = new OsascriptError("some error (-1743)", 1);
    assert.ok(err instanceof Error);
  });

  it("has name 'OsascriptError'", () => {
    const err = new OsascriptError("some error (-1743)", 1);
    assert.equal(err.name, "OsascriptError");
  });

  it("parses error code -1743 (permission denied)", () => {
    const err = new OsascriptError("Error: Error: some action (-1743)", 1);
    assert.equal(err.appleScriptErrorCode, -1743);
    assert.ok(err.message.includes("Permission denied"));
    assert.ok(err.message.includes("Automation"));
  });

  it("parses error code -1728 (element not found)", () => {
    const err = new OsascriptError("execution error: can't get document (-1728)", 1);
    assert.equal(err.appleScriptErrorCode, -1728);
    assert.ok(err.message.includes("Element not found"));
  });

  it("parses error code -128 (user cancelled)", () => {
    const err = new OsascriptError("execution error: User cancelled (-128)", 1);
    assert.equal(err.appleScriptErrorCode, -128);
    assert.ok(err.message.includes("cancelled"));
  });

  it("parses error code -10810 (app not running)", () => {
    const err = new OsascriptError("error (-10810)", 1);
    assert.equal(err.appleScriptErrorCode, -10810);
    assert.ok(err.message.includes("not running"));
  });

  it("parses error code -1700 (invalid data type)", () => {
    const err = new OsascriptError("error (-1700)", 1);
    assert.equal(err.appleScriptErrorCode, -1700);
    assert.ok(err.message.includes("Invalid data type"));
  });

  it("parses error code -1708 (does not understand)", () => {
    const err = new OsascriptError("error (-1708)", 1);
    assert.equal(err.appleScriptErrorCode, -1708);
    assert.ok(err.message.includes("does not understand"));
  });

  it("handles unknown error codes gracefully", () => {
    const err = new OsascriptError("error (-99999)", 1);
    assert.equal(err.appleScriptErrorCode, -99999);
    assert.ok(err.message.includes("osascript failed"));
    assert.ok(err.message.includes("-99999"));
  });

  it("handles stderr with no error code", () => {
    const err = new OsascriptError("something went wrong", 1);
    assert.equal(err.appleScriptErrorCode, undefined);
    assert.ok(err.message.includes("osascript failed"));
    assert.ok(err.message.includes("something went wrong"));
  });

  it("preserves stderr in the error", () => {
    const stderr = "execution error: blah (-1728)";
    const err = new OsascriptError(stderr, 1);
    assert.equal(err.stderr, stderr);
  });
});

// Helper: build a script similar to what runJXA produces after rewriteAppNames
function fakeScript(appName: string, usesDocName: boolean): string {
  const paramsLine = "const params = JSON.parse(argv[0]);";
  const body = usesDocName
    ? `const app = Application("${appName}"); const doc = app.documents.byName(params.documentName);`
    : `const app = Application("${appName}"); return app.documents().length;`;
  return `\nfunction run(argv) {\n  ${paramsLine}\n  ${body}\n}\n`;
}

describe("injectDocumentNameResolution", () => {
  it("injects resolution block for Numbers Creator Studio", () => {
    const script = fakeScript("Numbers Creator Studio", true);
    const result = injectDocumentNameResolution(script);
    assert.ok(result.includes("[iwork-mcp] Creator Studio document name resolution"));
    assert.ok(result.includes('Application("Numbers Creator Studio")'));
    assert.ok(result.includes('".numbers"'));
  });

  it("injects resolution block for Pages Creator Studio", () => {
    const script = fakeScript("Pages Creator Studio", true);
    const result = injectDocumentNameResolution(script);
    assert.ok(result.includes("[iwork-mcp] Creator Studio document name resolution"));
    assert.ok(result.includes('".pages"'));
  });

  it("injects resolution block for Keynote Creator Studio", () => {
    const script = fakeScript("Keynote Creator Studio", true);
    const result = injectDocumentNameResolution(script);
    assert.ok(result.includes("[iwork-mcp] Creator Studio document name resolution"));
    assert.ok(result.includes('".key"'));
  });

  it("does not inject when params.documentName is absent", () => {
    const script = fakeScript("Numbers Creator Studio", false);
    const result = injectDocumentNameResolution(script);
    assert.equal(result, script);
  });

  it("does not inject when no recognizable app is in the script", () => {
    const script = `\nfunction run(argv) {\n  const params = JSON.parse(argv[0]);\n  const app = Application("Safari"); app.documents.byName(params.documentName);\n}\n`;
    const result = injectDocumentNameResolution(script);
    assert.equal(result, script);
  });

  it("does not inject when the params anchor line is missing", () => {
    const script = `\nfunction run(argv) {\n  const app = Application("Numbers Creator Studio"); app.documents.byName(params.documentName);\n}\n`;
    const result = injectDocumentNameResolution(script);
    assert.equal(result, script);
  });

  it("injected code is syntactically valid JavaScript", () => {
    const script = fakeScript("Numbers Creator Studio", true);
    const result = injectDocumentNameResolution(script);
    // Should not throw when parsed as a function
    assert.doesNotThrow(() => new Function(result));
  });

  it("refuses names that match several open documents", () => {
    const script = fakeScript("Pages Creator Studio", true);
    const result = injectDocumentNameResolution(script);
    assert.ok(result.includes("Several open documents match the name"));
    assert.doesNotThrow(() => new Function(result));
  });

  it("injects after the params line, before the script body", () => {
    const script = fakeScript("Numbers Creator Studio", true);
    const result = injectDocumentNameResolution(script);
    const paramsIdx = result.indexOf("const params = JSON.parse(argv[0]);");
    const injectionIdx = result.indexOf("[iwork-mcp] Creator Studio document name resolution");
    const bodyIdx = result.indexOf("const doc = app.documents.byName(params.documentName)");
    assert.ok(paramsIdx < injectionIdx, "injection should come after params line");
    assert.ok(injectionIdx < bodyIdx, "injection should come before script body");
  });

  it("isCreatorStudio returns a boolean", () => {
    const result = isCreatorStudio();
    assert.equal(typeof result, "boolean");
  });

  it("uses correct extension per app", () => {
    const cases: Array<[string, string]> = [
      ["Numbers Creator Studio", ".numbers"],
      ["Pages Creator Studio", ".pages"],
      ["Keynote Creator Studio", ".key"],
    ];
    for (const [appName, ext] of cases) {
      const script = fakeScript(appName, true);
      const result = injectDocumentNameResolution(script);
      assert.ok(
        result.includes(`params.documentName + ${JSON.stringify(ext)}`),
        `expected extension ${ext} for ${appName}`,
      );
    }
  });
});
