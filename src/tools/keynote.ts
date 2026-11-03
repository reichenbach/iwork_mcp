import { execFile } from "node:child_process";
import { existsSync, realpathSync } from "node:fs";
import { copyFile, mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { promisify } from "node:util";
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runJXA, OsascriptError, isCreatorStudio, creatorStudioSaveAs, creatorStudioKeynoteExportPDF, rememberCreatedDocument, untitledAutoSavePath, trashUntitledAutoSave, clickMenuItem, resolveAppName } from "../jxa.js";
import { ANNOTATIONS } from "../annotations.js";

const execFileAsync = promisify(execFile);

function rememberCreated<T extends { name: string }>(result: T): T {
  rememberCreatedDocument("Keynote", result.name.replace(/\.key$/, ""));
  return result;
}

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

// AppleScript bridge for what JXA can't do (master slides, "after slide N").
// A failed call returns a truthy nil wrapper, hence isNil(); Ref() for the error crashes osascript.
const KEYNOTE_AS_HELPERS = `
    ObjC.import("Foundation");
    function asQuote(s) {
      return '"' + String(s).replace(/\\\\/g, "\\\\\\\\").replace(/"/g, '\\\\"') + '"';
    }
    function runAS(src) {
      const script = $.NSAppleScript.alloc.initWithSource(src);
      const err = $();
      const result = script.executeAndReturnError(err);
      if (!result || (result.isNil && result.isNil())) {
        let info = {};
        try { info = ObjC.deepUnwrap(err) || {}; } catch (e) {}
        const msg = info.NSAppleScriptErrorBriefMessage || info.NSAppleScriptErrorMessage || "AppleScript execution failed";
        throw new Error(info.NSAppleScriptErrorNumber !== undefined ? msg + " (" + info.NSAppleScriptErrorNumber + ")" : msg);
      }
      return result;
    }
    function docRef() { return "document " + asQuote(params.documentName); }
`;

const SLIDE_SIZE_SCHEMA = {
  width: z.number().int().positive().optional().describe("Slide width in points (e.g. 1920 for 16:9, 1024 for 4:3). Default: theme default. Theme layouts are not rescaled, so placeholder text keeps its point size"),
  height: z.number().int().positive().optional().describe("Slide height in points (e.g. 1080 for 16:9, 768 for 4:3). Default: theme default"),
};

const APPLY_SLIDE_SIZE = `
      if (params.width !== null) doc.width = params.width;
      if (params.height !== null) doc.height = params.height;
      if ((params.width !== null && doc.width() !== params.width) || (params.height !== null && doc.height() !== params.height)) {
        throw new Error("Keynote did not apply the slide size (document " + doc.name() + " is " + doc.width() + "x" + doc.height() + ")");
      }
`;

// Fails on a missing document or slide instead of reading an empty slide
const SLIDE_LOOKUP = `
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slideCount = doc.slides.length;
      if (params.slideNumber > slideCount) {
        throw new Error("Slide " + params.slideNumber + " is out of range (presentation has " + slideCount + " slides)");
      }
      const slide = doc.slides[params.slideNumber - 1];
`;

// Placeholders show up in both textItems() and shapes(); same text + geometry = same box
const ITEM_KEY_HELPER = `
      function itemKey(item, text) {
        try {
          const p = item.position();
          return [text, Math.round(p.x), Math.round(p.y), Math.round(item.width()), Math.round(item.height())].join("|");
        } catch (e) { return null; }
      }
`;

const TEXT_ITEM_TYPE = z.enum(["shape", "textItem"]).optional().describe("Item type: 'shape' (default) or 'textItem' (text box); indexes as in keynote_list_slide_items");

// Expects slide, params.itemType and params.itemIndex. Locked items ignore deletes and fail edits with -10003.
const ITEM_LOOKUP = `
      const collections = { shape: slide.shapes, textItem: slide.textItems, image: slide.images, table: slide.tables, line: slide.lines, chart: slide.charts };
      const itemList = collections[params.itemType]();
      if (params.itemIndex >= itemList.length) {
        throw new Error(params.itemType + " index " + params.itemIndex + " out of range (slide has " + itemList.length + ")");
      }
      const item = itemList[params.itemIndex];
      function assertUnlocked() {
        let locked = false;
        try { locked = item.locked(); } catch (e) {}
        if (locked) throw new Error(params.itemType + " " + params.itemIndex + " is locked; unlock it first with keynote_set_item_locked");
      }
`;

// Layouts like Blank/Quote hide the title/body placeholder; text written there is invisible
const PLACEHOLDER_HELPER = `
      function showPlaceholder(slide, slideNumber, prop, label) {
        let showing;
        try { showing = slide[prop](); } catch (e) { return false; }
        if (showing) return false;
        try { slide[prop] = true; } catch (e) {}
        if (!slide[prop]()) {
          throw new Error("The layout of slide " + slideNumber + " has no visible " + label + " placeholder. Change the layout (keynote_set_slide_layout) or add a text shape instead.");
        }
        return true;
      }
`;

function normalizeTransitionEffect(effect: string): string {
  const e = effect.trim().toLowerCase();
  return e === "none" ? "no transition effect" : e;
}

// Valid effect names from Keynote's scripting dictionary, for error messages
let transitionEffectNames: string[] | undefined;
async function keynoteTransitionEffects(): Promise<string[]> {
  if (transitionEffectNames) return transitionEffectNames;
  let names: string[] = [];
  try {
    const { stdout } = await execFileAsync("/usr/bin/sdef", [`/Applications/${resolveAppName("Keynote")}.app`], { maxBuffer: 20 * 1024 * 1024 });
    for (const m of stdout.matchAll(/<enumeration\b[^>]*>([\s\S]*?)<\/enumeration>/g)) {
      const found = [...m[1].matchAll(/<enumerator\b[^>]*\bname="([^"]+)"/g)].map(x => x[1]);
      if (found.includes("dissolve") && found.includes("magic move")) { names = found; break; }
    }
  } catch {}
  transitionEffectNames = names;
  return names;
}

export function registerKeynoteTools(server: McpServer): void {
  // ── Presentation Management ──

  server.tool(
    "keynote_list_presentations",
    "List all open Keynote presentations",
    {},
    ANNOTATIONS.readOnly,
    async () => handleJXA(() => runJXA<string[]>(`
      const app = Application("Keynote");
      const docs = app.documents();
      return JSON.stringify(docs.map(d => ({ name: d.name(), path: d.file() ? d.file().toString() : null })));
    `)),
  );

  server.tool(
    "keynote_list_themes",
    "List all available Keynote themes (e.g. White, Black, Gradient)",
    {},
    ANNOTATIONS.readOnly,
    async () => handleJXA(() => runJXA<string[]>(`
      const app = Application("Keynote");
      return JSON.stringify(app.themes().map(t => t.name()));
    `)),
  );

  server.tool(
    "keynote_get_theme",
    "Get the current theme of a Keynote presentation",
    {
      documentName: z.string().describe("Name of the open presentation"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      return JSON.stringify({ theme: doc.documentTheme().name() });
    `, { documentName })),
  );

  server.tool(
    "keynote_set_theme",
    "Change the theme of an existing Keynote presentation (use keynote_list_themes to see available themes). Warning: Keynote resets custom text formatting (font, size, color) on existing slides to the new theme's defaults, and this can't be undone via scripting — on a finished deck, work on a copy.",
    {
      documentName: z.string().describe("Name of the open presentation"),
      themeName: z.string().describe("Theme name to apply (e.g. 'White', 'Black', 'Gradient')"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, themeName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      doc.documentTheme = app.themes.byName(params.themeName);
      return JSON.stringify({ themeSet: true, theme: doc.documentTheme().name() });
    `, { documentName, themeName })),
  );

  server.tool(
    "keynote_create_presentation",
    "Create a new Keynote presentation (optionally with a theme — use keynote_list_themes to see available themes — and a slide size)",
    {
      themeName: z.string().optional().describe("Theme name (optional, e.g. 'White', 'Black', 'Gradient')"),
      ...SLIDE_SIZE_SCHEMA,
    },
    ANNOTATIONS.readWrite,
    async ({ themeName, width, height }) => handleJXA(async () => rememberCreated(await runJXA<{ name: string }>(`
      const app = Application("Keynote");
      let doc;
      if (params.themeName) {
        doc = app.Document({ documentTheme: app.themes[params.themeName] });
        app.documents.push(doc);
      } else {
        doc = app.Document();
        app.documents.push(doc);
      }
      ${APPLY_SLIDE_SIZE}
      return JSON.stringify({ name: doc.name(), slideCount: doc.slides.length, width: doc.width(), height: doc.height() });
    `, { themeName: themeName ?? null, width: width ?? null, height: height ?? null }))),
  );

  server.tool(
    "keynote_open_presentation",
    "Open a .key file from disk",
    {
      filePath: z.string().startsWith("/").describe("Absolute path to the .key file"),
    },
    ANNOTATIONS.readWrite,
    async ({ filePath }) => handleJXA(async () => {
      try {
        return await runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.open(Path(params.filePath));
      return JSON.stringify({ name: doc.name(), slideCount: doc.slides.length });
    `, { filePath });
      } catch (err) {
        // The first open from a folder can time out although the document opens
        const timedOut = err instanceof OsascriptError && (err.appleScriptErrorCode === undefined || err.appleScriptErrorCode === -1712);
        if (!timedOut || !existsSync(filePath)) throw err;
        const paths = [filePath];
        try { paths.push(realpathSync(filePath)); } catch {}
        const opened = await runJXA<unknown>(`
      const app = Application("Keynote");
      for (let i = 0; i < 30; i++) {
        const docs = app.documents();
        for (let j = 0; j < docs.length; j++) {
          try {
            const f = docs[j].file();
            if (f && params.paths.indexOf(f.toString().replace(/\\/$/, "")) >= 0) {
              return JSON.stringify({ name: docs[j].name(), slideCount: docs[j].slides.length });
            }
          } catch (e) {}
        }
        delay(1);
      }
      return JSON.stringify(null);
    `, { paths: paths.map(p => p.replace(/\/$/, "")) }, { timeout: 45_000, label: "keynote_open_presentation:wait" });
        if (opened) return opened;
        throw err;
      }
    }),
  );

  server.tool(
    "keynote_save_presentation",
    "Save a Keynote presentation as .key (use this to save to disk — use keynote_export_presentation for PDF/PowerPoint/HTML)",
    {
      documentName: z.string().describe("Name of the open presentation"),
      filePath: z.string().startsWith("/").optional().describe("File path to save to (for Save As)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, filePath }) => handleJXA(async () => {
      if (isCreatorStudio()) {
        if (filePath) {
          const newName = await creatorStudioSaveAs("Keynote", documentName, filePath);
          return JSON.stringify({ saved: true, name: newName });
        }
        return runJXA<string>(`
      ${KEYNOTE_AS_HELPERS}
      const app = Application("Keynote");
      const name = app.documents.byName(params.documentName).name();
      runAS('tell application "Keynote" to save ' + docRef());
      return JSON.stringify({ saved: true, name: name });
    `, { documentName });
      }
      return runJXA<string>(`
      const app = Application("Keynote");
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
    "keynote_export_presentation",
    "Export a Keynote presentation to a different format: PDF, PowerPoint (.pptx), HTML, or images (not .key — use keynote_save_presentation for that)",
    {
      documentName: z.string().describe("Name of the open presentation"),
      filePath: z.string().startsWith("/").describe("Absolute path for the exported file"),
      format: z.enum(["PDF", "PowerPoint", "HTML", "images"]).describe("Export format"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, filePath, format }) => handleJXA(async () => {
      if (isCreatorStudio() && format === "PDF") {
        const { method, pageCount } = await creatorStudioKeynoteExportPDF(documentName, filePath);
        return JSON.stringify({ exported: true, path: filePath, format: "PDF", method, pageCount });
      }
      return runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const formatMap = {
        "PDF": "PDF",
        "PowerPoint": "Microsoft PowerPoint",
        "HTML": "HTML",
        "images": "slide images",
      };
      const fmt = formatMap[params.format];
      app.export(doc, { to: Path(params.filePath), as: fmt });
      return JSON.stringify({ exported: true, path: params.filePath, format: params.format });
    `, { documentName, filePath, format });
    }),
  );

  server.tool(
    "keynote_export_slide_image",
    "Render a single slide as a PNG and return the image, e.g. to check a slide after editing it",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      filePath: z.string().startsWith("/").regex(/\.png$/i).optional().describe("Where to save the PNG (default: a temporary file)"),
      maxSize: z.number().int().min(64).optional().describe("Longest side in pixels; larger images are scaled down (default: 1280)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, filePath, maxSize }) => {
      try {
        const workDir = await mkdtemp(join(filePath ? dirname(filePath) : tmpdir(), ".iwork-mcp-slide-"));
        try {
          // Keynote only exports whole decks, so skip every other slide for the export
          await runJXA<void>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slides = doc.slides();
      if (params.slideNumber > slides.length) {
        throw new Error("Slide " + params.slideNumber + " is out of range (presentation has " + slides.length + " slides)");
      }
      const wasSkipped = slides.map(function (s) { return s.skipped(); });
      try {
        for (let i = 0; i < slides.length; i++) {
          const skip = i !== params.slideNumber - 1;
          if (wasSkipped[i] !== skip) slides[i].skipped = skip;
        }
        app.export(doc, { to: Path(params.folder), as: "slide images", withProperties: { skippedSlides: false } });
      } finally {
        for (let i = 0; i < slides.length; i++) {
          if (slides[i].skipped() !== wasSkipped[i]) slides[i].skipped = wasSkipped[i];
        }
      }
    `, { documentName, slideNumber, folder: join(workDir, "slide") }, { timeout: 60_000, label: "keynote_export_slide_image" });

          const { stdout } = await execFileAsync("/usr/bin/find", [workDir, "-type", "f", "-iname", "*.png"]);
          const image = stdout.split("\n").filter(Boolean)[0];
          if (!image) throw new Error("Keynote did not export the slide image");
          const limit = maxSize ?? 1280;
          const dims = await execFileAsync("/usr/bin/sips", ["-g", "pixelWidth", "-g", "pixelHeight", image]);
          const sizes = [...dims.stdout.matchAll(/pixel(?:Width|Height): (\d+)/g)].map(m => Number(m[1]));
          if (Math.max(...sizes) > limit) await execFileAsync("/usr/bin/sips", ["-Z", String(limit), image]);

          const target = filePath ?? join(tmpdir(), `keynote-slide-${slideNumber}-${Date.now()}.png`);
          await copyFile(image, target);
          const data = (await readFile(target)).toString("base64");
          return {
            content: [
              { type: "image" as const, data, mimeType: "image/png" },
              { type: "text" as const, text: JSON.stringify({ exported: true, slideNumber, path: target }) },
            ],
          };
        } finally {
          await rm(workDir, { recursive: true, force: true });
        }
      } catch (err) {
        return toolResult(err instanceof Error ? err.message : String(err), true);
      }
    },
  );

  server.tool(
    "keynote_close_presentation",
    "Close a Keynote presentation",
    {
      documentName: z.string().describe("Name of the open presentation"),
      saving: z.enum(["yes", "no", "ask"]).optional().describe("Whether to save before closing"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, saving }) => handleJXA(async () => {
      // Creator Studio keeps the iCloud auto-save of an untitled document even when closed without saving
      const autoSave = isCreatorStudio() && saving === "no" ? await untitledAutoSavePath("Keynote", documentName) : null;
      const result = await runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const saveOpts = { yes: "yes", no: "no", ask: "ask" };
      if (params.saving) {
        doc.close({ saving: saveOpts[params.saving] });
      } else {
        doc.close();
      }
      return JSON.stringify({ closed: true });
    `, { documentName, saving: saving ?? null });
      if (autoSave) await trashUntitledAutoSave("Keynote", documentName, autoSave);
      return result;
    }),
  );

  // ── Slide Tools ──

  server.tool(
    "keynote_list_slides",
    "List all slides in a presentation with their titles (plus the first visible text when a slide has no visible title)",
    {
      documentName: z.string().describe("Name of the open presentation"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slides = doc.slides();
      return JSON.stringify(slides.map((s, i) => {
        let title = "";
        try {
          const titleItem = s.defaultTitleItem();
          if (titleItem) title = titleItem.objectText();
        } catch(e) {}
        const entry = { index: i, slideNumber: i + 1, title: title, skipped: s.skipped() };
        let titleShowing = true;
        try { titleShowing = s.titleShowing(); } catch(e) {}
        if (!titleShowing) entry.titleShowing = false;
        // Decks built from free text boxes have no title to tell slides apart
        if (!title || !titleShowing) {
          try {
            const boxes = s.textItems().concat(s.shapes());
            for (let j = 0; j < boxes.length; j++) {
              let t = "";
              try { t = boxes[j].objectText(); } catch(e) {}
              if (t && t.trim()) { entry.firstText = t.trim().split("\\n")[0].slice(0, 80); break; }
            }
          } catch(e) {}
        }
        return entry;
      }));
    `, { documentName })),
  );

  server.tool(
    "keynote_get_slide_content",
    "Read all content from a slide: title, body, presenter notes, and list of items",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, slideNumber }) => handleJXA(() => runJXA<string>(`
      ${SLIDE_LOOKUP}
      ${ITEM_KEY_HELPER}

      let title = "";
      try { const t = slide.defaultTitleItem(); if (t) title = t.objectText(); } catch(e) {}
      let body = "";
      try { const b = slide.defaultBodyItem(); if (b) body = b.objectText(); } catch(e) {}
      let notes = "";
      try { notes = slide.presenterNotes(); } catch(e) {}

      const seen = {};
      function isRepeat(item, text) {
        const key = itemKey(item, text);
        if (key === null) return false;
        if (seen[key]) return true;
        seen[key] = true;
        return false;
      }

      const textItems = [];
      try {
        const items = slide.textItems();
        for (let i = 0; i < items.length; i++) {
          const text = items[i].objectText();
          if (isRepeat(items[i], text)) continue;
          textItems.push({ index: i, text: text, width: items[i].width(), height: items[i].height() });
        }
      } catch(e) {}

      const images = [];
      try {
        const imgs = slide.images();
        for (let i = 0; i < imgs.length; i++) {
          images.push({ index: i, width: imgs[i].width(), height: imgs[i].height() });
        }
      } catch(e) {}

      const shapes = [];
      try {
        const shps = slide.shapes();
        for (let i = 0; i < shps.length; i++) {
          const text = shps[i].objectText();
          if (isRepeat(shps[i], text)) continue;
          shapes.push({ index: i, text: text, width: shps[i].width(), height: shps[i].height() });
        }
      } catch(e) {}

      return JSON.stringify({ slideNumber: params.slideNumber, title, body, presenterNotes: notes, textItems, images, shapes });
    `, { documentName, slideNumber })),
  );

  server.tool(
    "keynote_list_slide_items",
    "List all items on a slide (shapes, text boxes, images, tables, lines, charts) with their types, positions, sizes, and whether they are locked",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, slideNumber }) => handleJXA(() => runJXA<string>(`
      ${SLIDE_LOOKUP}
      ${ITEM_KEY_HELPER}
      const items = [];
      const byKey = {};

      function pos(item) {
        try { var p = item.position(); return { x: p.x || p[0] || 0, y: p.y || p[1] || 0 }; } catch(e) { return null; }
      }
      function markLocked(item, entry) {
        try { if (item.locked()) entry.locked = true; } catch(e) {}
        return entry;
      }

      try {
        var shapes = slide.shapes();
        for (var i = 0; i < shapes.length; i++) {
          var s = shapes[i];
          var entry = markLocked(s, { type: "shape", index: i, width: s.width(), height: s.height(), position: pos(s) });
          try { entry.text = s.objectText(); } catch(e) {}
          var key = itemKey(s, entry.text);
          if (key !== null) {
            if (byKey[key]) continue;
            byKey[key] = entry;
          }
          items.push(entry);
        }
      } catch(e) {}

      try {
        var textItems = slide.textItems();
        for (var i = 0; i < textItems.length; i++) {
          var t = textItems[i];
          var entry = markLocked(t, { type: "textItem", index: i, width: t.width(), height: t.height(), position: pos(t) });
          try { entry.text = t.objectText(); } catch(e) {}
          var key = itemKey(t, entry.text);
          if (key !== null) {
            var same = byKey[key];
            if (same) {
              if (same.type === "shape" && same.textItemIndex === undefined) same.textItemIndex = i;
              continue;
            }
            byKey[key] = entry;
          }
          items.push(entry);
        }
      } catch(e) {}

      try {
        var images = slide.images();
        for (var i = 0; i < images.length; i++) {
          var img = images[i];
          items.push(markLocked(img, { type: "image", index: i, width: img.width(), height: img.height(), position: pos(img) }));
        }
      } catch(e) {}

      try {
        var tables = slide.tables();
        for (var i = 0; i < tables.length; i++) {
          var tbl = tables[i];
          items.push(markLocked(tbl, { type: "table", index: i, name: tbl.name(), rows: tbl.rowCount(), columns: tbl.columnCount(), width: tbl.width(), height: tbl.height(), position: pos(tbl) }));
        }
      } catch(e) {}

      try {
        var lines = slide.lines();
        for (var i = 0; i < lines.length; i++) {
          var ln = lines[i];
          items.push(markLocked(ln, { type: "line", index: i, startPoint: ln.startPoint(), endPoint: ln.endPoint() }));
        }
      } catch(e) {}

      try {
        var charts = slide.charts();
        for (var i = 0; i < charts.length; i++) {
          var ch = charts[i];
          items.push(markLocked(ch, { type: "chart", index: i, width: ch.width(), height: ch.height(), position: pos(ch) }));
        }
      } catch(e) {}

      return JSON.stringify({ slideNumber: params.slideNumber, itemCount: items.length, items: items });
    `, { documentName, slideNumber })),
  );

  server.tool(
    "keynote_delete_slide",
    "Delete a slide from the presentation",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number to delete (1-based)"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, slideNumber }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      app.delete(doc.slides[params.slideNumber - 1]);
      return JSON.stringify({ deleted: true, slideNumber: params.slideNumber, remainingSlides: doc.slides.length });
    `, { documentName, slideNumber })),
  );

  server.tool(
    "keynote_duplicate_slide",
    "Duplicate an existing slide, optionally placing the copy at a given position",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number to duplicate (1-based)"),
      toSlideNumber: z.number().int().min(1).optional().describe("Slide number the copy should end up at (1-based, up to slide count + 1). Default: end of presentation."),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, toSlideNumber }) => handleJXA(() => runJXA<string>(`
      ${KEYNOTE_AS_HELPERS}
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const before = doc.slides.length;
      if (params.slideNumber > before) {
        throw new Error("Slide " + params.slideNumber + " is out of range (presentation has " + before + " slides)");
      }
      let target;
      if (params.toSlideNumber === null) {
        app.duplicate(doc.slides[params.slideNumber - 1]);
        target = before + 1;
      } else {
        target = params.toSlideNumber;
        if (target > before + 1) {
          throw new Error("toSlideNumber " + target + " is out of range (max " + (before + 1) + ")");
        }
        const loc = target === 1 ? 'before slide 1' : 'after slide ' + (target - 1);
        runAS('tell application "Keynote" to tell ' + docRef() + ' to duplicate slide ' + params.slideNumber + ' to ' + loc);
      }
      const total = doc.slides.length;
      if (total !== before + 1) {
        throw new Error("Slide was not duplicated (slide count " + before + " -> " + total + ")");
      }
      return JSON.stringify({ duplicated: true, slideNumber: target, totalSlides: total });
    `, { documentName, slideNumber, toSlideNumber: toSlideNumber ?? null })),
  );

  server.tool(
    "keynote_list_master_slides",
    "List all available master slide layouts in the current theme",
    {
      documentName: z.string().describe("Name of the open presentation"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName }) => handleJXA(() => runJXA<string>(`
      ${KEYNOTE_AS_HELPERS}
      const result = runAS('tell application "Keynote" to tell ' + docRef() + ' to return name of every master slide');
      const count = result.numberOfItems;
      const names = [];
      for (let i = 1; i <= count; i++) {
        names.push({ index: i - 1, name: result.descriptorAtIndex(i).stringValue.js });
      }
      return JSON.stringify(names);
    `, { documentName })),
  );

  server.tool(
    "keynote_reorder_slide",
    "Move a slide from one position to another",
    {
      documentName: z.string().describe("Name of the open presentation"),
      fromSlideNumber: z.number().int().min(1).describe("Current slide number (1-based)"),
      toSlideNumber: z.number().int().min(1).describe("Target slide number (1-based)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, fromSlideNumber, toSlideNumber }) => handleJXA(() => runJXA<string>(`
      ${KEYNOTE_AS_HELPERS}
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const total = doc.slides.length;
      const from = params.fromSlideNumber;
      const to = params.toSlideNumber;
      if (from > total || to > total) {
        throw new Error("Slide number out of range (presentation has " + total + " slides)");
      }
      if (from !== to) {
        // "move ... to slide N" replaces slide N
        runAS('tell application "Keynote" to tell ' + docRef() + ' to move slide ' + from +
          ' to ' + (to < from ? 'before' : 'after') + ' slide ' + to);
        const after = doc.slides.length;
        if (after !== total) {
          throw new Error("Slide count changed during move (" + total + " -> " + after + ")");
        }
      }
      return JSON.stringify({ moved: true, from: from, to: to, totalSlides: total });
    `, { documentName, fromSlideNumber, toSlideNumber })),
  );

  server.tool(
    "keynote_skip_slide",
    "Mark a slide as skipped (hidden) or unskipped",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      skipped: z.boolean().describe("True to skip/hide, false to unskip/show"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, skipped }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      doc.slides[params.slideNumber - 1].skipped = params.skipped;
      return JSON.stringify({ slideNumber: params.slideNumber, skipped: params.skipped });
    `, { documentName, slideNumber, skipped })),
  );

  server.tool(
    "keynote_set_slide_layout",
    "Change the master slide layout of an existing slide (use keynote_list_master_slides to see available layouts)",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      masterSlideName: z.string().describe("Master slide layout name (e.g. 'Title & Subtitle', 'Bullets', 'Blank')"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, masterSlideName }) => handleJXA(() => runJXA<string>(`
      ${KEYNOTE_AS_HELPERS}
      runAS('tell application "Keynote" to tell ' + docRef() + ' to set base slide of slide ' + params.slideNumber + ' to master slide ' + asQuote(params.masterSlideName));
      return JSON.stringify({ layoutSet: true, slideNumber: params.slideNumber, masterSlide: params.masterSlideName });
    `, { documentName, slideNumber, masterSlideName })),
  );

  server.tool(
    "keynote_stop_slideshow",
    "Stop a running slideshow",
    {
      documentName: z.string().describe("Name of the open presentation"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName }) => handleJXA(() => runJXA<string>(`
      ${KEYNOTE_AS_HELPERS}
      const app = Application("Keynote");
      app.documents.byName(params.documentName).name(); // throws if missing
      if (!app.playing()) return JSON.stringify({ stopped: true, wasPlaying: false });
      // JXA app.stop() can leave the show running
      function stoppedWithin(seconds) {
        for (let i = 0; i < seconds * 4; i++) {
          if (!app.playing()) return true;
          delay(0.25);
        }
        return false;
      }
      runAS('tell application "Keynote" to stop ' + docRef());
      if (!stoppedWithin(2)) {
        try { runAS('tell application "Keynote" to stop front document'); } catch (e) {}
      }
      if (!stoppedWithin(2)) throw new Error("The slideshow is still playing");
      return JSON.stringify({ stopped: true, wasPlaying: true });
    `, { documentName })),
  );

  server.tool(
    "keynote_add_slide",
    "Add a new slide to the presentation",
    {
      documentName: z.string().describe("Name of the open presentation"),
      masterSlideName: z.string().optional().describe("Master slide / layout name (e.g. 'Title & Subtitle', 'Blank')"),
      afterSlide: z.number().int().min(1).optional().describe("Insert after this slide number (1-based). Default: end."),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, masterSlideName, afterSlide }) => handleJXA(() => runJXA<string>(`
      ${KEYNOTE_AS_HELPERS}
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const before = doc.slides.length;
      const hasAfter = params.afterSlide !== null;
      if (hasAfter && params.afterSlide > before) {
        throw new Error("afterSlide " + params.afterSlide + " is out of range (presentation has " + before + " slides)");
      }

      if (params.masterSlideName || hasAfter) {
        // JXA masterSlides bridge is broken and slides.splice() fails (-10002)
        let asCode = 'tell application "Keynote" to tell ' + docRef() + ' to make new slide';
        if (hasAfter) asCode += ' at after slide ' + params.afterSlide;
        if (params.masterSlideName) asCode += ' with properties {base slide:master slide ' + asQuote(params.masterSlideName) + '}';
        runAS(asCode);
      } else {
        doc.slides.push(app.Slide({}));
      }

      const total = doc.slides.length;
      if (total !== before + 1) {
        throw new Error("Slide was not added (slide count " + before + " -> " + total + ")");
      }
      return JSON.stringify({ slideNumber: hasAfter ? params.afterSlide + 1 : total, added: true, totalSlides: total });
    `, { documentName, masterSlideName: masterSlideName ?? null, afterSlide: afterSlide ?? null })),
  );

  server.tool(
    "keynote_set_slide_title",
    "Set the title text of a slide (turns the title placeholder on if the layout hides it)",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      title: z.string().describe("Title text"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, title }) => handleJXA(() => runJXA<string>(`
      ${PLACEHOLDER_HELPER}
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
      const madeVisible = showPlaceholder(slide, params.slideNumber, "titleShowing", "title");
      const titleItem = slide.defaultTitleItem();
      titleItem.objectText = params.title;
      return JSON.stringify({ slideNumber: params.slideNumber, title: params.title, set: true, titleMadeVisible: madeVisible });
    `, { documentName, slideNumber, title })),
  );

  server.tool(
    "keynote_set_slide_body",
    "Set the body text of a slide (bullet points separated by newlines; turns the body placeholder on if the layout hides it)",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      body: z.string().describe("Body text (use newlines for bullet points)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, body }) => handleJXA(() => runJXA<string>(`
      ${PLACEHOLDER_HELPER}
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
      const madeVisible = showPlaceholder(slide, params.slideNumber, "bodyShowing", "body");
      const bodyItem = slide.defaultBodyItem();
      bodyItem.objectText = params.body;
      return JSON.stringify({ slideNumber: params.slideNumber, set: true, bodyMadeVisible: madeVisible });
    `, { documentName, slideNumber, body })),
  );

  server.tool(
    "keynote_add_image_to_slide",
    "Add an image to a slide",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      filePath: z.string().startsWith("/").describe("Absolute path to the image file"),
      x: z.number().optional().describe("X position in points"),
      y: z.number().optional().describe("Y position in points"),
      width: z.number().positive().optional().describe("Width in points"),
      height: z.number().positive().optional().describe("Height in points"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, filePath, x, y, width, height }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
      const props = { file: Path(params.filePath) };
      if (params.width !== null) props.width = params.width;
      if (params.height !== null) props.height = params.height;
      slide.images.push(app.Image(props));
      const image = slide.images[slide.images.length - 1];
      // [x, y] is silently ignored; position needs {x, y}
      if (params.x !== null || params.y !== null) {
        const cur = image.position();
        image.position = { x: params.x !== null ? params.x : cur.x, y: params.y !== null ? params.y : cur.y };
      }
      const p = image.position();
      return JSON.stringify({ added: true, slideNumber: params.slideNumber, x: p.x, y: p.y, width: image.width(), height: image.height() });
    `, { documentName, slideNumber, filePath, x: x ?? null, y: y ?? null, width: width ?? null, height: height ?? null })),
  );

  server.tool(
    "keynote_add_shape",
    "Add a shape with text to a slide",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      text: z.string().optional().describe("Text content for the shape"),
      x: z.number().optional().describe("X position in points"),
      y: z.number().optional().describe("Y position in points"),
      width: z.number().positive().optional().describe("Width in points (default: 200)"),
      height: z.number().positive().optional().describe("Height in points (default: 100)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, text, x, y, width, height }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
      const props = {};
      if (params.width !== null) props.width = params.width;
      if (params.height !== null) props.height = params.height;
      slide.shapes.push(app.Shape(props));
      const shape = slide.shapes[slide.shapes.length - 1];
      // [x, y] is silently ignored; position needs {x, y}
      if (params.x !== null || params.y !== null) {
        const cur = shape.position();
        shape.position = { x: params.x !== null ? params.x : cur.x, y: params.y !== null ? params.y : cur.y };
      }
      if (params.text) {
        shape.objectText = params.text;
      }
      const p = shape.position();
      return JSON.stringify({ added: true, slideNumber: params.slideNumber, x: p.x, y: p.y, width: shape.width(), height: shape.height() });
    `, { documentName, slideNumber, text: text ?? null, x: x ?? null, y: y ?? null, width: width ?? 200, height: height ?? 100 })),
  );

  server.tool(
    "keynote_add_text_box",
    "Add a free text box to a slide",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      text: z.string().describe("Text content (newlines start new paragraphs)"),
      x: z.number().optional().describe("X position in points"),
      y: z.number().optional().describe("Y position in points"),
      width: z.number().positive().optional().describe("Width in points"),
      height: z.number().positive().optional().describe("Height in points"),
      fontName: z.string().optional().describe("PostScript font name (e.g. 'HelveticaNeue-Bold')"),
      fontSize: z.number().positive().optional().describe("Font size in points"),
      textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Text color as hex, e.g. '#FFFFFF'"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, text, x, y, width, height, fontName, fontSize, textColor }) => handleJXA(() => runJXA<string>(`
      ${SLIDE_LOOKUP}
      ${ITEM_KEY_HELPER}
      function keysOf(items) {
        return items.map(function (t) { let text = ""; try { text = t.objectText(); } catch (e) {} return itemKey(t, text); });
      }
      const beforeKeys = keysOf(slide.textItems());
      slide.textItems.push(app.TextItem({ objectText: params.text }));
      const items = slide.textItems();
      if (items.length !== beforeKeys.length + 1) throw new Error("Text box was not added");
      // The new box isn't necessarily last (Keynote can insert it before hidden placeholders):
      // it's the first item whose text + geometry occurs more often than before
      const remaining = {};
      beforeKeys.forEach(function (k) { remaining[k] = (remaining[k] || 0) + 1; });
      const index = keysOf(items).findIndex(function (k) {
        if (remaining[k]) { remaining[k]--; return false; }
        return true;
      });
      const box = items[index];
      if (params.width !== null) box.width = params.width;
      if (params.height !== null) box.height = params.height;
      if (params.fontName !== null) box.objectText.font = params.fontName;
      if (params.fontSize !== null) box.objectText.size = params.fontSize;
      if (params.textColor !== null) {
        const hex = params.textColor;
        box.objectText.color = [parseInt(hex.slice(1, 3), 16) * 257, parseInt(hex.slice(3, 5), 16) * 257, parseInt(hex.slice(5, 7), 16) * 257];
      }
      // Last: a font change resizes the box and Keynote keeps it centered
      if (params.x !== null || params.y !== null) {
        const cur = box.position();
        box.position = { x: params.x !== null ? params.x : cur.x, y: params.y !== null ? params.y : cur.y };
      }
      const p = box.position();
      return JSON.stringify({ added: true, slideNumber: params.slideNumber, textItemIndex: index, x: p.x, y: p.y, width: box.width(), height: box.height() });
    `, {
      documentName, slideNumber, text, x: x ?? null, y: y ?? null, width: width ?? null, height: height ?? null,
      fontName: fontName ?? null, fontSize: fontSize ?? null, textColor: textColor ?? null,
    })),
  );

  server.tool(
    "keynote_set_item_text",
    "Replace the text of a text box or shape on a slide, including layout placeholders such as a quote or attribution (use keynote_list_slide_items to find the index)",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      itemType: z.enum(["shape", "textItem"]).describe("Item type"),
      itemIndex: z.number().int().min(0).describe("Item index (0-based)"),
      text: z.string().describe("New text (newlines start new paragraphs)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, itemType, itemIndex, text }) => handleJXA(() => runJXA<string>(`
      ${SLIDE_LOOKUP}
      ${ITEM_LOOKUP}
      assertUnlocked();
      item.objectText = params.text;
      return JSON.stringify({ set: true, slideNumber: params.slideNumber, itemType: params.itemType, itemIndex: params.itemIndex, text: item.objectText() });
    `, { documentName, slideNumber, itemType, itemIndex, text })),
  );

  server.tool(
    "keynote_delete_item",
    "Delete an item (shape, text box, image, table, line, or chart) from a slide (use keynote_list_slide_items to find the index). Locked items must be unlocked first (keynote_set_item_locked).",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      itemType: z.enum(["shape", "textItem", "image", "table", "line", "chart"]).describe("Item type"),
      itemIndex: z.number().int().min(0).describe("Item index (0-based)"),
    },
    ANNOTATIONS.destructive,
    async ({ documentName, slideNumber, itemType, itemIndex }) => handleJXA(() => runJXA<string>(`
      ${SLIDE_LOOKUP}
      ${ITEM_LOOKUP}
      assertUnlocked();
      const before = itemList.length;
      app.delete(item);
      const after = collections[params.itemType]().length;
      if (after >= before) throw new Error("Item was not deleted");
      return JSON.stringify({ deleted: true, slideNumber: params.slideNumber, itemType: params.itemType, itemIndex: params.itemIndex });
    `, { documentName, slideNumber, itemType, itemIndex })),
  );

  server.tool(
    "keynote_set_item_locked",
    "Lock or unlock an item on a slide. Locked items can't be moved, edited, or deleted until unlocked.",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      itemType: z.enum(["shape", "textItem", "image", "table", "line", "chart"]).describe("Item type"),
      itemIndex: z.number().int().min(0).describe("Item index (0-based)"),
      locked: z.boolean().describe("True to lock, false to unlock"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, itemType, itemIndex, locked }) => handleJXA(() => runJXA<string>(`
      ${SLIDE_LOOKUP}
      ${ITEM_LOOKUP}
      item.locked = params.locked;
      return JSON.stringify({ slideNumber: params.slideNumber, itemType: params.itemType, itemIndex: params.itemIndex, locked: item.locked() });
    `, { documentName, slideNumber, itemType, itemIndex, locked })),
  );

  server.tool(
    "keynote_add_chart",
    "Add a native Keynote chart to a slide. data[r][c] is the value for row r and column c. groupBy 'row' (default): each row is one series (one line/bar color) and the column names label the x-axis; 'column': each column is a series and the row names go on the x-axis. Keynote's scripting can't add a legend or chart title — add a text box for that.",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      chartType: z.enum([
        "vertical_bar_2d", "stacked_vertical_bar_2d", "horizontal_bar_2d", "stacked_horizontal_bar_2d",
        "line_2d", "area_2d", "stacked_area_2d", "pie_2d", "scatterplot_2d",
        "vertical_bar_3d", "stacked_vertical_bar_3d", "horizontal_bar_3d", "stacked_horizontal_bar_3d",
        "line_3d", "area_3d", "stacked_area_3d", "pie_3d",
      ]).optional().describe("Chart type (default: vertical_bar_2d)"),
      rowNames: z.array(z.string()).min(1).describe("Row labels, one per data row"),
      columnNames: z.array(z.string()).min(1).describe("Column labels, one per data column"),
      data: z.array(z.array(z.number())).min(1).describe("Values as rows x columns"),
      groupBy: z.enum(["row", "column"]).optional().describe("Whether rows or columns form the data series (default: row)"),
      x: z.number().optional().describe("X position in points"),
      y: z.number().optional().describe("Y position in points"),
      width: z.number().positive().optional().describe("Width in points"),
      height: z.number().positive().optional().describe("Height in points"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, chartType, rowNames, columnNames, data, groupBy, x, y, width, height }) => {
      if (data.length !== rowNames.length || data.some(row => row.length !== columnNames.length)) {
        return toolResult(`data needs one row per row name (${rowNames.length}) and one value per column name (${columnNames.length})`, true);
      }
      return handleJXA(() => runJXA<string>(`
      ${KEYNOTE_AS_HELPERS}
      ${SLIDE_LOOKUP}
      const before = slide.charts.length;
      const list = function (values) { return "{" + values.map(function (v) { return typeof v === "number" ? String(v) : asQuote(v); }).join(", ") + "}"; };
      runAS('tell application "Keynote" to add chart slide ' + params.slideNumber + ' of ' + docRef() +
        ' row names ' + list(params.rowNames) + ' column names ' + list(params.columnNames) +
        ' data {' + params.data.map(list).join(", ") + '}' +
        ' type ' + params.chartType + ' group by ' + (params.groupBy === "column" ? "chart column" : "chart row"));
      if (slide.charts.length !== before + 1) throw new Error("Chart was not added");
      const chart = slide.charts[slide.charts.length - 1];
      if (params.width !== null) chart.width = params.width;
      if (params.height !== null) chart.height = params.height;
      if (params.x !== null || params.y !== null) {
        const cur = chart.position();
        chart.position = { x: params.x !== null ? params.x : cur.x, y: params.y !== null ? params.y : cur.y };
      }
      const p = chart.position();
      return JSON.stringify({ added: true, slideNumber: params.slideNumber, chartIndex: slide.charts.length - 1, x: p.x, y: p.y, width: chart.width(), height: chart.height() });
    `, {
        documentName, slideNumber, chartType: chartType ?? "vertical_bar_2d", rowNames, columnNames, data, groupBy: groupBy ?? "row",
        x: x ?? null, y: y ?? null, width: width ?? null, height: height ?? null,
      }));
    },
  );

  server.tool(
    "keynote_add_image_slides",
    "Add one new slide per image file at the end of the presentation (removes the duplicate slides Keynote 15.4 creates)",
    {
      documentName: z.string().describe("Name of the open presentation"),
      filePaths: z.array(z.string().startsWith("/")).min(1).describe("Absolute paths to the image files, in slide order"),
      setTitles: z.boolean().optional().describe("Use each file name as the slide title (default: false)"),
      masterSlideName: z.string().optional().describe("Slide layout for the new slides (default: the layout Keynote last used for image slides)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, filePaths, setTitles, masterSlideName }) => {
      const missing = filePaths.filter(p => !existsSync(p));
      if (missing.length) return toolResult(`File not found: ${missing.join(", ")}`, true);
      return handleJXA(() => runJXA<string>(`
      ${KEYNOTE_AS_HELPERS}
      ${PLACEHOLDER_HELPER}
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const before = doc.slides.length;
      const originalCurrent = doc.currentSlide().slideNumber();
      function setCurrent(n) { runAS('tell application "Keynote" to set current slide of ' + docRef() + ' to slide ' + n + ' of ' + docRef()); }

      function baseName(p) { return String(p).split("/").pop(); }
      function stem(n) { return n.replace(/\\.[^.]*$/, ""); }
      // Keynote may rename an imported copy (e.g. "photo 2.png"): match on the file stem
      function fileFor(imageName) {
        return Object.keys(wanted).find(function (k) {
          return stem(imageName) === stem(k) || stem(imageName).indexOf(stem(k)) === 0;
        });
      }
      const wanted = {};
      params.filePaths.forEach(function (p) { const n = baseName(p); wanted[n] = (wanted[n] || 0) + 1; });

      let total = before, extra = [];
      const hiddenTitles = [], resizedForTitle = [], coveredTitles = [];

      // Photo layouts can stretch the image over the title, and scripting can't reorder layers:
      // shrink the image into the larger free band above or below the title, keeping its proportions
      function fitBesideTitle(image, title) {
        const ip = image.position(), iw = image.width(), ih = image.height();
        let tp = title.position(), th = title.height();
        const tw = title.width();
        // Title boxes are often far taller than one line of text: shrink the box to the text first,
        // keeping the edge away from the image's center so the freed space joins the image's side
        try {
          const size = title.objectText.paragraphs[0].size();
          const perLine = Math.max(1, Math.floor((tw - 20) / (size * 0.55)));
          const lines = title.objectText().split(/\\r|\\n/).reduce(function (n, para) {
            return n + Math.max(1, Math.ceil(para.length / perLine));
          }, 0);
          const needed = Math.ceil(lines * size * 1.25 + 20);
          if (needed < th) {
            const y = tp.y + th / 2 >= ip.y + ih / 2 ? tp.y + th - needed : tp.y;
            title.height = needed;
            title.position = { x: tp.x, y: y };
            tp = { x: tp.x, y: y };
            th = needed;
          }
        } catch (e) {}
        const overlaps = tp.y < ip.y + ih && tp.y + th > ip.y && tp.x < ip.x + iw && tp.x + tw > ip.x;
        if (!overlaps) return "clear";
        const gap = 10;
        const above = { top: ip.y, height: tp.y - gap - ip.y };
        const below = { top: tp.y + th + gap, height: ip.y + ih - (tp.y + th + gap) };
        const band = above.height >= below.height ? above : below;
        if (band.height < ih / 4) return "covered";
        const scale = band.height / ih;
        const w = Math.round(iw * scale), h = Math.round(band.height);
        image.width = w;
        image.height = h;
        image.position = { x: Math.round(ip.x + (iw - w) / 2), y: Math.round(band.top) };
        return "resized";
      }
      try {
        // New slides go after the current slide: make that the last one
        setCurrent(before);
        let src = 'set fileList to {' + params.filePaths.map(function (p) { return "POSIX file " + asQuote(p); }).join(", ") + '}\\n' +
          'tell application "Keynote" to make image slides ' + docRef() + ' files fileList set titles ' + (params.setTitles ? "true" : "false");
        if (params.masterSlideName) src += ' slide layout (slide layout ' + asQuote(params.masterSlideName) + ' of ' + docRef() + ')';
        runAS(src);

        // Keynote 15.4 creates every image slide twice. Keep as many slides per file as requested.
        // Photo layouts also bring a sample image of their own, which is removed.
        const after = doc.slides.length;
        const kept = [];
        for (let n = before + 1; n <= after; n++) {
          const slide = doc.slides[n - 1];
          let file = null;
          const samples = [];
          slide.images().forEach(function (img) {
            let name = "";
            try { name = baseName(img.fileName()); } catch (e) {}
            const match = name && fileFor(name);
            if (match && !file) file = match; else samples.push(img);
          });
          if (!file) {
            throw new Error("Slide " + n + " is not one of the new image slides; Keynote inserted them elsewhere (slides " + before + " -> " + after + ")");
          }
          if (wanted[file] > 0) {
            wanted[file]--;
            kept.push({ n: n, file: file, samples: samples });
          } else {
            extra.push(n);
          }
        }
        kept.forEach(function (k) {
          k.samples.forEach(function (img) { try { app.delete(img); } catch (e) {} });
        });
        for (let i = extra.length - 1; i >= 0; i--) app.delete(doc.slides[extra[i] - 1]);
        total = doc.slides.length;
        if (total !== before + params.filePaths.length) {
          throw new Error("Expected " + params.filePaths.length + " new slides, got " + (total - before));
        }

        // Some layouts hide the title or leave it empty, so file-name titles wouldn't show
        if (params.setTitles) {
          kept.forEach(function (k, i) {
            const n = before + i + 1;
            const slide = doc.slides[n - 1];
            try {
              showPlaceholder(slide, n, "titleShowing", "title");
              const title = slide.defaultTitleItem();
              if (!title.objectText()) title.objectText = stem(k.file);
              const image = slide.images().find(function (img) {
                try { return fileFor(baseName(img.fileName())) === k.file; } catch (e) { return false; }
              });
              if (image) {
                const fit = fitBesideTitle(image, title);
                if (fit === "resized") resizedForTitle.push(n);
                else if (fit === "covered") coveredTitles.push(n);
              }
            } catch (e) { hiddenTitles.push(n); }
          });
        }
      } finally {
        try { setCurrent(originalCurrent); } catch (e) {}
      }

      const result = { added: params.filePaths.length, firstSlideNumber: before + 1, totalSlides: total, duplicatesRemoved: extra.length };
      if (hiddenTitles.length) result.titlesHiddenOnSlides = hiddenTitles;
      if (resizedForTitle.length) result.imagesShrunkToShowTitleOnSlides = resizedForTitle;
      if (coveredTitles.length) result.titlesCoveredByImageOnSlides = coveredTitles;
      return JSON.stringify(result);
    `, { documentName, filePaths, setTitles: setTitles ?? false, masterSlideName: masterSlideName ?? null }));
    },
  );

  server.tool(
    "keynote_add_line",
    "Add a line to a slide between two points",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      startX: z.number().describe("Start X position in points"),
      startY: z.number().describe("Start Y position in points"),
      endX: z.number().describe("End X position in points"),
      endY: z.number().describe("End Y position in points"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, startX, startY, endX, endY }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
      const line = app.Line({
        startPoint: { x: params.startX, y: params.startY },
        endPoint: { x: params.endX, y: params.endY },
      });
      slide.lines.push(line);
      const added = slide.lines[slide.lines.length - 1];
      return JSON.stringify({
        added: true,
        startPoint: added.startPoint(),
        endPoint: added.endPoint(),
      });
    `, { documentName, slideNumber, startX, startY, endX, endY })),
  );

  server.tool(
    "keynote_set_presenter_notes",
    "Set presenter notes for a slide",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      notes: z.string().describe("Presenter notes text"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, notes }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
      slide.presenterNotes = params.notes;
      return JSON.stringify({ slideNumber: params.slideNumber, set: true });
    `, { documentName, slideNumber, notes })),
  );

  server.tool(
    "keynote_set_transition",
    "Set a transition effect on a slide",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      effect: z.string().describe("Transition effect name as in Keynote's scripting dictionary (e.g. 'dissolve', 'push', 'wipe', 'magic move', or 'none'). An invalid name returns the list of valid ones."),
      duration: z.number().positive().optional().describe("Duration in seconds (default: 1.0)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, effect, duration }) => handleJXA(async () => {
      try {
        return await runJXA<string>(`
      ${SLIDE_LOOKUP}
      const transition = {
        transitionEffect: params.effect,
        transitionDuration: params.duration || 1.0,
      };
      try {
        slide.transitionProperties = transition;
      } catch (e) {
        throw new Error("Could not set transition effect '" + params.effect + "': " + e.message);
      }
      return JSON.stringify({ slideNumber: params.slideNumber, effect: params.effect, set: true });
    `, { documentName, slideNumber, effect: normalizeTransitionEffect(effect), duration: duration ?? null });
      } catch (err) {
        if (err instanceof OsascriptError && err.message.includes("Could not set transition effect")) {
          const names = await keynoteTransitionEffects();
          if (names.length) throw new Error(`${err.stderr.trim()}\nValid effects: ${names.join(", ")} (or 'none')`);
        }
        throw err;
      }
    }),
  );

  server.tool(
    "keynote_list_transition_effects",
    "List the transition effect names keynote_set_transition accepts (read from Keynote's scripting dictionary)",
    {},
    ANNOTATIONS.readOnly,
    async () => {
      const names = await keynoteTransitionEffects();
      if (!names.length) return toolResult("Could not read the transition effects from Keynote's scripting dictionary", true);
      return toolResult(JSON.stringify({ effects: names, note: "'none' is accepted as an alias for 'no transition effect'" }, null, 2));
    },
  );

  server.tool(
    "keynote_start_slideshow",
    "Start playing the presentation slideshow",
    {
      documentName: z.string().describe("Name of the open presentation"),
      fromSlide: z.number().int().min(1).optional().describe("Start from this slide number (1-based, default: 1)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, fromSlide }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      if (params.fromSlide) {
        doc.slides[params.fromSlide - 1].skipped = false;
        app.startFrom(doc.slides[params.fromSlide - 1]);
      } else {
        app.start(doc);
      }
      return JSON.stringify({ playing: true });
    `, { documentName, fromSlide: fromSlide ?? null })),
  );

  // ── Layout Tools ──

  server.tool(
    "keynote_position_item",
    "Move and/or resize an item (shape, image, textItem, table, chart, or line) on a slide by type and index",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      itemType: z.enum(["shape", "image", "textItem", "table", "chart", "line"]).describe("Type of item"),
      itemIndex: z.number().int().min(0).describe("Item index (0-based)"),
      x: z.number().optional().describe("New X position in points"),
      y: z.number().optional().describe("New Y position in points"),
      width: z.number().positive().optional().describe("New width in points"),
      height: z.number().positive().optional().describe("New height in points"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, itemType, itemIndex, x, y, width, height }) => handleJXA(() => runJXA<string>(`
      ${SLIDE_LOOKUP}
      ${ITEM_LOOKUP}
      assertUnlocked();

      // Size first, so the requested position holds even if Keynote re-centers on resize
      if (params.width !== null) item.width = params.width;
      if (params.height !== null) item.height = params.height;

      const pos = item.position();
      const newX = params.x !== null ? params.x : pos.x;
      const newY = params.y !== null ? params.y : pos.y;
      item.position = {x: newX, y: newY};

      return JSON.stringify({
        positioned: true,
        x: item.position().x,
        y: item.position().y,
        width: item.width(),
        height: item.height(),
      });
    `, { documentName, slideNumber, itemType, itemIndex, x: x ?? null, y: y ?? null, width: width ?? null, height: height ?? null })),
  );

  server.tool(
    "keynote_align_items",
    "Align two or more items on a slide along an edge or center",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      items: z.array(z.object({
        itemType: z.enum(["shape", "image", "textItem"]).describe("Type of item"),
        itemIndex: z.number().int().min(0).describe("Item index (0-based)"),
      })).min(2).describe("Items to align (at least 2)"),
      alignment: z.enum(["left", "center", "right", "top", "middle", "bottom"]).describe("Alignment edge or center"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, items, alignment }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];

      const collections = { shape: slide.shapes, image: slide.images, textItem: slide.textItems };
      const resolved = params.items.map(function(spec) {
        var col = collections[spec.itemType]();
        if (spec.itemIndex >= col.length) throw new Error(spec.itemType + " index " + spec.itemIndex + " out of range");
        return col[spec.itemIndex];
      });

      // Compute target coordinate
      var target;
      var align = params.alignment;
      if (align === "left") {
        target = Math.min.apply(null, resolved.map(function(it) { return it.position().x; }));
        resolved.forEach(function(it) { var p = it.position(); it.position = {x: target, y: p.y}; });
      } else if (align === "right") {
        target = Math.max.apply(null, resolved.map(function(it) { return it.position().x + it.width(); }));
        resolved.forEach(function(it) { var p = it.position(); it.position = {x: target - it.width(), y: p.y}; });
      } else if (align === "center") {
        var centers = resolved.map(function(it) { return it.position().x + it.width() / 2; });
        target = centers.reduce(function(a, b) { return a + b; }, 0) / centers.length;
        resolved.forEach(function(it) { var p = it.position(); it.position = {x: target - it.width() / 2, y: p.y}; });
      } else if (align === "top") {
        target = Math.min.apply(null, resolved.map(function(it) { return it.position().y; }));
        resolved.forEach(function(it) { var p = it.position(); it.position = {x: p.x, y: target}; });
      } else if (align === "bottom") {
        target = Math.max.apply(null, resolved.map(function(it) { return it.position().y + it.height(); }));
        resolved.forEach(function(it) { var p = it.position(); it.position = {x: p.x, y: target - it.height()}; });
      } else if (align === "middle") {
        var middles = resolved.map(function(it) { return it.position().y + it.height() / 2; });
        target = middles.reduce(function(a, b) { return a + b; }, 0) / middles.length;
        resolved.forEach(function(it) { var p = it.position(); it.position = {x: p.x, y: target - it.height() / 2}; });
      }

      return JSON.stringify({ aligned: true, alignment: params.alignment, itemCount: resolved.length });
    `, { documentName, slideNumber, items, alignment })),
  );

  server.tool(
    "keynote_distribute_items",
    "Evenly space three or more items on a slide horizontally or vertically",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      items: z.array(z.object({
        itemType: z.enum(["shape", "image", "textItem"]).describe("Type of item"),
        itemIndex: z.number().int().min(0).describe("Item index (0-based)"),
      })).min(3).describe("Items to distribute (at least 3)"),
      direction: z.enum(["horizontal", "vertical"]).describe("Distribution direction"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, items, direction }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];

      const collections = { shape: slide.shapes, image: slide.images, textItem: slide.textItems };
      const resolved = params.items.map(function(spec) {
        var col = collections[spec.itemType]();
        if (spec.itemIndex >= col.length) throw new Error(spec.itemType + " index " + spec.itemIndex + " out of range");
        return col[spec.itemIndex];
      });

      if (params.direction === "horizontal") {
        // Sort by x position
        resolved.sort(function(a, b) { return a.position().x - b.position().x; });
        var firstX = resolved[0].position().x;
        var lastX = resolved[resolved.length - 1].position().x;
        var totalWidth = resolved.reduce(function(sum, it) { return sum + it.width(); }, 0);
        var totalSpace = (lastX + resolved[resolved.length - 1].width()) - firstX;
        var gap = (totalSpace - totalWidth) / (resolved.length - 1);
        var currentX = firstX;
        for (var i = 0; i < resolved.length; i++) {
          var p = resolved[i].position();
          resolved[i].position = {x: currentX, y: p.y};
          currentX += resolved[i].width() + gap;
        }
      } else {
        // Sort by y position
        resolved.sort(function(a, b) { return a.position().y - b.position().y; });
        var firstY = resolved[0].position().y;
        var lastY = resolved[resolved.length - 1].position().y;
        var totalHeight = resolved.reduce(function(sum, it) { return sum + it.height(); }, 0);
        var totalSpaceV = (lastY + resolved[resolved.length - 1].height()) - firstY;
        var gapV = (totalSpaceV - totalHeight) / (resolved.length - 1);
        var currentY = firstY;
        for (var j = 0; j < resolved.length; j++) {
          var pv = resolved[j].position();
          resolved[j].position = {x: pv.x, y: currentY};
          currentY += resolved[j].height() + gapV;
        }
      }

      return JSON.stringify({ distributed: true, direction: params.direction, itemCount: resolved.length });
    `, { documentName, slideNumber, items, direction })),
  );

  // ── Shape Formatting Tools ──

  server.tool(
    "keynote_get_shape_info",
    "Read shape or text box properties: position, size, text, opacity, rotation, and text formatting",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      shapeIndex: z.number().int().min(0).describe("Item index (0-based)"),
      itemType: TEXT_ITEM_TYPE,
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, slideNumber, shapeIndex, itemType }) => handleJXA(() => runJXA<string>(`
      ${SLIDE_LOOKUP}
      ${ITEM_LOOKUP}
      const shape = item;

      const info = {};
      try { info.position = shape.position(); } catch(e) {}
      try { info.width = shape.width(); } catch(e) {}
      try { info.height = shape.height(); } catch(e) {}
      try { info.opacity = shape.opacity(); } catch(e) {}
      try { info.rotation = shape.rotation(); } catch(e) {}
      try { info.text = shape.objectText(); } catch(e) {}
      try {
        const p = shape.objectText.paragraphs[0];
        info.textFormat = {};
        try { info.textFormat.font = p.font(); } catch(e) {}
        try { info.textFormat.size = p.size(); } catch(e) {}
        try {
          const c = p.color();
          if (Array.isArray(c) && c.length >= 3) {
            const scale = Math.max(c[0], c[1], c[2]) > 1 ? 65535 : 1;
            const toHex = function(v) { return Math.round(v / scale * 255).toString(16).padStart(2, "0"); };
            info.textFormat.color = "#" + toHex(c[0]) + toHex(c[1]) + toHex(c[2]);
          }
        } catch(e) {}
      } catch(e) {}

      return JSON.stringify(info);
    `, { documentName, slideNumber, itemIndex: shapeIndex, itemType: itemType ?? "shape" })),
  );

  server.tool(
    "keynote_format_shape",
    "Set shape or text box properties: opacity, rotation, and text formatting (font, size, color; use a bold PostScript font name for bold). Note: fill and border colors and text alignment are not accessible via scripting.",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      shapeIndex: z.number().int().min(0).describe("Item index (0-based)"),
      itemType: TEXT_ITEM_TYPE,
      format: z.object({
        opacity: z.number().min(0).max(100).optional().describe("Opacity (0-100)"),
        rotation: z.number().optional().describe("Rotation in degrees"),
        fontName: z.string().optional().describe("PostScript font name for shape text"),
        fontSize: z.number().positive().optional().describe("Font size in points"),
        textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Text color as hex, e.g. '#FF0000'"),
      }).strict().describe("Format options"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, shapeIndex, itemType, format }) => handleJXA(() => runJXA<string>(`
      ${SLIDE_LOOKUP}
      ${ITEM_LOOKUP}
      assertUnlocked();
      const shape = item;
      const fmt = params.format;

      // No rollback: report what already took effect
      const applied = [];
      function apply(name, fn) {
        try { fn(); } catch (e) {
          throw new Error("Failed to set " + name + ": " + e.message + (applied.length ? " (already applied: " + applied.join(", ") + ")" : ""));
        }
        applied.push(name);
      }

      if (fmt.opacity !== undefined) apply("opacity", function () { shape.opacity = fmt.opacity; });
      if (fmt.rotation !== undefined) apply("rotation", function () { shape.rotation = fmt.rotation; });

      // Text formatting applies to all paragraphs
      const paras = shape.objectText.paragraphs;
      const paraCount = paras.length;
      if (fmt.fontName !== undefined) apply("fontName", function () { for (let i = 0; i < paraCount; i++) paras[i].font = fmt.fontName; });
      if (fmt.fontSize !== undefined) apply("fontSize", function () { for (let i = 0; i < paraCount; i++) paras[i].size = fmt.fontSize; });
      if (fmt.textColor !== undefined) apply("textColor", function () {
        for (let i = 0; i < paraCount; i++) {
          const hex = fmt.textColor;
          // Keynote expects 0-65535, not 0-1
          const r = parseInt(hex.slice(1, 3), 16) * 257;
          const g = parseInt(hex.slice(3, 5), 16) * 257;
          const b = parseInt(hex.slice(5, 7), 16) * 257;
          paras[i].color = [r, g, b];
        }
      });

      return JSON.stringify({ formatted: true, shapeIndex: params.itemIndex, applied: applied });
    `, { documentName, slideNumber, itemIndex: shapeIndex, itemType: itemType ?? "shape", format })),
  );

  // ── Slide Table Tools ──

  server.tool(
    "keynote_add_table_to_slide",
    "Add a table to a slide with optional data",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      rows: z.number().int().min(1).optional().describe("Number of rows (default: 3)"),
      columns: z.number().int().min(1).optional().describe("Number of columns (default: 3)"),
      data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).optional().describe("2D array of data — first row can be headers"),
      x: z.number().optional().describe("X position in points"),
      y: z.number().optional().describe("Y position in points"),
      width: z.number().positive().optional().describe("Width in points"),
      height: z.number().positive().optional().describe("Height in points"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, rows, columns, data, x, y, width, height }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
      const rowCount = params.data ? params.data.length : (params.rows || 3);
      const colCount = params.data && params.data[0] ? params.data[0].length : (params.columns || 3);
      const tbl = app.Table({ rowCount: Math.max(rowCount, 1), columnCount: Math.max(colCount, 1) });
      slide.tables.push(tbl);
      const addedTable = slide.tables[slide.tables.length - 1];
      if (params.x != null || params.y != null) {
        const pos = addedTable.position();
        addedTable.position = { x: params.x != null ? params.x : pos.x, y: params.y != null ? params.y : pos.y };
      }
      if (params.width != null) addedTable.width = params.width;
      if (params.height != null) addedTable.height = params.height;
      if (params.data) {
        for (var r = 0; r < params.data.length; r++) {
          for (var c = 0; c < params.data[r].length; c++) {
            if (params.data[r][c] !== null && params.data[r][c] !== undefined) {
              var colLetter = String.fromCharCode(65 + c);
              addedTable.cells[colLetter + (r + 1)].value = params.data[r][c];
            }
          }
        }
      }
      return JSON.stringify({
        added: true,
        tableName: addedTable.name(),
        rows: addedTable.rowCount(),
        columns: addedTable.columnCount(),
      });
    `, { documentName, slideNumber, rows: rows ?? null, columns: columns ?? null, data: data ?? null, x: x ?? null, y: y ?? null, width: width ?? null, height: height ?? null })),
  );

  server.tool(
    "keynote_read_slide_table",
    "Read all data from a table on a slide as a 2D array",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      tableIndex: z.number().int().min(0).optional().describe("Table index on the slide (0-based, default: 0)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, slideNumber, tableIndex }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
      const tables = slide.tables();
      const idx = params.tableIndex || 0;
      if (idx >= tables.length) {
        throw new Error("Table index " + idx + " out of range (slide has " + tables.length + " tables)");
      }
      const tbl = tables[idx];
      const rows = tbl.rowCount();
      const cols = tbl.columnCount();
      const data = [];
      for (var r = 0; r < rows; r++) {
        var row = [];
        for (var c = 0; c < cols; c++) {
          var colLetter = String.fromCharCode(65 + c);
          try { row.push(tbl.cells[colLetter + (r + 1)].value()); } catch(e) { row.push(null); }
        }
        data.push(row);
      }
      return JSON.stringify({
        tableName: tbl.name(),
        rows: rows,
        columns: cols,
        data: data,
      });
    `, { documentName, slideNumber, tableIndex: tableIndex ?? 0 })),
  );

  server.tool(
    "keynote_write_slide_table",
    "Write data to cells of an existing table on a slide",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      tableIndex: z.number().int().min(0).optional().describe("Table index on the slide (0-based, default: 0)"),
      data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).describe("2D array of data to write (starting from A1). Null cells are skipped."),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, tableIndex, data }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
      const tables = slide.tables();
      const idx = params.tableIndex || 0;
      if (idx >= tables.length) {
        throw new Error("Table index " + idx + " out of range (slide has " + tables.length + " tables)");
      }
      const tbl = tables[idx];
      var written = 0;
      for (var r = 0; r < params.data.length; r++) {
        for (var c = 0; c < params.data[r].length; c++) {
          if (params.data[r][c] !== null && params.data[r][c] !== undefined) {
            var colLetter = String.fromCharCode(65 + c);
            tbl.cells[colLetter + (r + 1)].value = params.data[r][c];
            written++;
          }
        }
      }
      return JSON.stringify({
        written: true,
        cellsWritten: written,
        tableName: tbl.name(),
      });
    `, { documentName, slideNumber, tableIndex: tableIndex ?? 0, data })),
  );

  server.tool(
    "keynote_format_slide_table",
    "Format cells in a table on a slide: text and background color, font (use a bold PostScript font name for bold), size, alignment, number format, and the width/height of the columns/rows in the range. The currency symbol can't be chosen via scripting.",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      tableIndex: z.number().int().min(0).optional().describe("Table index on the slide (0-based, default: 0)"),
      cellRange: z.string().regex(/^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$/).describe("Cell or range reference, e.g. 'A1' or 'A1:C3'"),
      format: z.object({
        textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Text color as hex, e.g. '#FF0000'"),
        backgroundColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Cell background color as hex"),
        fontName: z.string().optional().describe("PostScript font name, e.g. 'HelveticaNeue-Bold'"),
        fontSize: z.number().positive().optional().describe("Font size in points"),
        alignment: z.enum(["left", "center", "right", "justify"]).optional().describe("Horizontal text alignment"),
        verticalAlignment: z.enum(["top", "center", "bottom"]).optional().describe("Vertical text alignment"),
        numberFormat: z.enum(["automatic", "number", "currency", "percent", "fraction", "scientific", "text"]).optional().describe("Number format for the cells"),
        columnWidth: z.number().positive().optional().describe("Width in points for every column in the range"),
        rowHeight: z.number().positive().optional().describe("Height in points for every row in the range"),
      }).strict().describe("Formatting options to apply"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, tableIndex, cellRange, format }) => handleJXA(() => runJXA<string>(`
      ${SLIDE_LOOKUP}
      const tables = slide.tables();
      const idx = params.tableIndex || 0;
      if (idx >= tables.length) {
        throw new Error("Table index " + idx + " out of range (slide has " + tables.length + " tables)");
      }
      const tbl = tables[idx];
      const range = params.cellRange;
      // Parse range into list of cell refs
      var cells = [];
      if (range.indexOf(":") >= 0) {
        var parts = range.split(":");
        var startCol = parts[0].replace(/[0-9]/g, "");
        var startRow = parseInt(parts[0].replace(/[A-Z]/g, ""), 10);
        var endCol = parts[1].replace(/[0-9]/g, "");
        var endRow = parseInt(parts[1].replace(/[A-Z]/g, ""), 10);
        for (var r = startRow; r <= endRow; r++) {
          for (var c = startCol.charCodeAt(0); c <= endCol.charCodeAt(0); c++) {
            cells.push(String.fromCharCode(c) + r);
          }
        }
      } else {
        cells.push(range);
      }
      const fmt = params.format;
      function rgb(hex) {
        return [parseInt(hex.slice(1,3), 16) * 257, parseInt(hex.slice(3,5), 16) * 257, parseInt(hex.slice(5,7), 16) * 257];
      }
      const cellProps = {
        textColor: fmt.textColor && rgb(fmt.textColor),
        backgroundColor: fmt.backgroundColor && rgb(fmt.backgroundColor),
        fontName: fmt.fontName,
        fontSize: fmt.fontSize,
        alignment: fmt.alignment,
        verticalAlignment: fmt.verticalAlignment,
        format: fmt.numberFormat,
      };
      // No rollback: report what already took effect
      const applied = [];
      function apply(name, fn) {
        try { fn(); } catch (e) {
          throw new Error("Failed to set " + name + ": " + e.message + (applied.length ? " (already applied: " + applied.join(", ") + ")" : ""));
        }
        applied.push(name);
      }
      Object.keys(cellProps).forEach(function (prop) {
        if (cellProps[prop] === undefined) return;
        apply(prop === "format" ? "numberFormat" : prop, function () {
          for (var i = 0; i < cells.length; i++) tbl.cells[cells[i]][prop] = cellProps[prop];
        });
      });
      var refs = range.split(":");
      var firstCol = refs[0].replace(/[0-9]/g, "").charCodeAt(0) - 65;
      var lastCol = refs[refs.length - 1].replace(/[0-9]/g, "").charCodeAt(0) - 65;
      var firstRow = parseInt(refs[0].replace(/[A-Z]/g, ""), 10) - 1;
      var lastRow = parseInt(refs[refs.length - 1].replace(/[A-Z]/g, ""), 10) - 1;
      if (fmt.columnWidth !== undefined) apply("columnWidth", function () {
        for (var c = firstCol; c <= lastCol; c++) tbl.columns[c].width = fmt.columnWidth;
      });
      if (fmt.rowHeight !== undefined) apply("rowHeight", function () {
        for (var r = firstRow; r <= lastRow; r++) tbl.rows[r].height = fmt.rowHeight;
      });
      return JSON.stringify({ formatted: true, cellsFormatted: cells.length, applied: applied });
    `, { documentName, slideNumber, tableIndex: tableIndex ?? 0, cellRange, format })),
  );

  // ── Compound Tool ──

  server.tool(
    "keynote_create_presentation_with_slides",
    "Create a Keynote presentation with multiple fully-configured slides in one call (theme, layout, title, body, notes, transitions)",
    {
      themeName: z.string().optional().describe("Theme name (e.g. 'White', 'Black', 'Gradient')"),
      slides: z.array(z.object({
        masterSlideName: z.string().optional().describe("Master slide layout (e.g. 'Title & Subtitle', 'Bullets', 'Blank')"),
        title: z.string().optional().describe("Slide title text"),
        body: z.string().optional().describe("Slide body text (use newlines for bullet points)"),
        presenterNotes: z.string().optional().describe("Presenter notes"),
        transition: z.object({
          effect: z.string().describe("Transition effect (e.g. 'dissolve', 'push', 'wipe', 'magic move', 'none')"),
          duration: z.number().positive().optional().describe("Duration in seconds"),
        }).optional().describe("Slide transition"),
      })).min(1).describe("Array of slides to create"),
      ...SLIDE_SIZE_SCHEMA,
    },
    ANNOTATIONS.readWrite,
    async ({ themeName, slides, width, height }) => handleJXA(async () => rememberCreated(await runJXA<{ name: string }>(`
      const app = Application("Keynote");

      // Create presentation
      let doc;
      if (params.themeName) {
        doc = app.Document({ documentTheme: app.themes[params.themeName] });
        app.documents.push(doc);
      } else {
        doc = app.Document();
        app.documents.push(doc);
      }
      ${APPLY_SLIDE_SIZE}

      ${KEYNOTE_AS_HELPERS}
      ${PLACEHOLDER_HELPER}
      const docAS = "document " + asQuote(doc.name());
      function addSlideWithMaster(masterSlideName, i) {
        try {
          runAS('tell application "Keynote" to tell ' + docAS + ' to make new slide with properties {base slide:master slide ' + asQuote(masterSlideName) + '}');
        } catch (e) {
          throw new Error("Slide " + (i + 1) + ": " + e.message);
        }
      }
      const results = [];

      // New presentations always have 1 auto-generated slide.
      // Reuse it for the first user slide, then add the rest.
      for (let i = 0; i < params.slides.length; i++) {
        const s = params.slides[i];

        if (i === 0) {
          // Reuse the existing first slide — but if a master slide is requested,
          // add the new slide first, then delete the default one.
          if (s.masterSlideName) {
            addSlideWithMaster(s.masterSlideName, i);
            // Delete the original default slide (now at index 0)
            app.delete(doc.slides[0]);
          }
          // else: reuse the auto-generated slide as-is
        } else {
          // Add subsequent slides
          if (s.masterSlideName) {
            addSlideWithMaster(s.masterSlideName, i);
          } else {
            doc.slides.push(app.Slide({}));
          }
        }

        const slide = doc.slides[doc.slides.length - 1];

        // Set title
        if (s.title) {
          showPlaceholder(slide, i + 1, "titleShowing", "title");
          try { slide.defaultTitleItem().objectText = s.title; } catch(e) {}
        }

        // Set body
        if (s.body) {
          showPlaceholder(slide, i + 1, "bodyShowing", "body");
          try { slide.defaultBodyItem().objectText = s.body; } catch(e) {}
        }

        // Set presenter notes
        if (s.presenterNotes) {
          slide.presenterNotes = s.presenterNotes;
        }

        // Set transition
        if (s.transition) {
          slide.transitionProperties = {
            transitionEffect: s.transition.effect,
            transitionDuration: s.transition.duration || 1.0,
          };
        }

        let title = "";
        try { const t = slide.defaultTitleItem(); if (t) title = t.objectText(); } catch(e) {}
        let hasBody = false;
        try { const b = slide.defaultBodyItem(); if (b) hasBody = b.objectText().length > 0; } catch(e) {}
        let hasNotes = false;
        try { hasNotes = slide.presenterNotes().length > 0; } catch(e) {}

        results.push({ slideNumber: i + 1, title, hasBody, hasNotes });
      }

      return JSON.stringify({ name: doc.name(), slideCount: doc.slides.length, width: doc.width(), height: doc.height(), slides: results });
    `, {
      themeName: themeName ?? null,
      slides: slides.map(s => s.transition ? { ...s, transition: { ...s.transition, effect: normalizeTransitionEffect(s.transition.effect) } } : s),
      width: width ?? null,
      height: height ?? null,
    }))),
  );

  // ── Creator Studio Features (require subscription) ──

  server.tool(
    "keynote_clean_up_slide",
    "Auto-adjust a slide's layout, spacing, alignment, and typography using AI (Creator Studio only).",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideIndex: z.number().int().min(1).describe("Slide number (1-based)"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideIndex }) => {
      if (!isCreatorStudio()) {
        return toolResult("Clean Up Slide requires Apple Creator Studio (iWork 15.1+).", true);
      }
      return handleJXA(async () => {
        // Navigate to the target slide
        await runJXA<void>(`
          const app = Application("Keynote");
          app.activate();
          const doc = app.documents.byName(params.documentName);
          doc.currentSlide = doc.slides[params.slideIndex - 1];
        `, { documentName, slideIndex },
        { label: "clean_up_slide:navigate" });

        // Click Slide > Clean Up Slide
        await clickMenuItem("Keynote", ["Slide", "Clean Up Slide"], { postdelay: 3 });

        return JSON.stringify({ success: true, message: "Slide " + slideIndex + " cleaned up." });
      });
    },
  );

  server.tool(
    "keynote_super_resolution",
    "Upscale an image on a slide using AI Super Resolution (Creator Studio only). Increases resolution while preserving quality.",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideIndex: z.number().int().min(1).describe("Slide number (1-based)"),
      imageIndex: z.number().int().min(1).describe("Image index (1-based) on the slide"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideIndex, imageIndex }) => {
      if (!isCreatorStudio()) {
        return toolResult("Super Resolution requires Apple Creator Studio (iWork 15.1+).", true);
      }
      return handleJXA(async () => {
        await runJXA<void>(`
          const app = Application("Keynote");
          app.activate();
          const doc = app.documents.byName(params.documentName);
          doc.currentSlide = doc.slides[params.slideIndex - 1];
          const slide = doc.slides[params.slideIndex - 1];
          const images = slide.images();
          if (params.imageIndex > images.length) throw new Error("Image index " + params.imageIndex + " out of range (slide has " + images.length + " images)");
          app.selection = [images[params.imageIndex - 1]];
        `, { documentName, slideIndex, imageIndex },
        { label: "super_resolution:select" });

        await clickMenuItem("Keynote", ["Format", "Image", "Super Resolution"], { postdelay: 2 });

        return JSON.stringify({ success: true, message: "Super Resolution started on slide " + slideIndex + ", image " + imageIndex + "." });
      });
    },
  );

  server.tool(
    "keynote_remove_background",
    "Remove the background from an image on a slide using AI (Creator Studio only).",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideIndex: z.number().int().min(1).describe("Slide number (1-based)"),
      imageIndex: z.number().int().min(1).describe("Image index (1-based) on the slide"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideIndex, imageIndex }) => {
      if (!isCreatorStudio()) {
        return toolResult("Remove Background requires Apple Creator Studio (iWork 15.1+).", true);
      }
      return handleJXA(async () => {
        await runJXA<void>(`
          const app = Application("Keynote");
          app.activate();
          const doc = app.documents.byName(params.documentName);
          doc.currentSlide = doc.slides[params.slideIndex - 1];
          const slide = doc.slides[params.slideIndex - 1];
          const images = slide.images();
          if (params.imageIndex > images.length) throw new Error("Image index " + params.imageIndex + " out of range (slide has " + images.length + " images)");
          app.selection = [images[params.imageIndex - 1]];
        `, { documentName, slideIndex, imageIndex },
        { label: "remove_background:select" });

        await clickMenuItem("Keynote", ["Format", "Image", "Remove Background"], { postdelay: 2 });

        return JSON.stringify({ success: true, message: "Background removal started on slide " + slideIndex + ", image " + imageIndex + "." });
      });
    },
  );
}
