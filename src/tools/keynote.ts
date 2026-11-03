import { execFile } from "node:child_process";
import { existsSync, realpathSync } from "node:fs";
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
    "List all items on a slide with their types, positions, and sizes",
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

      try {
        var shapes = slide.shapes();
        for (var i = 0; i < shapes.length; i++) {
          var s = shapes[i];
          var entry = { type: "shape", index: i, width: s.width(), height: s.height(), position: pos(s) };
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
          var entry = { type: "textItem", index: i, width: t.width(), height: t.height(), position: pos(t) };
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
          items.push({ type: "image", index: i, width: img.width(), height: img.height(), position: pos(img) });
        }
      } catch(e) {}

      try {
        var tables = slide.tables();
        for (var i = 0; i < tables.length; i++) {
          var tbl = tables[i];
          items.push({ type: "table", index: i, name: tbl.name(), rows: tbl.rowCount(), columns: tbl.columnCount(), width: tbl.width(), height: tbl.height(), position: pos(tbl) });
        }
      } catch(e) {}

      try {
        var lines = slide.lines();
        for (var i = 0; i < lines.length; i++) {
          var ln = lines[i];
          items.push({ type: "line", index: i, startPoint: ln.startPoint(), endPoint: ln.endPoint() });
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
    "Move and/or resize an item (shape, image, or textItem) on a slide by type and index",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      itemType: z.enum(["shape", "image", "textItem"]).describe("Type of item"),
      itemIndex: z.number().int().min(0).describe("Item index (0-based)"),
      x: z.number().optional().describe("New X position in points"),
      y: z.number().optional().describe("New Y position in points"),
      width: z.number().positive().optional().describe("New width in points"),
      height: z.number().positive().optional().describe("New height in points"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, itemType, itemIndex, x, y, width, height }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];

      const collections = { shape: slide.shapes, image: slide.images, textItem: slide.textItems };
      const items = collections[params.itemType]();
      if (params.itemIndex >= items.length) {
        throw new Error(params.itemType + " index " + params.itemIndex + " out of range (slide has " + items.length + ")");
      }
      const item = items[params.itemIndex];

      const pos = item.position();
      const newX = params.x !== null ? params.x : pos.x;
      const newY = params.y !== null ? params.y : pos.y;
      item.position = {x: newX, y: newY};

      if (params.width !== null) item.width = params.width;
      if (params.height !== null) item.height = params.height;

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
    "Read shape properties: position, size, text, opacity, rotation, and text formatting",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      shapeIndex: z.number().int().min(0).describe("Shape index (0-based)"),
    },
    ANNOTATIONS.readOnly,
    async ({ documentName, slideNumber, shapeIndex }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
      const shapes = slide.shapes();
      if (params.shapeIndex >= shapes.length) {
        throw new Error("Shape index " + params.shapeIndex + " out of range (slide has " + shapes.length + " shapes)");
      }
      const shape = shapes[params.shapeIndex];

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
    `, { documentName, slideNumber, shapeIndex })),
  );

  server.tool(
    "keynote_format_shape",
    "Set shape properties: opacity, rotation, and text formatting (font, size, color). Note: fill and border colors and text alignment are not accessible via scripting.",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      shapeIndex: z.number().int().min(0).describe("Shape index (0-based)"),
      format: z.object({
        opacity: z.number().min(0).max(100).optional().describe("Opacity (0-100)"),
        rotation: z.number().optional().describe("Rotation in degrees"),
        fontName: z.string().optional().describe("PostScript font name for shape text"),
        fontSize: z.number().positive().optional().describe("Font size in points"),
        textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Text color as hex, e.g. '#FF0000'"),
      }).strict().describe("Format options"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, shapeIndex, format }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
      const shapes = slide.shapes();
      if (params.shapeIndex >= shapes.length) {
        throw new Error("Shape index " + params.shapeIndex + " out of range (slide has " + shapes.length + " shapes)");
      }
      const shape = shapes[params.shapeIndex];
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

      return JSON.stringify({ formatted: true, shapeIndex: params.shapeIndex, applied: applied });
    `, { documentName, slideNumber, shapeIndex, format })),
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
    "Format cells in a table on a slide (text color, number format)",
    {
      documentName: z.string().describe("Name of the open presentation"),
      slideNumber: z.number().int().min(1).describe("Slide number (1-based)"),
      tableIndex: z.number().int().min(0).optional().describe("Table index on the slide (0-based, default: 0)"),
      cellRange: z.string().regex(/^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$/).describe("Cell or range reference, e.g. 'A1' or 'A1:C3'"),
      format: z.object({
        textColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional().describe("Text color as hex, e.g. '#FF0000'"),
        numberFormat: z.enum(["automatic", "number", "currency", "percent", "fraction", "scientific", "text"]).optional().describe("Number format for the cells"),
      }).describe("Formatting options to apply"),
    },
    ANNOTATIONS.readWrite,
    async ({ documentName, slideNumber, tableIndex, cellRange, format }) => handleJXA(() => runJXA<string>(`
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      const slide = doc.slides[params.slideNumber - 1];
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
      var formatted = 0;
      for (var i = 0; i < cells.length; i++) {
        var cell = tbl.cells[cells[i]];
        if (params.format.textColor) {
          var hex = params.format.textColor;
          var r = parseInt(hex.slice(1,3), 16) * 257;
          var g = parseInt(hex.slice(3,5), 16) * 257;
          var b = parseInt(hex.slice(5,7), 16) * 257;
          cell.textColor = [r, g, b];
        }
        if (params.format.numberFormat) {
          cell.format = params.format.numberFormat;
        }
        formatted++;
      }
      return JSON.stringify({ formatted: true, cellsFormatted: formatted });
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
