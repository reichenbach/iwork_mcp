import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, rm, copyFile, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

// ── iWork app name resolution ──
// iWork 15.1 "Creator Studio" ships as separate bundles (e.g. "Numbers Creator Studio.app").
// Prefer Creator Studio when available; fall back to standard names.

const IWORK_APPS = ["Numbers", "Pages", "Keynote"] as const;
type IWorkApp = (typeof IWORK_APPS)[number];

let resolvedNames: Record<IWorkApp, string> | undefined;

function detectAppNames(): Record<IWorkApp, string> {
  const result = {} as Record<IWorkApp, string>;
  for (const app of IWORK_APPS) {
    result[app] = existsSync(`/Applications/${app} Creator Studio.app`)
      ? `${app} Creator Studio`
      : app;
  }
  return result;
}

/** Resolve the JXA application name for an iWork app. */
export function resolveAppName(app: IWorkApp): string {
  resolvedNames ??= detectAppNames();
  return resolvedNames[app];
}

/** Apply app-name substitutions to a script string. */
function rewriteAppNames(script: string): string {
  resolvedNames ??= detectAppNames();
  let s = script;
  for (const app of IWORK_APPS) {
    const resolved = resolvedNames[app];
    if (resolved === app) continue;
    // JXA: Application("Numbers") → Application("Numbers Creator Studio")
    s = s.replaceAll(`Application("${app}")`, `Application("${resolved}")`);
    // AppleScript (used in Keynote master slide ops):
    // application "Keynote" → application "Keynote Creator Studio"
    s = s.replaceAll(`application "${app}"`, `application "${resolved}"`);
  }
  return s;
}

/** Check if the current machine has Creator Studio iWork apps installed. */
export function isCreatorStudio(): boolean {
  resolvedNames ??= detectAppNames();
  return IWORK_APPS.some((app) => resolvedNames![app] !== app);
}

/** Map from iWork app name to System Events process name. */
export function processName(app: IWorkApp): string {
  return app; // System Events always uses the base name ("Keynote", not "Keynote Creator Studio")
}

/**
 * Click a menu item by name via System Events (resolution-independent).
 * menuPath is top-level menu followed by nested item names,
 * e.g. ["Table", "Magic Fill Cells"] or ["Format", "Image", "Super Resolution"].
 */
export async function clickMenuItem(
  app: IWorkApp,
  menuPath: string[],
  options?: { predelay?: number; postdelay?: number },
): Promise<void> {
  const proc = processName(app);
  const predelay = options?.predelay ?? 0.3;
  const postdelay = options?.postdelay ?? 0.5;

  // Build the AppleScript menu traversal
  // ["Format", "Image", "Super Resolution"] →
  //   menu item "Super Resolution" of menu "Image" of menu item "Image" of menu "Format" of menu bar 1
  let selector: string;
  if (menuPath.length === 2) {
    selector = `menu item "${menuPath[1]}" of menu "${menuPath[0]}" of menu bar 1`;
  } else if (menuPath.length === 3) {
    selector = `menu item "${menuPath[2]}" of menu "${menuPath[1]}" of menu item "${menuPath[1]}" of menu "${menuPath[0]}" of menu bar 1`;
  } else {
    throw new Error("clickMenuItem supports 2 or 3 level menu paths");
  }

  const script = `
    tell application "${resolveAppName(app)}" to activate
    delay ${predelay}
    tell application "System Events"
      tell process "${proc}"
        click ${selector}
      end tell
    end tell
    delay ${postdelay}
  `;

  await execFileAsync("/usr/bin/osascript", ["-e", script], { timeout: 15_000 });
}

/**
 * Creator Studio workaround for table insertion in Pages.
 * `doc.tables.push()` and AppleScript `make new table` both fail with -2763
 * ("Don't know how to create TMAScriptTableInfoProxy") on Pages 15.x, but
 * existing tables can be read/resized via AppleScript. So: insert via the
 * Insert > Table > Basic menu (System Events), then resize to the requested
 * dimensions. An insertion point in the body text is required to enable the
 * menu — established by clicking into the body, scanning top-down (a click
 * can land on a table/image, which selects it and disables Insert > Table).
 * Synthetic keystrokes are deliberately avoided: secure keyboard entry (e.g.
 * a password field focused in another app) silently blocks them, but not
 * mouse clicks or menu clicks.
 * Returns the new table's name.
 */
export async function pagesInsertTableViaUI(
  documentName: string,
  rows?: number,
  columns?: number,
): Promise<string> {
  // Resolve the document name first (auto-save may have appended ".pages")
  const resolvedDocName = await runJXA<string>(
    `
    const app = Application("Pages");
    return JSON.stringify(app.documents.byName(params.documentName).name());
    `,
    { documentName },
    { label: "pagesInsertTableViaUI:resolveName" },
  );

  const script = `
on run argv
  set docName to item 1 of argv
  set newRows to (item 2 of argv) as integer
  set newCols to (item 3 of argv) as integer
  tell application "${resolveAppName("Pages")}"
    activate
    try
      set index of window docName to 1
    on error
      try
        set index of window (docName & ".pages") to 1
      end try
    end try
    set frontName to name of front document
    if frontName is not docName and frontName is not (docName & ".pages") then
      error "Could not bring document " & docName & " to front (front document is " & frontName & ")"
    end if
    set beforeNames to name of every table of front document
  end tell
  tell application "System Events"
    tell process "${processName("Pages")}"
      repeat 10 times
        set frontmost to true
        delay 0.3
        if frontmost then exit repeat
      end repeat
      set sa to scroll area 1 of splitter group 1 of window 1
      set {sx, sy} to position of sa
      set {sw, sh} to size of sa
      set ok to false
      set yOff to 90
      repeat while yOff < (sh - 100)
        click at {sx + (sw div 2), sy + yOff}
        delay 0.4
        if enabled of menu item "Basic" of menu "Table" of menu item "Table" of menu "Insert" of menu bar 1 then
          set ok to true
          exit repeat
        end if
        set yOff to yOff + 70
      end repeat
      if not ok then error "Could not place an insertion point in the document body to insert the table. Make sure this app has Accessibility permission (System Settings > Privacy & Security > Accessibility) and the document body is editable."
      click menu item "Table" of menu "Insert" of menu bar 1
      delay 0.4
      click menu item "Basic" of menu "Table" of menu item "Table" of menu "Insert" of menu bar 1
    end tell
  end tell
  delay 1.0
  tell application "${resolveAppName("Pages")}"
    tell front document
      set afterNames to name of every table
      if (count of afterNames) is not ((count of beforeNames) + 1) then
        error "Table insertion via UI scripting did not take effect. Make sure this app has Accessibility permission (System Settings > Privacy & Security > Accessibility)."
      end if
      set newName to ""
      repeat with n in afterNames
        if beforeNames does not contain (n as text) then set newName to (n as text)
      end repeat
      if newRows > 0 then set row count of table newName to newRows
      if newCols > 0 then set column count of table newName to newCols
      return newName
    end tell
  end tell
end run
`;

  try {
    const { stdout } = await execFileAsync(
      "/usr/bin/osascript",
      ["-e", script, resolvedDocName, String(rows ?? 0), String(columns ?? 0)],
      { timeout: 30_000 },
    );
    return stdout.trim();
  } catch (err) {
    const e = err as { stderr?: string; message?: string; code?: number };
    throw new OsascriptError(e.stderr || e.message || String(err), typeof e.code === "number" ? e.code : null);
  }
}

// Documents this server created that haven't been saved to a user path yet
const createdDocuments = new Set<string>();

/** Record a document created by a tool, so its untitled auto-save can be cleaned up. */
export function rememberCreatedDocument(appName: IWorkApp, name: string): void {
  createdDocuments.add(createdKey(appName, name));
}

function createdKey(appName: IWorkApp, name: string): string {
  const ext = EXT_MAP[appName];
  return `${appName}:${name.endsWith(ext) ? name.slice(0, -ext.length) : name}`;
}

/** The iCloud auto-save of a document a tool created, or null. */
export async function untitledAutoSavePath(appName: IWorkApp, documentName: string): Promise<string | null> {
  if (!createdDocuments.has(createdKey(appName, documentName))) return null;
  const path = await runJXA<string | null>(
    `
    const f = Application("${resolveAppName(appName)}").documents.byName(params.documentName).file();
    return JSON.stringify(f ? f.toString() : null);
    `,
    { documentName },
    { label: "untitledAutoSavePath" },
  ).catch(() => null);
  return path && path.includes("/Library/Mobile Documents/") ? path : null;
}

/** Move a tool-created document's untitled auto-save to the Trash once no open document uses it. */
export async function trashUntitledAutoSave(appName: IWorkApp, documentName: string, path: string): Promise<void> {
  createdDocuments.delete(createdKey(appName, documentName));
  // Waiting first, or Keynote may report the document can't be opened
  await runJXA<void>(
    `
    ObjC.import("Foundation");
    const app = Application("${resolveAppName(appName)}");
    function inUse() {
      return app.documents().some(function (d) {
        try { const f = d.file(); return f && f.toString() === params.path; } catch (e) { return false; }
      });
    }
    for (let i = 0; i < 10 && inUse(); i++) delay(0.5);
    if (inUse()) return;
    $.NSFileManager.defaultManager.trashItemAtURLResultingItemURLError($.NSURL.fileURLWithPath(params.path), null, null);
    `,
    { path },
    { label: "trashUntitledAutoSave" },
  ).catch(() => {});
}

/**
 * Creator Studio workaround for "Save As" (doc.save({in:}) hangs).
 * Copies the auto-saved iCloud file, then closes + reopens from the new path.
 * Returns the new document name.
 */
export async function creatorStudioSaveAs(
  appName: IWorkApp,
  documentName: string,
  targetPath: string,
): Promise<string> {
  // Get the auto-save path, close with save to flush changes, copy, and reopen
  const autoSavePath = await runJXA<string>(
    `
    const app = Application("${resolveAppName(appName)}");
    const doc = app.documents.byName(params.documentName);
    // Poll for file() to become available (auto-save can take a moment)
    for (let i = 0; i < 15; i++) {
      const f = doc.file();
      if (f) return f.toString();
      delay(1);
    }
    throw new Error("Timed out waiting for document auto-save path");
    `,
    { documentName },
    { timeout: 20_000, label: "creatorStudioSaveAs:getFile" },
  );

  // Close with saving:"yes" to flush all in-memory changes to the auto-save location
  await runJXA<void>(
    `
    const app = Application("${resolveAppName(appName)}");
    const doc = app.documents.byName(params.documentName);
    doc.close({ saving: "yes" });
    `,
    { documentName },
    { timeout: 10_000, label: "creatorStudioSaveAs:close" },
  );

  // Copy the saved file to the target path
  await execFileAsync("/bin/cp", ["-R", autoSavePath, targetPath]);

  // Reopen from the new path
  const newName = await runJXA<string>(
    `
    const app = Application("${resolveAppName(appName)}");
    const newDoc = app.open(Path(params.targetPath));
    return JSON.stringify(newDoc.name());
    `,
    { targetPath },
    { label: "creatorStudioSaveAs:reopen" },
  );

  // Untitled documents are auto-saved to iCloud before Save As; don't leave that copy behind
  if (createdDocuments.has(createdKey(appName, documentName)) && autoSavePath.includes("/Library/Mobile Documents/") && autoSavePath !== targetPath) {
    await trashUntitledAutoSave(appName, documentName, autoSavePath);
  }

  return newName;
}

/**
 * Creator Studio workaround for PDF export (app.export() fails with error 6).
 * Uses qlmanage Quick Look to generate a PDF from the auto-saved file.
 */
export async function creatorStudioExportPDF(
  appName: IWorkApp,
  documentName: string,
  targetPath: string,
): Promise<void> {
  // Get the auto-save file path
  const autoSavePath = await runJXA<string>(
    `
    const app = Application("${resolveAppName(appName)}");
    const doc = app.documents.byName(params.documentName);
    for (let i = 0; i < 15; i++) {
      const f = doc.file();
      if (f) return f.toString();
      delay(1);
    }
    throw new Error("Timed out waiting for document auto-save path");
    `,
    { documentName },
    { timeout: 20_000, label: "creatorStudioExportPDF:getFile" },
  );

  // Generate Quick Look preview (contains a PDF attachment)
  const tmpDir = await mkdtemp(join(tmpdir(), "iwork-mcp-ql-"));
  try {
    await execFileAsync("/usr/bin/qlmanage", ["-p", "-o", tmpDir, autoSavePath], {
      timeout: 30_000,
    });

    // Find the generated PDF
    const { stdout } = await execFileAsync("/usr/bin/find", [tmpDir, "-name", "*.pdf", "-type", "f"]);
    const pdfPath = stdout.trim().split("\n")[0];
    if (!pdfPath) {
      throw new Error("Quick Look did not generate a PDF preview for this document");
    }

    await copyFile(pdfPath, targetPath);
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
}

/**
 * Creator Studio PDF export for Keynote (Quick Look renders only one blank page for .key).
 * Tries the native export, then falls back to slide images assembled with PDFKit (raster pages).
 */
export async function creatorStudioKeynoteExportPDF(
  documentName: string,
  targetPath: string,
): Promise<{ method: "native" | "slide images"; pageCount?: number }> {
  const mtimeBefore = existsSync(targetPath) ? (await stat(targetPath)).mtimeMs : -1;
  try {
    await runJXA<void>(
      `
      const app = Application("Keynote");
      app.export(app.documents.byName(params.documentName), { to: Path(params.targetPath), as: "PDF" });
      `,
      { documentName, targetPath },
      { timeout: 60_000, label: "creatorStudioKeynoteExportPDF:native" },
    );
    if (existsSync(targetPath) && (await stat(targetPath)).mtimeMs > mtimeBefore) return { method: "native" };
  } catch {}

  // Next to the target: the sandboxed app may not be able to write to $TMPDIR
  const tmpDir = await mkdtemp(join(dirname(targetPath), ".iwork-mcp-export-"));
  try {
    await runJXA<void>(
      `
      const app = Application("Keynote");
      const doc = app.documents.byName(params.documentName);
      // No imageFormat: JXA can't pass that enum (-1700), and PNG is the default
      app.export(doc, { to: Path(params.folder), as: "slide images", withProperties: { skippedSlides: false } });
      `,
      { documentName, folder: join(tmpDir, "slides") },
      { timeout: 120_000, label: "creatorStudioKeynoteExportPDF:images" },
    );

    const { stdout } = await execFileAsync("/usr/bin/find", [tmpDir, "-type", "f"]);
    const images = stdout
      .split("\n")
      .filter((f) => /\.(png|jpe?g|tiff?)$/i.test(f))
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    if (images.length === 0) {
      throw new Error("Keynote did not export any slide images");
    }

    const pageCount = await runJXA<number>(
      `
      ObjC.import("AppKit");
      ObjC.import("Quartz");
      const pdf = $.PDFDocument.alloc.init;
      for (let i = 0; i < params.images.length; i++) {
        const img = $.NSImage.alloc.initWithContentsOfFile(params.images[i]);
        if (!img || img.isNil()) throw new Error("Could not read exported slide image " + params.images[i]);
        pdf.insertPageAtIndex($.PDFPage.alloc.initWithImage(img), i);
      }
      if (!pdf.writeToFile(params.targetPath)) throw new Error("Could not write PDF to " + params.targetPath);
      return JSON.stringify(Number(pdf.pageCount));
      `,
      { images, targetPath },
      { timeout: 60_000, label: "creatorStudioKeynoteExportPDF:assemble" },
    );
    return { method: "slide images", pageCount };
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
}

// ── Creator Studio document name resolution ──
// Creator Studio auto-saves rename documents by appending file extensions
// (e.g. "Untitled 1" → "Untitled 1.numbers"), breaking documents.byName() lookups.
// We inject a resolution block that tries the extended name on -1728 errors.

const EXT_MAP: Record<IWorkApp, string> = {
  Numbers: ".numbers",
  Pages: ".pages",
  Keynote: ".key",
};

function detectCreatorStudioAppFromScript(script: string): IWorkApp | undefined {
  for (const app of IWORK_APPS) {
    const csName = `${app} Creator Studio`;
    if (script.includes(`Application("${csName}")`)) return app;
    if (script.includes(`application "${csName}"`)) return app;
  }
  return undefined;
}

/**
 * Inject document-name resolution into a JXA script.
 *
 * After `const params = JSON.parse(argv[0]);`, inserts a block that:
 * 1. Tries `app.documents.byName(params.documentName)` — if it works, done.
 * 2. On failure, tries with the file extension appended.
 * 3. If that works, mutates `params.documentName` so all downstream code uses it.
 *
 * Guards: only activates when Creator Studio is detected, the script references
 * `params.documentName`, and the app can be identified.
 */
export function injectDocumentNameResolution(script: string): string {
  // Only relevant for scripts that use params.documentName
  if (!script.includes("params.documentName")) return script;

  // Detect Creator Studio app from the script content (not machine state)
  const app = detectCreatorStudioAppFromScript(script);
  if (!app) return script;

  const ext = EXT_MAP[app];
  const csName = `${app} Creator Studio`;
  const appExpr = `Application("${csName}")`;

  const injection = `
  // [iwork-mcp] Creator Studio document name resolution
  try {
    const __app = ${appExpr};
    __app.documents.byName(params.documentName).name();
  } catch(__e) {
    try {
      const __app2 = ${appExpr};
      const __extName = params.documentName + ${JSON.stringify(ext)};
      __app2.documents.byName(__extName).name();
      params.documentName = __extName;
    } catch(__e2) { /* leave params unchanged; original error will surface */ }
  }`;

  // Insert after the params line
  const anchor = "const params = JSON.parse(argv[0]);";
  if (!script.includes(anchor)) return script;

  return script.replace(anchor, anchor + injection);
}

const COMMON_ERRORS: Record<number, string> = {
  [-1743]: "Permission denied. Open System Settings → Privacy & Security → Automation and allow this app to control the iWork application.",
  [-1728]: "Element not found. The specified document, sheet, table, or cell does not exist.",
  [-128]: "User cancelled the operation.",
  [-10810]: "The application is not running and could not be launched.",
  [-1700]: "Invalid data type for the operation.",
  [-1708]: "The application does not understand this command.",
};

export class OsascriptError extends Error {
  readonly appleScriptErrorCode: number | undefined;
  readonly stderr: string;

  constructor(stderr: string, exitCode: number | null) {
    const errorCode = parseErrorCode(stderr);
    const friendly = errorCode !== undefined ? COMMON_ERRORS[errorCode] : undefined;
    const message = friendly
      ? `${friendly}\n\nosascript stderr: ${stderr}`
      : `osascript failed (exit ${exitCode}): ${stderr}`;

    super(message);
    this.name = "OsascriptError";
    this.appleScriptErrorCode = errorCode;
    this.stderr = stderr;
  }
}

function parseErrorCode(stderr: string): number | undefined {
  // JXA errors look like: "Error: Error: ... (-1743)"
  // or "execution error: ... (-1728)"
  const match = stderr.match(/\((-?\d+)\)\s*$/);
  return match ? parseInt(match[1], 10) : undefined;
}

export interface RunJXAOptions {
  timeout?: number;
  maxBuffer?: number;
  label?: string;
}

/**
 * Execute a JXA (JavaScript for Automation) script via osascript.
 *
 * The script body is wrapped in `function run(argv) { ... }` automatically.
 * If `params` is provided, it's JSON-serialized and passed as argv[0].
 * The script should return a value; it will be JSON.stringify'd and parsed back.
 */
export function runJXA<T = unknown>(
  scriptBody: string,
  params?: Record<string, unknown>,
  options?: RunJXAOptions,
): Promise<T> {
  const timeout = options?.timeout ?? 30_000;
  const maxBuffer = options?.maxBuffer ?? 10 * 1024 * 1024;
  const timing = process.env.IWORK_MCP_TIMING === "1";
  const t0 = timing ? Date.now() : 0;

  // Wrap in run(argv) so osascript calls it with our args
  const fullScript = `
function run(argv) {
  ${params !== undefined ? "const params = JSON.parse(argv[0]);" : ""}
  ${scriptBody}
}
`;

  const rewritten = injectDocumentNameResolution(rewriteAppNames(fullScript));
  const args = ["-l", "JavaScript", "-e", rewritten];
  if (params !== undefined) {
    args.push(JSON.stringify(params));
  }

  const label = options?.label ?? "jxa";

  return new Promise<T>((resolve, reject) => {
    execFile("/usr/bin/osascript", args, { timeout, maxBuffer }, (error, stdout, stderr) => {
      if (timing) {
        const ms = Date.now() - t0;
        process.stderr.write(`[iwork-mcp] ${label} completed in ${ms}ms\n`);
      }

      if (error) {
        const exitCode = typeof error.code === "number" ? error.code : null;
        reject(new OsascriptError(stderr || error.message, exitCode));
        return;
      }

      const trimmed = stdout.trim();
      if (!trimmed) {
        resolve(undefined as T);
        return;
      }

      try {
        resolve(JSON.parse(trimmed) as T);
      } catch {
        // If the output isn't JSON, return the raw string
        resolve(trimmed as T);
      }
    });
  });
}
