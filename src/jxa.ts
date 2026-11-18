import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, rm, copyFile, stat, readdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, dirname, join } from "node:path";
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

// ── Menus in the app's UI language ──
// Tools name menu items by their English titles. When the app runs in another language,
// the titles are translated with the app's own MainMenu.strings: en_GB/en_AU hold the
// English titles under the same keys as every other localization (Base is a compiled nib).

type MenuStrings = Record<string, string>;
const menuStringsCache = new Map<string, Promise<MenuStrings | null>>();
const resolvedMenuPaths = new Map<string, string[]>();
const US = "\u001f"; // separates titles in osascript arguments and results

function readMenuStrings(file: string): Promise<MenuStrings | null> {
  let p = menuStringsCache.get(file);
  if (!p) {
    p = execFileAsync("/usr/bin/plutil", ["-convert", "json", "-o", "-", file], { timeout: 10_000, maxBuffer: 8 * 1024 * 1024 })
      .then(({ stdout }) => JSON.parse(stdout) as MenuStrings)
      .catch(() => null);
    menuStringsCache.set(file, p);
  }
  return p;
}

const normalizeTitle = (t: string) => t.replace(/(…|\.\.\.)$/, "").trim().toLowerCase();

/** For each English title, the title itself plus its translations in the UI language (best match for the menu bar). */
export async function menuTitleCandidates(app: IWorkApp, englishPath: string[], menuBarTitles: string[]): Promise<string[][]> {
  const res = `/Applications/${resolveAppName(app)}.app/Contents/Resources`;
  let english: MenuStrings | null = null;
  for (const l of ["en_GB", "en_AU", "en"]) {
    english = await readMenuStrings(`${res}/${l}.lproj/MainMenu.strings`);
    if (english) break;
  }
  let local: MenuStrings | null = null;
  if (english) {
    const dirs = (await readdir(res).catch(() => [] as string[])).filter((d) => d.endsWith(".lproj"));
    const tables = await Promise.all(dirs.map((d) => readMenuStrings(`${res}/${d}/MainMenu.strings`)));
    let bestScore = 0;
    for (const t of tables) {
      if (!t) continue;
      const titles = new Set(Object.values(t));
      const score = menuBarTitles.filter((b) => titles.has(b)).length;
      if (score > bestScore) { bestScore = score; local = t; }
    }
  }
  return englishPath.map((name) => {
    const out = [name];
    if (english && local) {
      for (const [key, value] of Object.entries(english)) {
        const translated = local[key];
        if (translated && normalizeTitle(value) === normalizeTitle(name) && !out.includes(translated)) out.push(translated);
      }
    }
    return out;
  });
}

/** Find the first existing menu path among per-level candidate titles. */
async function findMenuPath(app: IWorkApp, candidates: string[][]): Promise<{ path: string[] | null; menuBarTitles: string[] }> {
  const script = `
on splitText(t)
  set AppleScript's text item delimiters to (character id 31)
  set r to text items of t
  set AppleScript's text item delimiters to ""
  return r
end splitText
on run argv
  set us to character id 31
  tell application "System Events" to tell process "${processName(app)}"
    repeat with a in my splitText(item 1 of argv)
      set a to a as text
      if exists menu a of menu bar 1 then
        if (count of argv) is 1 then return a
        repeat with b in my splitText(item 2 of argv)
          set b to b as text
          if exists menu item b of menu a of menu bar 1 then
            if (count of argv) is 2 then return a & us & b
            repeat with c in my splitText(item 3 of argv)
              set c to c as text
              if exists menu item c of menu b of menu item b of menu a of menu bar 1 then return a & us & b & us & c
            end repeat
          end if
        end repeat
      end if
    end repeat
    set AppleScript's text item delimiters to us
    set bar to (name of every menu bar item of menu bar 1) as text
    set AppleScript's text item delimiters to ""
    return (character id 30) & bar
  end tell
end run
`;
  let out: string;
  try {
    ({ stdout: out } = await execFileAsync("/usr/bin/osascript", ["-e", script, ...candidates.map((c) => c.join(US))], { timeout: 15_000 }));
  } catch (err) {
    const e = err as { stderr?: string; message?: string; code?: number };
    throw new OsascriptError(e.stderr || e.message || String(err), typeof e.code === "number" ? e.code : null);
  }
  out = out.replace(/\n$/, "");
  if (out.startsWith("\u001e")) return { path: null, menuBarTitles: out.slice(1).split(US) };
  return { path: out.split(US), menuBarTitles: [] };
}

/** Resolve an English menu path to the titles the app shows (cached). */
export async function resolveMenuPath(app: IWorkApp, menuPath: string[]): Promise<string[]> {
  if (menuPath.length < 1 || menuPath.length > 3) throw new Error("Menu paths have 1 to 3 levels");
  const key = [app, ...menuPath].join(US);
  const cached = resolvedMenuPaths.get(key);
  if (cached) return cached;
  let found = await findMenuPath(app, menuPath.map((t) => [t]));
  if (!found.path) {
    const candidates = await menuTitleCandidates(app, menuPath, found.menuBarTitles);
    if (candidates.some((c) => c.length > 1)) found = await findMenuPath(app, candidates);
  }
  if (!found.path) {
    throw new Error(`Menu item ${menuPath.join(" > ")} not found in ${resolveAppName(app)}, neither by its English title nor in the app's UI language.`);
  }
  resolvedMenuPaths.set(key, found.path);
  return found.path;
}

/** Forget cached menu titles (e.g. after the app was restarted in another language). */
export function forgetMenuPath(app: IWorkApp, menuPath: string[]): void {
  resolvedMenuPaths.delete([app, ...menuPath].join(US));
}

/** AppleScript reference to a menu item from resolved titles. */
export function menuItemRef(path: string[]): string {
  const q = (t: string) => `"${t.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
  if (path.length === 2) return `menu item ${q(path[1])} of menu ${q(path[0])} of menu bar 1`;
  if (path.length === 3) return `menu item ${q(path[2])} of menu ${q(path[1])} of ${menuItemRef(path.slice(0, 2))}`;
  throw new Error("Menu items have 2 or 3 levels");
}

/**
 * Click a menu item by its English titles via System Events, in any UI language.
 * menuPath is top-level menu followed by nested item names,
 * e.g. ["Table", "Magic Fill Cells"] or ["Format", "Image", "Super Resolution"].
 * Submenu items are clicked in two steps (open the parent, then the item): a direct
 * click on a third-level item silently does nothing. A disabled item is an error,
 * checked before any menu opens. If the click opens a sheet in the front window (for
 * example a subscription prompt), the sheet is cancelled and its text becomes the error.
 */
export async function clickMenuItem(
  app: IWorkApp,
  menuPath: string[],
  options?: { predelay?: number; postdelay?: number },
): Promise<void> {
  if (menuPath.length < 2 || menuPath.length > 3) throw new Error("clickMenuItem supports 2 or 3 level menu paths");
  const predelay = options?.predelay ?? 0.3;
  const postdelay = options?.postdelay ?? 0.5;

  for (let attempt = 0; ; attempt++) {
    const path = await resolveMenuPath(app, menuPath);
    const leaf = menuItemRef(path);
    // A sheet that opens instead (e.g. "Subscribe to Apple Creator Studio" for the AI image
    // tools) blocks every later call: cancel it and report its text.
    const script = `
    tell application "${resolveAppName(app)}" to activate
    delay ${predelay}
    tell application "System Events"
      tell process "${processName(app)}"
        if not (enabled of ${leaf}) then error "DISABLED"
        set w to missing value
        set sheetsBefore to 0
        try
          set w to window 1
          set sheetsBefore to count of sheets of w
        end try
        ${path.length === 3 ? `click ${menuItemRef(path.slice(0, 2))}
        delay 0.4` : ""}
        click ${leaf}
      end tell
    end tell
    repeat ${Math.max(1, Math.round(postdelay / 0.25))} times
      delay 0.25
      if w is not missing value then
        tell application "System Events"
          if (count of sheets of w) > sheetsBefore then
            delay 0.3
            set s to sheet 1 of w
            set msg to ""
            try
              set AppleScript's text item delimiters to " "
              set msg to (value of every static text of s) as text
              set AppleScript's text item delimiters to ""
            end try
            try
              click (value of attribute "AXCancelButton" of s)
            end try
            return "SHEET" & (character id 30) & msg
          end if
        end tell
      end if
    end repeat
    return ""
  `;
    let stdout: string;
    try {
      ({ stdout } = await execFileAsync("/usr/bin/osascript", ["-e", script], { timeout: 15_000 + postdelay * 1000 }));
    } catch (err) {
      const e = err as { stderr?: string; message?: string; code?: number };
      const stderr = e.stderr || e.message || String(err);
      if (stderr.includes("DISABLED")) {
        throw new Error(`Menu item ${menuPath.join(" > ")} (${path.join(" > ")}) is disabled for the current selection.`);
      }
      if (attempt === 0 && /\(-1728\)\s*$/.test(stderr.trim())) {
        forgetMenuPath(app, menuPath); // titles changed since they were cached
        continue;
      }
      throw new OsascriptError(stderr, typeof e.code === "number" ? e.code : null);
    }
    if (stdout.startsWith("SHEET\u001e")) {
      const text = stdout.slice(6).trim();
      throw new Error(`${resolveAppName(app)} answered ${menuPath.join(" > ")} with a dialog instead of doing it${text ? `: "${text}"` : ""}. The dialog was cancelled.`);
    }
    return;
  }
}

/** AppleScript string literal. */
export function asString(t: string): string {
  return `"${t.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

/**
 * Select one item of a document and bring its window to the front, so that a menu
 * command applies to it. JXA can't set the selection on 15.x (`app.selection = [...]`
 * fails with -1728, `doc.selection = [...]` with -1700), but AppleScript
 * `set selection of document ... to {...}` works in all three apps. Menus act on the
 * front window, so it is raised first. (Reading the selection back fails in Numbers.)
 * `itemRef` is relative to the document, e.g. `image 2` or `image 1 of slide 3`;
 * `prepare` runs first with the document as `d` (e.g. to show the item's slide).
 */
export async function selectForMenu(app: IWorkApp, documentName: string, itemRef: string, prepare = ""): Promise<void> {
  const script = `
on run argv
  set docName to item 1 of argv
  tell application "${resolveAppName(app)}"
    activate
    set d to missing value
    repeat with nm in {docName, docName & "${EXT_MAP[app]}"}
      if exists document (nm as text) then
        set d to document (nm as text)
        exit repeat
      end if
    end repeat
    if d is missing value then error "Document not found: " & docName
    set docId to id of d
    repeat with w in windows
      try
        if id of document of w is docId then
          set index of w to 1
          exit repeat
        end if
      end try
    end repeat
    ${prepare}
    delay 0.3
    set selection of d to {${itemRef} of d}
  end tell
end run
`;
  try {
    await execFileAsync("/usr/bin/osascript", ["-e", script, documentName], { timeout: 15_000 });
  } catch (err) {
    const e = err as { stderr?: string; message?: string; code?: number };
    throw new OsascriptError(e.stderr || e.message || String(err), typeof e.code === "number" ? e.code : null);
  }
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

  // Insert > Table > Basic, in the app's UI language
  const [insertMenu, tableItem, basicItem] = await resolveMenuPath("Pages", ["Insert", "Table", "Basic"]);

  const script = `
on run argv
  set docName to item 1 of argv
  set newRows to (item 2 of argv) as integer
  set newCols to (item 3 of argv) as integer
  set insertMenu to item 4 of argv
  set tableItem to item 5 of argv
  set basicItem to item 6 of argv
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
        if enabled of menu item basicItem of menu tableItem of menu item tableItem of menu insertMenu of menu bar 1 then
          set ok to true
          exit repeat
        end if
        set yOff to yOff + 70
      end repeat
      if not ok then error "Could not place an insertion point in the document body to insert the table. Make sure this app has Accessibility permission (System Settings > Privacy & Security > Accessibility) and the document body is editable."
      click menu item tableItem of menu insertMenu of menu bar 1
      delay 0.4
      click menu item basicItem of menu tableItem of menu item tableItem of menu insertMenu of menu bar 1
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
      ["-e", script, resolvedDocName, String(rows ?? 0), String(columns ?? 0), insertMenu, tableItem, basicItem],
      { timeout: 30_000 },
    );
    return stdout.trim();
  } catch (err) {
    const e = err as { stderr?: string; message?: string; code?: number };
    throw new OsascriptError(e.stderr || e.message || String(err), typeof e.code === "number" ? e.code : null);
  }
}

// ── Timeouts ──
// osascript runs are killed (SIGTERM, "exit null") when they exceed their timeout.
// Saves of large documents need far longer than ordinary reads, so both limits are
// overridable via environment variables.

export const DEFAULT_JXA_TIMEOUT_MS = 30_000;
export const DEFAULT_SAVE_TIMEOUT_MS = 120_000;

/** Read a positive millisecond value from an env var, falling back to a default. */
export function envTimeout(name: string, fallback: number): number {
  const raw = process.env[name];
  if (!raw) return fallback;
  const n = Number(raw);
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : fallback;
}

/** Timeout for ordinary JXA calls (IWORK_MCP_TIMEOUT_MS, default 30s). */
export function jxaTimeout(): number {
  return envTimeout("IWORK_MCP_TIMEOUT_MS", DEFAULT_JXA_TIMEOUT_MS);
}

/** Timeout for save operations (IWORK_MCP_SAVE_TIMEOUT_MS, default 120s). */
export function saveTimeout(): number {
  return envTimeout("IWORK_MCP_SAVE_TIMEOUT_MS", DEFAULT_SAVE_TIMEOUT_MS);
}

/** If `path` was modified at or after `sinceMs`, return a size+mtime signature; otherwise null. */
async function writtenSince(path: string, sinceMs: number): Promise<string | null> {
  try {
    const s = await stat(path);
    return s.mtimeMs >= sinceMs ? `${s.size}:${s.mtimeMs}` : null;
  } catch {
    return null;
  }
}

/**
 * Whether a document file is an untitled draft: auto-saved by the app (iCloud, Creator Studio)
 * but never saved by the user. AppKit marks those with the com.apple.document.draft attribute
 * and removes it on the first real save. A plain save of a draft opens a Save sheet and blocks.
 */
export async function isDraftFile(path: string): Promise<boolean> {
  try {
    const { stdout } = await execFileAsync("/usr/bin/xattr", [path], { timeout: 5_000 });
    return stdout.split("\n").some((name) => name.startsWith("com.apple.document.draft"));
  } catch {
    return false;
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

/** The file of an open document, or null if it has none (or the lookup fails). */
async function documentFile(appName: IWorkApp, documentName: string): Promise<string | null> {
  return runJXA<string | null>(
    `
    const f = Application("${resolveAppName(appName)}").documents.byName(params.documentName).file();
    return JSON.stringify(f ? f.toString() : null);
    `,
    { documentName },
    { label: "documentFile" },
  ).catch(() => null);
}

/**
 * The iCloud auto-save of an untitled document, or null: a draft (new or imported,
 * never saved — also one created before a server restart), or a document a tool created.
 */
export async function untitledAutoSavePath(appName: IWorkApp, documentName: string): Promise<string | null> {
  const path = await documentFile(appName, documentName);
  if (!path || !path.includes("/Library/Mobile Documents/")) return null;
  return createdDocuments.has(createdKey(appName, documentName)) || await isDraftFile(path) ? path : null;
}

/**
 * Close a document. A draft can't be saved in place (a Save sheet would open and block),
 * so saving "yes" on one is an error; saving "no" also moves its iCloud draft to the
 * Trash, which Creator Studio would otherwise keep (closing it in the app deletes it).
 */
export async function closeDocument(appName: IWorkApp, documentName: string, saving?: "yes" | "no" | "ask"): Promise<string> {
  if (saving === "yes") {
    const file = await documentFile(appName, documentName);
    if (file && await isDraftFile(file)) {
      throw new Error('The document has never been saved (it is new or was imported from another format); save it with a filePath first, or close it with saving "no"');
    }
  }
  const autoSave = saving === "no" ? await untitledAutoSavePath(appName, documentName) : null;
  const result = await runJXA<string>(
    `
    const doc = Application("${resolveAppName(appName)}").documents.byName(params.documentName);
    if (params.saving) doc.close({ saving: params.saving });
    else doc.close();
    return JSON.stringify(params.trashDraft ? { closed: true, draftTrashed: true } : { closed: true });
    `,
    { documentName, saving: saving ?? null, trashDraft: Boolean(autoSave) },
    { label: "closeDocument" },
  );
  if (autoSave) await trashUntitledAutoSave(appName, documentName, autoSave);
  return result;
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
 *
 * Uses `close saving yes saving in <path>`, which writes the file directly — it
 * works for brand-new documents that have no auto-save file yet (the old
 * copy-the-auto-save approach polled forever on those and got killed). The app
 * sometimes writes the file but never replies to the Apple Event, so the target
 * file is watched in parallel and the save counts as done once it is written and
 * stable. The document is then reopened via LaunchServices (`open -a`), because
 * JXA `app.open()` also never replies on Creator Studio.
 * Returns the document name, and whether the app actually reopened it.
 */
export async function creatorStudioSaveAs(
  appName: IWorkApp,
  documentName: string,
  targetPath: string,
): Promise<{ name: string; reopened: boolean }> {
  const timeout = saveTimeout();
  const start = Date.now();
  // Untitled documents are auto-saved to iCloud before Save As; don't leave that copy behind
  const autoSave = await untitledAutoSavePath(appName, documentName);

  const closing = runJXA<void>(
    `
    const app = Application("${resolveAppName(appName)}");
    const doc = app.documents.byName(params.documentName);
    doc.close({ saving: "yes", savingIn: Path(params.targetPath) });
    `,
    { documentName, targetPath },
    { timeout, label: "creatorStudioSaveAs:close" },
  ).then(() => null, (err: unknown) => err);

  // Wait for either the close to reply, or the file to be written and stable.
  let lastSig: string | null = null;
  let stableSince = 0;
  for (;;) {
    const done = await Promise.race([closing, sleep(500).then(() => undefined)]);
    if (done !== undefined) {
      if (await writtenSince(targetPath, start)) break; // a failed reply is fine if the file landed
      throw done ?? new Error(`${appName} closed the document but did not write ${targetPath}`);
    }
    const sig = await writtenSince(targetPath, start);
    if (sig && sig === lastSig) {
      if (Date.now() - stableSince >= 2_000) break;
    } else {
      lastSig = sig;
      stableSince = Date.now();
    }
  }

  const reopened = await reopenDocument(appName, targetPath);
  if (autoSave && autoSave !== targetPath) await trashUntitledAutoSave(appName, documentName, autoSave);
  return reopened;
}

/** Open a saved file via LaunchServices and return its document name once the app has it. */
async function reopenDocument(appName: IWorkApp, targetPath: string): Promise<{ name: string; reopened: boolean }> {
  const fallbackName = basename(targetPath).replace(/\.(numbers|pages|key)$/i, "");
  try {
    await execFileAsync("/usr/bin/open", ["-g", "-a", resolveAppName(appName), targetPath], { timeout: 15_000 });
    const name = await runJXA<string | null>(
      `
      const app = Application("${resolveAppName(appName)}");
      const t0 = Date.now();
      while (Date.now() - t0 < 15000) {
        const docs = app.documents();
        for (const d of docs) {
          const f = d.file();
          if (f && f.toString() === params.targetPath) return JSON.stringify(d.name());
        }
        delay(0.5);
      }
      return JSON.stringify(null);
      `,
      { targetPath },
      { timeout: 25_000, label: "creatorStudioSaveAs:reopen" },
    );
    return name ? { name, reopened: true } : { name: fallbackName, reopened: false };
  } catch {
    // The file is saved; failing to reopen it must not turn the save into an error.
    return { name: fallbackName, reopened: false };
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
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
    throw new Error("The document has no file on disk yet, and this PDF export works from the saved file. Save it with a file path first.");
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

    // The preview can hold several PDF attachments (Attachment1.pdf, Attachment2.pdf, ...)
    const { stdout } = await execFileAsync("/usr/bin/find", [tmpDir, "-name", "*.pdf", "-type", "f"]);
    const pdfs = stdout.split("\n").filter(Boolean).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    if (pdfs.length === 0) {
      throw new Error("Quick Look did not generate a PDF preview for this document");
    }
    if (pdfs.length === 1) {
      await copyFile(pdfs[0], targetPath);
      return;
    }
    await runJXA<void>(
      `
      ObjC.import("Quartz");
      const out = $.PDFDocument.alloc.init;
      for (const path of params.pdfs) {
        const part = $.PDFDocument.alloc.initWithURL($.NSURL.fileURLWithPath(path));
        if (!part || part.isNil()) throw new Error("Could not read " + path);
        for (let i = 0; i < Number(part.pageCount); i++) out.insertPageAtIndex(part.pageAtIndex(i), Number(out.pageCount));
      }
      if (!out.writeToFile(params.targetPath)) throw new Error("Could not write PDF to " + params.targetPath);
      `,
      { pdfs, targetPath },
      { timeout: 60_000, label: "creatorStudioExportPDF:merge" },
    );
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
}

async function pdfPageCount(path: string): Promise<number | null> {
  return runJXA<number | null>(
    `
    ObjC.import("Quartz");
    const pdf = $.PDFDocument.alloc.initWithURL($.NSURL.fileURLWithPath(params.path));
    return JSON.stringify(pdf && !pdf.isNil() ? Number(pdf.pageCount) : null);
    `,
    { path },
    { timeout: 20_000, label: "pdfPageCount" },
  );
}

/**
 * Creator Studio PDF export for Pages. Quick Look may render only the first page,
 * so try the native export first and compare the PDF's page count with the document's.
 */
export async function creatorStudioPagesExportPDF(
  documentName: string,
  targetPath: string,
): Promise<{ method: "native" | "Quick Look"; pageCount: number | null; documentPageCount: number | null }> {
  const documentPageCount = await runJXA<number | null>(
    `
    const doc = Application("Pages").documents.byName(params.documentName);
    try { return JSON.stringify(doc.pages.length); } catch (e) { return JSON.stringify(null); }
    `,
    { documentName },
    { timeout: 20_000, label: "creatorStudioPagesExportPDF:pages" },
  );

  const mtimeBefore = existsSync(targetPath) ? (await stat(targetPath)).mtimeMs : -1;
  let method: "native" | "Quick Look" = "native";
  try {
    await runJXA<void>(
      `
      const app = Application("Pages");
      app.export(app.documents.byName(params.documentName), { to: Path(params.targetPath), as: "PDF" });
      `,
      { documentName, targetPath },
      { timeout: 60_000, label: "creatorStudioPagesExportPDF:native" },
    );
  } catch {}
  if (!existsSync(targetPath) || (await stat(targetPath)).mtimeMs <= mtimeBefore) {
    method = "Quick Look";
    await creatorStudioExportPDF("Pages", documentName, targetPath);
  }

  const pageCount = await pdfPageCount(targetPath);
  if (pageCount !== null && documentPageCount !== null && pageCount < documentPageCount) {
    await rm(targetPath, { force: true });
    throw new Error(
      `PDF export is incomplete: ${method} produced ${pageCount} of ${documentPageCount} pages, so the file was not kept. ` +
        "Export the PDF from Pages by hand (File > Export To > PDF).",
    );
  }
  return { method, pageCount, documentPageCount };
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
  {
    // Untitled documents are renamed on auto-save ("X" -> "X${ext}"), so a new
    // "X" can coexist with it; refuse to guess which one is meant
    let __names = [];
    try { __names = ${appExpr}.documents.name(); } catch(__e) {}
    const __matches = __names.filter(function (n) { return n === params.documentName || n === params.documentName + ${JSON.stringify(ext)}; });
    if (__matches.length > 1) {
      throw new Error("Several open documents match the name " + JSON.stringify(params.documentName) + " (" + __matches.join(", ") +
        "). Save or close one of them first so each has its own name.");
    }
  }
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
    // -1719 is also "invalid index"; from System Events it means UI scripting isn't allowed
    const friendly = errorCode === -1719 && stderr.includes("System Events")
      ? "Accessibility permission missing. Open System Settings → Privacy & Security → Accessibility and allow the app running this server (e.g. Claude or your terminal)."
      : errorCode !== undefined ? COMMON_ERRORS[errorCode] : undefined;
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
  const timeout = options?.timeout ?? jxaTimeout();
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
        if (error.killed || error.signal) {
          reject(new OsascriptError(
            `${label} timed out after ${timeout}ms and osascript was killed (${error.signal ?? "SIGTERM"}). ` +
            `Raise IWORK_MCP_TIMEOUT_MS (or IWORK_MCP_SAVE_TIMEOUT_MS for saves) if the app needs longer.` +
            (stderr ? `\n${stderr}` : ""),
            exitCode,
          ));
          return;
        }
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
