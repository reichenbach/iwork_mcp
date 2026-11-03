# CLAUDE.md

## Project Overview
- MCP server for Apple iWork (Numbers, Pages, Keynote) automation via JXA/osascript
- TypeScript, ESM, uses `@modelcontextprotocol/sdk` v1.26.0
- 117 tools total: 50 Numbers, 26 Pages, 41 Keynote (includes 8 Creator Studio AI tools)
- npm: `iwork-mcp`
- Requirements: macOS 13+, iWork 14.0+, Node.js 18+
- Supports both standard iWork and iWork 15.1+ "Creator Studio" app bundles (tested on 15.1.1)

## Key Technical Details
- MCP SDK imports: `@modelcontextprotocol/sdk/server/mcp.js` and `.../server/stdio.js`
- `server.tool(name, description, zodSchema, annotations, callback)` — uses zod raw shapes (not z.object())
- Tool callbacks return `{ content: [{ type: "text", text: "..." }], isError?: boolean }`
- JXA scripts pass params via `JSON.stringify` in argv[0], return via `JSON.stringify` in stdout
- `execFile("/usr/bin/osascript", ["-l", "JavaScript", "-e", script, jsonArgs])`
- App name resolution in `src/jxa.ts`: prefers Creator Studio if installed, falls back to standard names. All tool scripts use standard names — `rewriteAppNames()` rewrites them transparently before execution.
- Numbers cell `format` property: set via JXA `cell.format = "percent"` (not "percentage"). Valid: automatic, number, currency, percent, fraction, scientific, text, checkbox, star rating.

## Critical JXA Bugs (FIXED)
- **Colors**: Numbers uses 0-65535 int range, NOT 0-1 floats. Multiply hex by 257.
- **Font names**: Use PostScript names (`HelveticaNeue-Bold`) not display (`Helvetica Neue Bold`). Display names cause -10000 error.
- **Merging**: Can't merge across header/non-header boundaries. Set headerRowCount/headerColumnCount to 0 first.
- **Charts**: JXA can't bind data to charts. `numbers_add_chart` uses AppleScript bridge via `$.NSAppleScript` to create data-bound charts from selection.
- **Pages 14.5 paragraphs**: `doc.paragraphs` completely broken — no push/read/format/index. Only `doc.bodyText` (plain string) works.
- **Pages JXA paragraph access**: Use `doc.bodyText.paragraphs[i]`, NOT `doc.paragraphs[i]`. Properties are `font`/`size`/`color` (not `fontName`/`fontSize`/`bold`/`italic`). Color writes use 0-65535 ints, reads return 0-1 floats. No `bold`/`italic` properties — use PostScript font names instead.
- **Pages bodyText formatting**: Setting `doc.bodyText = "..."` destroys ALL formatting. For structural changes (add/insert/delete), save formats first, set bodyText, then restore. For replace, use per-paragraph `doc.bodyText.paragraphs[i] = "new text"` which preserves other paragraphs' formatting.
- **Pages paragraph index shift**: When paragraph text contains `\n`, Pages creates multiple actual paragraphs from one input. `pages_create_document_with_content` must track real paragraph index (count `\n` per input) or formatting bleeds across paragraphs.
- **Numbers default table**: New sheets auto-create "Table 1". `numbers_add_sheet` now deletes it by default.
- **Keynote masterSlides JXA bridge broken**: `doc.masterSlides()`, `.byName()`, `.name()` all fail with -1700 "Can't convert types". Use `NSAppleScript` ObjC bridge from within JXA to run AppleScript for master slide operations.
- **Keynote shapeType**: No JXA API to set shape type programmatically. Parameter removed from `keynote_add_shape`.
- **Numbers add_row at beginning**: JXA `table.rows.push()` always appends. To insert at beginning, must shift existing cell data down first, then write new data at row 0.
- **JXA export format strings**: Use `"PDF"` not `"Numbers PDF"`/`"Pages PDF"`/`"Keynote PDF"`. The app-prefixed strings cause -1700 "Can't convert types". Keynote's `imageFormat` export option can't be passed from JXA at all (-1700; its enum shares the code `Knmf` with movie formats) — leave it out, slide images default to PNG.
- **JXA `table.ranges["B2:C3"]`**: Completely broken — always throws "Invalid index" or "Can't get object". `read_range` manually parses cell refs and iterates cells.
- **Numbers minimum table size**: `app.Table({columnCount: 1})` throws "Invalid column count (-10000)". Minimum is 2 columns AND 2 rows.
- **Keynote can't delete all slides**: Can't `app.delete` the only slide — Keynote requires at least 1. Compound tool reuses the auto-generated first slide.
- **Keynote `doc.save({ in: })` unreliable across processes**: `byName` + `save({ in: Path(...) })` fails with -1728 from a different osascript subprocess. Workaround: use `app.documents[0]` or combine create+save in one JXA call.
- **Creator Studio `doc.save({ in: })` hangs**: Use `creatorStudioSaveAs()` in `src/jxa.ts` — closes with auto-save, copies file, reopens from new path. A plain save (no path) works via AppleScript `save document ...` — don't skip it. Untitled documents are auto-saved to iCloud right away, and the copy stays after Save As or closing without saving; both move it to the Trash, but only for documents a tool created.
- **Creator Studio `app.export()` fails with error 6**: Use `creatorStudioExportPDF()` in `src/jxa.ts` — uses `qlmanage` Quick Look to generate PDF. Not for Keynote: Quick Look yields only one (blank) page for `.key`, so `creatorStudioKeynoteExportPDF()` tries the native export, then falls back to exporting slide images and assembling them into a PDF with PDFKit.
- **Creator Studio document name resolution**: Auto-save renames documents by appending file extensions (e.g. "Untitled 1" -> "Untitled 1.numbers"). `injectDocumentNameResolution()` in `src/jxa.ts` handles this transparently by trying the extended name on lookup failure.
- **NSAppleScript failure check**: `executeAndReturnError()` returns a nil wrapper on failure that is still truthy in JXA — `if (!result)` never fires and errors are swallowed. Check `!result || result.isNil()` (Keynote: `runAS()` in `KEYNOTE_AS_HELPERS`). Pass the error out-param as `$()` and read it with `ObjC.deepUnwrap(err)` — reading `Ref()[0]` crashes osascript (SIGBUS/SIGSEGV).
- **Keynote slide insertion/move**: JXA `slides.splice()` fails (-10002), so use AppleScript insertion locations: `make new slide at after slide N`, `duplicate slide N to after slide M`. `move slide N to slide M` (and JXA `app.move(slide, {to: slides[M]})`) REPLACES slide M — use `before slide M` when moving backward, `after slide M` when moving forward.
- **Keynote shape position**: Setting `shape.position = [x, y]` (array format) silently fails — reads back as (0,0). Must use object format: `shape.position = {x: x, y: y}`. Width/height setting works fine with direct assignment.
- **JXA `indexOf()` on scriptable objects**: `doc.sheets().indexOf(sheet)` returns -1 because JXA object references can't be compared with `===`. Iterate by name instead.
- **Pages tables inaccessible via JXA**: `doc.tables()` throws -2763 "Don't know how to create TMAScriptTableInfoProxy" — but the AppleScript dictionary exposes them fully. The four pages table tools (list/read/write_cells/resize) use the NSAppleScript bridge, same as Keynote master-slide ops. Cell writes with strings starting "=" set formulas; `set row count` grows/shrinks tables. CAVEAT: table CREATION (`make new table`) is broken on Creator Studio 15.x in BOTH JXA and AppleScript (-2763); on 15.x `pages_add_table` falls back to the UI-scripting workaround (see the "Pages 15.x table creation broken" bullet).
- **Keynote text colors**: Use 0-65535 int range (like Numbers), for shape text AND table cell `textColor`: `[r*257, g*257, b*257]`. 0-1 floats are accepted silently and render near-black. Reads may come back in either range — normalize by the max component.
- **Keynote hidden placeholders**: Layouts like Blank/Quote keep a title/body placeholder with `titleShowing`/`bodyShowing` = false. Text written there reads back fine but is invisible — set `slide.titleShowing = true` first.
- **Keynote text alignment**: Not settable via scripting (`objectText.alignment` always errors). Don't expose it.
- **Keynote transitions**: `transitionEffect` takes the dictionary enumerator name — "no transition effect", not "none". Invalid names fail with an opaque -1700; valid names come from `sdef` (the enumeration containing "magic move").
- **Keynote stop slideshow**: JXA `app.stop(doc)` can return without stopping. Use AppleScript `stop document ...` and verify with `app.playing()`.
- **Numbers auto-parses formatted strings**: Writing `"$1,234.56"` to a cell auto-converts to numeric 1234.56. To preserve strings, set `cell.format = "text"` before writing.
- **Pages 15.x table creation broken**: `doc.tables.push()` AND AppleScript `make new table` both fail with -2763 "TMAScriptTableInfoProxy" (reading/resizing existing tables via AppleScript works). `pages_add_table` falls back to `pagesInsertTableViaUI()` in `src/jxa.ts` on Creator Studio: System Events menu-clicks Insert > Table > Basic, then resizes via AppleScript `set row count`/`set column count`. The insertion point is established by clicking into the body with a top-down position scan (a click on a table/image selects it and disables the menu). Synthetic keystrokes are deliberately avoided — secure keyboard entry from ANY app (e.g. a browser password field) silently blocks them, but not mouse/menu clicks. Menu clicks must be two-step (open parent menu, then click child) — direct 3-level submenu clicks silently no-op.

## File Structure
- `src/index.ts` — entry point + install routing
- `src/install.ts` — auto-config for Claude Desktop
- `src/jxa.ts` — runJXA(), OsascriptError, Creator Studio workarounds
- `src/tools/{numbers,pages,keynote}.ts` — tool definitions
- `src/instructions.ts` — MCP instructions string (design guidance for AI clients)
- `scripts/create-screenshots.ts` — showcase document builder (budget, resume, pitch)
- `scripts/create-examples.ts` — example file generator (sample.numbers, sample.pages, sample.key)
- `.github/workflows/ci.yml` — CI: build + unit tests on push/PR, Node 18 & 22 matrix
- `test/helpers/{server,app-check}.ts` — test helpers (InMemoryTransport, app detection)
- `test/{registration,jxa}.test.ts` — fast unit tests (no apps needed)
- `test/{numbers,pages,keynote}.test.ts` — integration tests (require apps)

## Testing
- `npm test` — unit tests (registration + jxa), ~300ms, no apps needed, runs in CI
- `npm run test:integration` — CRUD tests for Numbers/Pages/Keynote, ~8s, needs apps
- `npm run test:all` — both tiers combined
- Uses `node:test` + `tsx`, in-memory MCP transport (no subprocess)
- 102 tests total: 34 unit + 68 integration

## Build & Publish
- `npm run build` -> `tsc && chmod +x dist/index.js`
- `npm publish` requires 2FA (OTP). User publishes manually.
- Always bump version before publish (npm rejects duplicates)
- **Keep version in sync**: bump -> build -> test -> commit -> push -> publish -> GitHub release
- After each npm publish, create the matching GitHub release server-side: `gh release create vX.Y.Z --generate-notes`. GitHub Releases do NOT update automatically on publish — they lag behind npm unless this step runs (v0.8.6 sat as "Latest" long after 0.8.7 shipped).
- **NEVER `git push --tags`** — stale local tags may pin old, rewritten history. Always create tags server-side via `gh release create`.
- Follow the repository's configured commit identity. Never include personal addresses in public files.
- Screenshots must be JPEG not PNG — large PNGs cause 503 errors on GitHub README rendering

## Known Issues
- All 6 Pages text tools (pages_add_text, pages_get_paragraphs, pages_format_text, pages_insert_text_at, pages_delete_text, pages_replace_text) work via `doc.bodyText.paragraphs` workarounds. The original `doc.paragraphs` API remains broken on Pages 14.5.
- Pages tables are not accessible via JXA (-2763) but ARE via AppleScript: pages_list_tables, pages_read_table, pages_write_table_cells, and pages_resize_table use the NSAppleScript bridge. On Creator Studio 15.x, table creation is broken in both JXA and AppleScript — pages_add_table falls back to a UI-scripting workaround there (see Critical JXA Bugs) which requires Accessibility permission for the host app.
- Keynote shape fill/border colors are not exposed by JXA. Only opacity, rotation, and text formatting are settable via `keynote_format_shape`.
- Pages paragraph styles (Title, Heading 1, Body, etc.) are NOT exposed by Apple's scripting dictionary. Only font, size, and color are accessible. No alignment, indent, or line spacing either.
