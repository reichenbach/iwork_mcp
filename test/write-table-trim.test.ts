import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { WRITE_TABLE_TRIM_JS, DEFAULT_NEW_TABLE_SIZE } from "../src/tools/numbers.js";
import { BODY_TEXT_POSITION_JS } from "../src/tools/pages.js";

// The helpers are JavaScript source that runs inside osascript; evaluate the same source here.
type Cell = string | number | boolean | null | undefined;
const { shouldTrimTable, trimmedSize } = new Function(
  `${WRITE_TABLE_TRIM_JS}; return { shouldTrimTable, trimmedSize };`,
)() as {
  shouldTrimTable: (trimToFit: boolean | null | undefined, rows: number, cols: number, values: Cell[]) => boolean;
  trimmedSize: (values: Cell[], rows: number, cols: number, needRows: number, needCols: number) => { rowCount: number; columnCount: number };
};
const { paraBounds, findCharPositions } = new Function(
  `${BODY_TEXT_POSITION_JS}; return { paraBounds, findCharPositions };`,
)() as {
  paraBounds: (text: string | string[]) => Array<{ start: number; end: number; sep: string }>;
  findCharPositions: (text: string, find: string, all: boolean) => number[];
};

/** Row-major values of a rows x cols table, with the given cells filled. */
function grid(rows: number, cols: number, filled: Record<string, Cell> = {}): Cell[] {
  const v: Cell[] = new Array(rows * cols).fill(null);
  for (const [k, val] of Object.entries(filled)) {
    const [r, c] = k.split(",").map(Number);
    v[r * cols + c] = val;
  }
  return v;
}

/** Simulates numbers_write_table on a table: grow, write data at (0,0), decide + trim. */
function simulateWrite(rows: number, cols: number, before: Cell[], data: Cell[][], trimToFit?: boolean | null) {
  const trim = shouldTrimTable(trimToFit ?? null, rows, cols, before);
  const needRows = data.length, needCols = Math.max(...data.map((r) => r.length));
  const R = Math.max(rows, needRows), C = Math.max(cols, needCols);
  const after: Cell[] = new Array(R * C).fill(null);
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) after[r * C + c] = before[r * cols + c];
  data.forEach((row, r) => row.forEach((val, c) => { if (val !== null) after[r * C + c] = val; }));
  return trim ? trimmedSize(after, R, C, needRows, needCols) : { rowCount: R, columnCount: C };
}

describe("numbers_write_table trimming", () => {
  const { rowCount: DR, columnCount: DC } = DEFAULT_NEW_TABLE_SIZE;

  it("never shrinks an existing table by default (10x4 with blank rows, 2x3 write)", () => {
    const before = grid(10, 4, { "0,0": "Name", "0,1": "Qty", "1,0": "a", "3,2": 7 });
    assert.equal(shouldTrimTable(null, 10, 4, before), false);
    assert.equal(shouldTrimTable(undefined, 10, 4, before), false);
    assert.deepEqual(simulateWrite(10, 4, before, [["x", "y", "z"], [1, 2, 3]]), { rowCount: 10, columnCount: 4 });
  });

  it("does not shrink a completely empty table that is not the default size", () => {
    assert.equal(shouldTrimTable(null, 10, 4, grid(10, 4)), false);
    assert.equal(shouldTrimTable(null, DR, DC + 1, grid(DR, DC + 1)), false);
    assert.equal(shouldTrimTable(null, DR - 1, DC, grid(DR - 1, DC)), false);
  });

  it("does not shrink a 22x7 table that already holds data", () => {
    assert.equal(shouldTrimTable(null, DR, DC, grid(DR, DC, { "15,5": "keep" })), false);
    assert.equal(shouldTrimTable(null, DR, DC, grid(DR, DC, { "0,0": 0 })), false, "0 is a value");
    assert.equal(shouldTrimTable(null, DR, DC, grid(DR, DC, { "0,0": false })), false, "false is a value");
  });

  it("trims the untouched empty 22x7 default table of a new document (5x2 write -> 5x2)", () => {
    assert.equal(shouldTrimTable(null, DR, DC, grid(DR, DC)), true);
    const data = [["h1", "h2"], [1, 2], [3, 4], [5, 6], [7, 8]];
    assert.deepEqual(simulateWrite(DR, DC, grid(DR, DC), data), { rowCount: 5, columnCount: 2 });
  });

  it("trimToFit: false never trims, even the default table", () => {
    assert.equal(shouldTrimTable(false, DR, DC, grid(DR, DC)), false);
    assert.deepEqual(simulateWrite(DR, DC, grid(DR, DC), [["a", "b"], [1, 2]], false), { rowCount: DR, columnCount: DC });
  });

  it("trimToFit: true trims any table, but only empty trailing rows/columns", () => {
    const before = grid(10, 4, { "0,0": "Name", "1,0": "a" });
    assert.equal(shouldTrimTable(true, 10, 4, before), true);
    assert.deepEqual(simulateWrite(10, 4, before, [["x", "y", "z"], [1, 2, 3]], true), { rowCount: 2, columnCount: 3 });
    // A value in row 7 / column D survives the trim
    const kept = grid(10, 4, { "6,3": "note" });
    assert.deepEqual(simulateWrite(10, 4, kept, [["x", "y"], [1, 2]], true), { rowCount: 7, columnCount: 4 });
  });

  it("trimmedSize never goes below the data or Numbers' 2x2 minimum", () => {
    assert.deepEqual(trimmedSize(grid(10, 5), 10, 5, 1, 1), { rowCount: 2, columnCount: 2 });
    assert.deepEqual(trimmedSize(grid(10, 5), 10, 5, 4, 3), { rowCount: 4, columnCount: 3 });
    // Empty string counts as empty; a blank row in the middle is kept
    const v = grid(6, 3, { "0,0": "a", "3,1": "b", "4,2": "" });
    assert.deepEqual(trimmedSize(v, 6, 3, 1, 1), { rowCount: 4, columnCount: 2 });
  });

  it("growing still works without trimming", () => {
    assert.deepEqual(simulateWrite(2, 2, grid(2, 2, { "0,0": "a" }), [[1, 2, 3], [4, 5, 6], [7, 8, 9]]), { rowCount: 3, columnCount: 3 });
  });
});

describe("Pages body text positions", () => {
  it("paraBounds splits on \\n and \\f (page break)", () => {
    assert.deepEqual(paraBounds("ab\ncd\fef"), [
      { start: 0, end: 2, sep: "\n" },
      { start: 3, end: 5, sep: "\f" },
      { start: 6, end: 8, sep: "" },
    ]);
  });

  it("paraBounds handles empty text, empty paragraphs and character arrays", () => {
    assert.deepEqual(paraBounds(""), [{ start: 0, end: 0, sep: "" }]);
    assert.deepEqual(paraBounds("a\n\nb\n"), [
      { start: 0, end: 1, sep: "\n" },
      { start: 2, end: 2, sep: "\n" },
      { start: 3, end: 4, sep: "\n" },
      { start: 5, end: 5, sep: "" },
    ]);
    // Array.from(body): an emoji is one entry, so bounds are in code points
    assert.deepEqual(paraBounds(Array.from("😀x\ny")), [
      { start: 0, end: 2, sep: "\n" },
      { start: 3, end: 4, sep: "" },
    ]);
  });

  it("findCharPositions returns code-point positions of all matches", () => {
    assert.deepEqual(findCharPositions("cat dog cat", "cat", true), [0, 8]);
    assert.deepEqual(findCharPositions("😀 cat 😀 cat", "cat", true), [2, 8]);
    assert.deepEqual(findCharPositions("x￼cat", "cat", true), [2]);
  });

  it("findCharPositions stops after the first match when all is false", () => {
    assert.deepEqual(findCharPositions("cat dog cat", "cat", false), [0]);
  });

  it("findCharPositions does not overlap matches and handles no match", () => {
    assert.deepEqual(findCharPositions("aaaa", "aa", true), [0, 2]);
    assert.deepEqual(findCharPositions("abc", "z", true), []);
    assert.deepEqual(findCharPositions("😀😀", "😀", true), [0, 1]);
  });
});
