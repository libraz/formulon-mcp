import assert from "node:assert/strict";
import { test } from "vitest";
import {
  cellToA1,
  colToIndex,
  indexToCol,
  MAX_COL,
  MAX_ROW,
  normalizeRange,
  parseCellRef,
  parseRangeRef,
  parseSelectionRef,
} from "../dist/a1.js";
import * as sessionShared from "../dist/session/shared.js";

test("converts between Excel columns and zero-based indexes", () => {
  assert.equal(colToIndex("A"), 0);
  assert.equal(colToIndex("Z"), 25);
  assert.equal(colToIndex("AA"), 26);
  assert.equal(indexToCol(0), "A");
  assert.equal(indexToCol(25), "Z");
  assert.equal(indexToCol(26), "AA");
});

test("parses and formats A1 cell references", () => {
  assert.deepEqual(parseCellRef("Sheet1!B2"), {
    sheetName: "Sheet1",
    row: 1,
    col: 1,
  });
  assert.deepEqual(parseCellRef("'Sales Q1'!$AA$10"), {
    sheetName: "Sales Q1",
    row: 9,
    col: 26,
  });
  assert.equal(cellToA1(9, 26), "AA10");
});

test("parses and normalizes A1 ranges", () => {
  const range = normalizeRange(parseRangeRef("Sheet1!C3:A1"));
  assert.equal(range.sheetName, "Sheet1");
  assert.deepEqual(range.start, { sheetName: "Sheet1", row: 0, col: 0 });
  assert.deepEqual(range.end, { sheetName: "Sheet1", row: 2, col: 2 });

  const unqualified = normalizeRange(parseRangeRef("A1:B2"));
  assert.equal(unqualified.sheetName, undefined);
  assert.deepEqual(unqualified.start, { sheetName: undefined, row: 0, col: 0 });
  assert.deepEqual(unqualified.end, { sheetName: undefined, row: 1, col: 1 });
});

test("rejects invalid cell references", () => {
  assert.throws(() => parseCellRef("1A"), /invalid A1 cell reference/);
  assert.throws(() => parseRangeRef("A1:B2:C3"), /invalid A1 range reference/);
  assert.throws(() => colToIndex("A1"), /invalid column/);
  assert.throws(() => colToIndex(""), /invalid column/);
  assert.throws(() => cellToA1(-1, 0), /invalid zero-based row index/);
});

test("exports the grid limits from one place", () => {
  assert.equal(MAX_ROW, 1_048_575);
  assert.equal(MAX_COL, 16_383);
  assert.equal(indexToCol(MAX_COL), "XFD");
  assert.equal(sessionShared.MAX_ROW, MAX_ROW);
  assert.equal(sessionShared.MAX_COL, MAX_COL);
});

const span = (sheetName, startRow, startCol, endRow, endCol) => ({
  sheetName,
  start: { sheetName, row: startRow, col: startCol },
  end: { sheetName, row: endRow, col: endCol },
});

const SELECTIONS = [
  // Whole columns span every row.
  ["B:D", span(undefined, 0, 1, MAX_ROW, 3)],
  ["$B:$D", span(undefined, 0, 1, MAX_ROW, 3)],
  ["$B:D", span(undefined, 0, 1, MAX_ROW, 3)],
  ["b:d", span(undefined, 0, 1, MAX_ROW, 3)],
  ["C:C", span(undefined, 0, 2, MAX_ROW, 2)],
  ["A:XFD", span(undefined, 0, 0, MAX_ROW, MAX_COL)],
  // Whole rows span every column.
  ["3:5", span(undefined, 2, 0, 4, MAX_COL)],
  ["$3:$5", span(undefined, 2, 0, 4, MAX_COL)],
  ["7:7", span(undefined, 6, 0, 6, MAX_COL)],
  ["1:1048576", span(undefined, 0, 0, MAX_ROW, MAX_COL)],
  // Reversed spans parse as written; normalizeRange orders them.
  ["D:B", span(undefined, 0, 3, MAX_ROW, 1)],
  ["5:3", span(undefined, 4, 0, 2, MAX_COL)],
  // A sheet prefix, quoted or not, applies to both ends.
  ["Sheet1!B:D", span("Sheet1", 0, 1, MAX_ROW, 3)],
  ["Sheet1!$3:$5", span("Sheet1", 2, 0, 4, MAX_COL)],
  ["'Sales Q1'!B:B", span("Sales Q1", 0, 1, MAX_ROW, 1)],
  ["'My ''Q'' Sheet'!1:2", span("My 'Q' Sheet", 0, 0, 1, MAX_COL)],
  ["'My ''Q'' Sheet'!C:D", span("My 'Q' Sheet", 0, 2, MAX_ROW, 3)],
  // Surrounding whitespace is trimmed.
  ["  B:D  ", span(undefined, 0, 1, MAX_ROW, 3)],
  // Everything else is an ordinary cell or range.
  ["A1", span(undefined, 0, 0, 0, 0)],
  ["A1:B2", span(undefined, 0, 0, 1, 1)],
  ["$A$1:$B$2", span(undefined, 0, 0, 1, 1)],
  ["Sheet1!A1:B2", span("Sheet1", 0, 0, 1, 1)],
  ["'Sales Q1'!C3", span("Sales Q1", 2, 2, 2, 2)],
  ["XFD1048576", span(undefined, MAX_ROW, MAX_COL, MAX_ROW, MAX_COL)],
];

for (const [ref, expected] of SELECTIONS) {
  test(`parses the selection ${JSON.stringify(ref)}`, () => {
    assert.deepEqual(parseSelectionRef(ref), expected);
  });
}

test("normalizes reversed whole-line selections", () => {
  assert.deepEqual(normalizeRange(parseSelectionRef("D:B")), span(undefined, 0, 1, MAX_ROW, 3));
  assert.deepEqual(normalizeRange(parseSelectionRef("5:3")), span(undefined, 2, 0, 4, MAX_COL));
  assert.deepEqual(
    normalizeRange(parseSelectionRef("Sheet1!D:B")),
    span("Sheet1", 0, 1, MAX_ROW, 3),
  );
});

test("keeps whole-line shapes out of the plain range parser", () => {
  assert.throws(() => parseRangeRef("B:D"), /invalid A1 cell reference/);
  assert.throws(() => parseRangeRef("3:5"), /invalid A1 cell reference/);
});

for (const ref of [
  "0:1",
  "1:0",
  "A:1",
  "1:A",
  ":",
  "A:",
  ":A",
  ":1",
  "",
  "  ",
  "A:B:C",
  "1:2:3",
  "B$:D",
  "Sheet1!",
  "Sheet1!:",
]) {
  test(`rejects the selection ${JSON.stringify(ref)}`, () => {
    assert.throws(() => parseSelectionRef(ref), /invalid A1/);
  });
}
