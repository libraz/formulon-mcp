import assert from "node:assert/strict";
import test from "node:test";
import {
  cellToA1,
  colToIndex,
  indexToCol,
  normalizeRange,
  parseCellRef,
  parseRangeRef,
} from "../dist/a1.js";

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
