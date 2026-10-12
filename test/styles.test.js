import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { cellToA1, MAX_COL, MAX_ROW } from "../dist/a1.js";
import {
  applySessionLayout,
  callWorkbookMethod,
  closeSession,
  listSessions,
  openSession,
  saveSession,
  sessionDefaultFont,
  setSessionRange,
  styleSessionRange,
} from "../dist/sessions.js";
import { parseColor, STYLE_VOCABULARY } from "../dist/styles.js";

/** Reads the resolved xf record a cell carries. */
function cellXf(id, sheet, row, col) {
  const { result } = callWorkbookMethod(id, "getCellXfIndex", [sheet, row, col]);
  return callWorkbookMethod(id, "getCellXf", [result.xfIndex]).result;
}

/**
 * Reads the format code applied to a cell. The id is asserted through the code
 * it resolves to, since `addNumFmt` reuses a built-in slot when one already
 * spells the requested format.
 */
function cellFormatCode(id, sheet, row, col) {
  const numFmtId = cellXf(id, sheet, row, col).numFmtId;
  return callWorkbookMethod(id, "getNumFmt", [numFmtId]).result.formatCode;
}

test("parses hex colors into AARRGGBB", () => {
  assert.equal(parseColor("#FF0000"), 0xffff0000);
  assert.equal(parseColor("00FF00"), 0xff00ff00);
  assert.equal(parseColor("#8000FF00"), 0x8000ff00);
  assert.throws(() => parseColor("#GGG"), /invalid color/);
});

test("names every ordinal table it maps", () => {
  // The ordinals are positional, so a table that drifts out of OOXML's order
  // would apply the wrong style silently.
  assert.equal(STYLE_VOCABULARY.borderStyle[0], "none");
  assert.equal(STYLE_VOCABULARY.borderStyle[1], "thin");
  assert.equal(STYLE_VOCABULARY.fillPattern[1], "solid");
  assert.equal(STYLE_VOCABULARY.horizontalAlign[2], "center");
  assert.equal(STYLE_VOCABULARY.verticalAlign[1], "center");
});

test("applies a font, fill, and number format across a range", async () => {
  await openSession(undefined, "style-basic");
  try {
    setSessionRange("style-basic", "A1", [["Item", "Amount"]], undefined, false);
    const applied = styleSessionRange(
      "style-basic",
      "A1:B1",
      {
        font: { bold: true, size: 14, color: "#1F4E79" },
        fill: { color: "#DCE6F1" },
        align: { horizontal: "center" },
      },
      undefined,
      "existing",
    );
    assert.equal(applied.regions.length, 1);
    assert.equal(applied.regions[0].range, "A1:B1");

    const xf = cellXf("style-basic", 0, 0, 1);
    assert.equal(xf.horizontalAlign, STYLE_VOCABULARY.horizontalAlign.indexOf("center"));
    assert.equal(xf.hasAlignment, true);

    const font = callWorkbookMethod("style-basic", "getFont", [xf.fontIndex]).result;
    assert.equal(font.bold, true);
    assert.equal(font.size, 14);
    assert.equal(font.colorArgb, parseColor("#1F4E79"));

    const fill = callWorkbookMethod("style-basic", "getFill", [xf.fillIndex]).result;
    assert.equal(fill.pattern, STYLE_VOCABULARY.fillPattern.indexOf("solid"));
    assert.equal(fill.fgArgb, parseColor("#DCE6F1"));
  } finally {
    closeSession("style-basic");
  }
});

test("keeps style a later pass did not state", async () => {
  await openSession(undefined, "style-delta");
  try {
    setSessionRange("style-delta", "A1", [[1000], [2000]], undefined, false);
    styleSessionRange("style-delta", "A1:A2", { numberFormat: '"¥"#,##0' }, undefined, "existing");
    assert.equal(cellFormatCode("style-delta", 0, 1, 0), '"¥"#,##0');

    // Bolding the total must not undo the currency format under it, which is
    // how Excel's own formatting behaves.
    styleSessionRange("style-delta", "A2", { font: { bold: true } }, undefined, "existing");
    const after = cellXf("style-delta", 0, 1, 0);
    assert.equal(cellFormatCode("style-delta", 0, 1, 0), '"¥"#,##0');
    assert.equal(callWorkbookMethod("style-delta", "getFont", [after.fontIndex]).result.bold, true);
  } finally {
    closeSession("style-delta");
  }
});

test("splits a mixed range by the style each cell already carries", async () => {
  await openSession(undefined, "style-mixed");
  try {
    setSessionRange("style-mixed", "A1", [["Label", 1000]], undefined, false);
    styleSessionRange("style-mixed", "B1", { numberFormat: '"¥"#,##0' }, undefined, "existing");

    const applied = styleSessionRange(
      "style-mixed",
      "A1:B1",
      { font: { italic: true } },
      undefined,
      "existing",
    );
    assert.equal(applied.regions.length, 2);
    assert.equal(new Set(applied.regions.map((region) => region.baseXfIndex)).size, 2);
    // The formatted cell keeps its format; the plain one is left General.
    assert.equal(cellXf("style-mixed", 0, 0, 0).numFmtId, 0);
    assert.equal(cellFormatCode("style-mixed", 0, 0, 1), '"¥"#,##0');
  } finally {
    closeSession("style-mixed");
  }
});

test("draws an outline around a range without ruling its interior", async () => {
  await openSession(undefined, "style-outline");
  try {
    const applied = styleSessionRange(
      "style-outline",
      "B2:D4",
      { border: { all: "thin", outline: "medium" } },
      undefined,
      "default",
    );
    // Three row bands by three column bands: corners, edges, and interior.
    assert.equal(applied.regions.length, 9);

    const thin = STYLE_VOCABULARY.borderStyle.indexOf("thin");
    const medium = STYLE_VOCABULARY.borderStyle.indexOf("medium");
    const borderOf = (row, col) =>
      callWorkbookMethod("style-outline", "getBorder", [
        cellXf("style-outline", 0, row, col).borderIndex,
      ]).result;

    const topLeft = borderOf(1, 1);
    assert.equal(topLeft.top.style, medium);
    assert.equal(topLeft.left.style, medium);
    assert.equal(topLeft.right.style, thin);
    assert.equal(topLeft.bottom.style, thin);

    const middle = borderOf(2, 2);
    assert.equal(middle.top.style, thin);
    assert.equal(middle.left.style, thin);
    assert.equal(middle.right.style, thin);
    assert.equal(middle.bottom.style, thin);

    const bottomRight = borderOf(3, 3);
    assert.equal(bottomRight.bottom.style, medium);
    assert.equal(bottomRight.right.style, medium);
  } finally {
    closeSession("style-outline");
  }
});

test("styles a single cell and materializes it as a styled blank", async () => {
  await openSession(undefined, "style-blank");
  try {
    const applied = styleSessionRange(
      "style-blank",
      "C3",
      { border: { all: "thin" } },
      undefined,
      "default",
    );
    assert.equal(applied.regions.length, 1);
    assert.equal(applied.regions[0].range, "C3:C3");
    assert.ok(cellXf("style-blank", 0, 2, 2).borderIndex > 0);
  } finally {
    closeSession("style-blank");
  }
});

test("names the accepted values when a style word is unknown", async () => {
  await openSession(undefined, "style-vocab");
  try {
    assert.throws(
      () =>
        styleSessionRange(
          "style-vocab",
          "A1",
          { border: { all: "hairline" } },
          undefined,
          "default",
        ),
      /unknown border style: hairline .*slantDashDot/,
    );
    assert.throws(
      () =>
        styleSessionRange(
          "style-vocab",
          "A1",
          { align: { horizontal: "middle" } },
          undefined,
          "default",
        ),
      /unknown horizontal alignment: middle/,
    );
  } finally {
    closeSession("style-vocab");
  }
});

test("redeclares the workbook default font in place", async () => {
  await openSession(undefined, "default-font");
  try {
    const seeded = sessionDefaultFont("default-font");
    assert.equal(seeded.font.name, "Calibri");
    assert.equal(seeded.font.size, 11);
    const seededCount = callWorkbookMethod("default-font", "fontCount", []).result.value;

    const applied = sessionDefaultFont("default-font", { name: "Meiryo", size: 10 });
    assert.equal(applied.font.name, "Meiryo");
    assert.equal(applied.font.size, 10);
    assert.equal(sessionDefaultFont("default-font").font.name, "Meiryo");
    // The default occupies font 0, so redeclaring it overwrites the slot rather
    // than appending a font every unstyled cell would still not resolve to.
    assert.equal(callWorkbookMethod("default-font", "fontCount", []).result.value, seededCount);
  } finally {
    closeSession("default-font");
  }
});

test("cuts a font's theme link when a typeface is named", async () => {
  await openSession(undefined, "font-scheme");
  try {
    // A ja-JP workbook's base font is theme-linked, which makes Excel resolve
    // the typeface from the theme and discard whatever name was written.
    callWorkbookMethod("font-scheme", "setDefaultFont", [
      { ...callWorkbookMethod("font-scheme", "getFont", [0]).result, scheme: 2 },
    ]);
    assert.equal(sessionDefaultFont("font-scheme").font.themeLink, "minor");

    // Size alone leaves the link alone: only the typeface comes from the theme.
    assert.equal(sessionDefaultFont("font-scheme", { size: 12 }).font.themeLink, "minor");
    assert.equal(sessionDefaultFont("font-scheme", { name: "Meiryo" }).font.themeLink, "none");

    setSessionRange("font-scheme", "A1", [["見出し"]], undefined, false);
    styleSessionRange("font-scheme", "A1", { font: { bold: true } }, undefined, "existing");
    const bolded = callWorkbookMethod("font-scheme", "getFont", [
      cellXf("font-scheme", 0, 0, 0).fontIndex,
    ]).result;
    assert.equal(bolded.name, "Meiryo");
    assert.equal(bolded.scheme, 0);
  } finally {
    closeSession("font-scheme");
  }
});

test("refuses a range too wide to materialize", async () => {
  await openSession(undefined, "style-huge");
  try {
    assert.throws(
      () =>
        styleSessionRange(
          "style-huge",
          "A1:Z100000",
          { font: { bold: true } },
          undefined,
          "default",
        ),
      /style a smaller range/,
    );
  } finally {
    closeSession("style-huge");
  }
});

test("styles whole columns and rows as columns and rows", async () => {
  await openSession(undefined, "style-lines");
  try {
    setSessionRange("style-lines", "B2", [[5]], undefined, false);
    const applied = styleSessionRange(
      "style-lines",
      "B:B",
      { fill: { color: "#FFFF00" } },
      undefined,
      "existing",
    );
    assert.equal(applied.range.end, "B1048576");
    // No blank cell of the column is materialized; the column carries the style.
    assert.equal(callWorkbookMethod("style-lines", "cellCount", [0]).result.value, 1);
    const columns = callWorkbookMethod("style-lines", "getSheetColumns", [0]).result.columns;
    assert.equal(columns[0].first, 1);
    assert.equal(columns[0].hasStyle, 1);
    const fill = (row) =>
      callWorkbookMethod("style-lines", "getFill", [cellXf("style-lines", 0, row, 1).fillIndex])
        .result.fgArgb;
    assert.equal(fill(1), parseColor("#FFFF00"));
    // A value written later into the column takes the column's style.
    setSessionRange("style-lines", "B500", [[1]], undefined, false);
    assert.equal(fill(499), parseColor("#FFFF00"));

    styleSessionRange("style-lines", "3:3", { font: { bold: true } }, undefined, "existing");
    const rows = callWorkbookMethod("style-lines", "getSheetRowOverrides", [0]).result.rows;
    assert.equal(
      rows.some((row) => row.row === 2 && row.hasStyle === 1),
      true,
    );

    assert.throws(
      () =>
        styleSessionRange(
          "style-lines",
          "C:D",
          { border: { outline: "thin" } },
          undefined,
          "existing",
        ),
      /bounded range/,
    );
  } finally {
    closeSession("style-lines");
  }
});

// ---------------------------------------------------------------------------
// Range style application across both entry points.
// ---------------------------------------------------------------------------

const ORANGE = parseColor("#FFC000");
const MEDIUM = STYLE_VOCABULARY.borderStyle.indexOf("medium");
const MAX_STYLED_CELLS = 100_000;

/** Calls an engine method on a session and returns the unwrapped result. */
function engine(id, method, ...args) {
  return callWorkbookMethod(id, method, args).result;
}

function sessionDirty(id) {
  return listSessions().find((session) => session.id === id)?.dirty;
}

/** Builds a cell xf that differs from the default xf only by a number format. */
function numberFormatXf(id, code) {
  const numFmtId = engine(id, "addNumFmt", code).numFmtId;
  const record = { ...engine(id, "getCellXf", 0), status: undefined };
  return engine(id, "addXf", { ...record, numFmtId, applyNumberFormat: true }).index;
}

const effectiveAt = (id, row, col, sheet = 0) => engine(id, "getEffectiveStyle", sheet, row, col);

const wholeColumns = (rect) => rect.firstRow === 0 && rect.lastRow === MAX_ROW;
const wholeRows = (rect) => rect.firstCol === 0 && rect.lastCol === MAX_COL;
const isWholeLines = (rect) => wholeColumns(rect) || wholeRows(rect);

/** Target shapes, all containing C3 so a single probe cell exists in every one. */
const SHAPES = {
  cell: { ref: "C3", rect: { firstRow: 2, lastRow: 2, firstCol: 2, lastCol: 2 } },
  rect: { ref: "B2:D4", rect: { firstRow: 1, lastRow: 3, firstCol: 1, lastCol: 3 } },
  col: { ref: "C:C", rect: { firstRow: 0, lastRow: MAX_ROW, firstCol: 2, lastCol: 2 } },
  cols: { ref: "B:D", rect: { firstRow: 0, lastRow: MAX_ROW, firstCol: 1, lastCol: 3 } },
  row: { ref: "3:3", rect: { firstRow: 2, lastRow: 2, firstCol: 0, lastCol: MAX_COL } },
  rows: { ref: "2:4", rect: { firstRow: 1, lastRow: 3, firstCol: 0, lastCol: MAX_COL } },
  // The qualifier names the sheet the range lives on, whatever the fallback is.
  sheetcol: {
    ref: "Sheet1!C:C",
    fallbackSheet: 1,
    rect: { firstRow: 0, lastRow: MAX_ROW, firstCol: 2, lastCol: 2 },
  },
};

/** The two public routes onto `applyRangeStyle`. */
const ENTRY_POINTS = {
  styleSessionRange: (id, shape, style, base) =>
    styleSessionRange(id, shape.ref, style, shape.fallbackSheet, base),
  applyLayoutStyle: (id, shape, style, base) =>
    applySessionLayout(id, 0, [{ type: "style", range: shape.ref, style, base }]).operations[0],
};

/** Prior state the target carries before the style pass; built straight on the engine. */
const PRIORS = {
  blank: () => {},
  numfmt: (id) => {
    const xf = numberFormatXf(id, "0.0%");
    for (const [row, col] of [
      [2, 2],
      [4, 4],
    ]) {
      engine(id, "setNumber", 0, row, col, 1000);
      engine(id, "setCellXfIndex", 0, row, col, xf);
    }
  },
  colstyled: (id) => engine(id, "setColumnStyle", 0, 2, 2, numberFormatXf(id, "0.00%")),
  rowstyled: (id) => engine(id, "setRowStyle", 0, 2, 2, numberFormatXf(id, "0.00%")),
};

const GRID_ROWS = 9;
const GRID_COLS = 7;

/** Everything a style pass may or may not touch, for before/after comparison. */
function snapshot(id) {
  const grid = [];
  for (let row = 0; row < GRID_ROWS; row += 1) {
    for (let col = 0; col < GRID_COLS; col += 1) {
      grid.push({ row, col, style: effectiveAt(id, row, col) });
    }
  }
  return {
    cellCount: engine(id, "cellCount", 0).value,
    columns: engine(id, "getSheetColumns", 0).columns,
    rows: engine(id, "getSheetRowOverrides", 0).rows,
    xfCount: engine(id, "xfCount").value,
    fontCount: engine(id, "fontCount").value,
    fillCount: engine(id, "fillCount").value,
    borderCount: engine(id, "borderCount").value,
    otherSheetCells: engine(id, "cellCount", 1).value,
    otherSheetStyle: effectiveAt(id, 2, 2, 1),
    grid,
  };
}

function bandsOf(first, last) {
  if (first === last) {
    return [{ first, last, atStart: true, atEnd: true }];
  }
  const edges = [
    { first, last: first, atStart: true, atEnd: false },
    { first: last, last, atStart: false, atEnd: true },
  ];
  if (last === first + 1) {
    return edges;
  }
  return [edges[0], { first: first + 1, last: last - 1, atStart: false, atEnd: false }, edges[1]];
}

/** The regions a pass must report, derived from the engine's own pre-pass read. */
function expectedRegions(id, rect, base, outline) {
  const rowBands = outline ? bandsOf(rect.firstRow, rect.lastRow) : [bandsOf(0, 1)[0]];
  const bandRows = outline
    ? rowBands
    : [{ first: rect.firstRow, last: rect.lastRow, atStart: true, atEnd: true }];
  const bandCols = outline
    ? bandsOf(rect.firstCol, rect.lastCol)
    : [{ first: rect.firstCol, last: rect.lastCol, atStart: true, atEnd: true }];
  const expected = [];
  for (const rowBand of bandRows) {
    for (const colBand of bandCols) {
      const bases = engine(
        id,
        "rangeStyleBases",
        0,
        rowBand.first,
        colBand.first,
        rowBand.last,
        colBand.last,
      ).xfIndices;
      const range = `${cellToA1(rowBand.first, colBand.first)}:${cellToA1(rowBand.last, colBand.last)}`;
      for (const baseXfIndex of base === "default" ? [0] : bases) {
        expected.push({ range, baseXfIndex });
      }
    }
  }
  return expected;
}

const regionKey = (region) => `${region.range}|${region.baseXfIndex}`;

/** Cells a pass is expected to add to the sheet. */
function expectedCellGrowth(priorName, base, rect) {
  if (!isWholeLines(rect)) {
    const area = (rect.lastRow - rect.firstRow + 1) * (rect.lastCol - rect.firstCol + 1);
    // C3 already holds a value when the prior state is "numfmt".
    return area - (priorName === "numfmt" ? 1 : 0);
  }
  // A row style wins over a column style, so a whole column crossing a styled
  // row must materialize the crossing cells to show through. A whole row
  // crossing a styled column wins on its own and only has to preserve what the
  // column contributed, which matters for "existing" and not for "default".
  if (priorName === "rowstyled" && wholeColumns(rect)) {
    return rect.lastCol - rect.firstCol + 1;
  }
  if (priorName === "colstyled" && wholeRows(rect) && base === "existing") {
    return rect.lastRow - rect.firstRow + 1;
  }
  return 0;
}

function assertWholeLineMetadata(id, rect, before) {
  const fillOf = (xfIndex) => {
    const xf = engine(id, "getCellXf", xfIndex);
    return engine(id, "getFill", xf.fillIndex).fgArgb;
  };
  const columns = engine(id, "getSheetColumns", 0).columns;
  const rows = engine(id, "getSheetRowOverrides", 0).rows;
  if (wholeColumns(rect)) {
    for (let col = rect.firstCol; col <= rect.lastCol; col += 1) {
      const entry = columns.find((item) => item.first <= col && col <= item.last);
      assert.ok(entry, `column ${col} carries a style entry`);
      assert.equal(entry.hasStyle, 1);
      assert.equal(fillOf(entry.styleXf), ORANGE);
    }
    assert.deepEqual(rows, before.rows, "a column pass leaves row overrides alone");
  } else {
    for (let row = rect.firstRow; row <= rect.lastRow; row += 1) {
      const entry = rows.find((item) => item.row === row);
      assert.ok(entry, `row ${row} carries a style entry`);
      assert.equal(entry.hasStyle, 1);
      assert.equal(fillOf(entry.styleXf), ORANGE);
    }
    assert.deepEqual(columns, before.columns, "a row pass leaves column entries alone");
  }
}

/**
 * Pairwise matrix (shape x base x outline x prior; the first 29 rows cover every
 * value pair) run through both entry points.
 */
const STYLE_CASES = [
  ["col", "existing", "absent", "blank"],
  ["cell", "default", "absent", "colstyled"],
  ["col", "default", "present", "rowstyled"],
  ["cols", "existing", "present", "colstyled"],
  ["rect", "existing", "absent", "rowstyled"],
  ["cell", "existing", "present", "numfmt"],
  ["rows", "default", "absent", "colstyled"],
  ["sheetcol", "existing", "absent", "numfmt"],
  ["row", "default", "present", "blank"],
  ["row", "existing", "absent", "numfmt"],
  ["cols", "default", "absent", "blank"],
  ["rows", "existing", "present", "colstyled"],
  ["sheetcol", "existing", "present", "colstyled"],
  ["rect", "default", "present", "numfmt"],
  ["cols", "existing", "present", "rowstyled"],
  ["sheetcol", "default", "present", "rowstyled"],
  ["cols", "existing", "absent", "numfmt"],
  ["col", "default", "absent", "colstyled"],
  ["col", "existing", "present", "numfmt"],
  ["rows", "existing", "present", "numfmt"],
  ["rows", "existing", "present", "rowstyled"],
  ["rows", "default", "absent", "blank"],
  ["sheetcol", "default", "present", "blank"],
  ["cell", "existing", "absent", "rowstyled"],
  ["cell", "default", "absent", "blank"],
  ["rect", "existing", "present", "blank"],
  ["row", "existing", "absent", "colstyled"],
  ["rect", "existing", "present", "colstyled"],
  ["row", "existing", "present", "rowstyled"],
  // Beyond the pairwise rows: a whole line crossing a styled perpendicular line.
  ["col", "default", "absent", "rowstyled"],
  ["cols", "default", "absent", "rowstyled"],
  ["cols", "existing", "absent", "rowstyled"],
  ["row", "default", "absent", "colstyled"],
  ["sheetcol", "existing", "absent", "rowstyled"],
];

async function runStyleCase(entryName, [shapeName, base, outlineFlag, priorName]) {
  const id = `style-matrix-${entryName}-${shapeName}-${base}-${outlineFlag}-${priorName}`;
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-style-"));
  const shape = SHAPES[shapeName];
  const outline = outlineFlag === "present";
  const style = {
    fill: { color: "#FFC000" },
    font: { italic: true },
    ...(outline ? { border: { outline: "medium" } } : {}),
  };
  await openSession(undefined, id);
  try {
    engine(id, "addSheet", "Sheet2");
    PRIORS[priorName](id);
    // Start from a clean session so a refused pass can be shown not to dirty it.
    await saveSession(id, path.join(dir, "prior.xlsx"));
    assert.equal(sessionDirty(id), false);
    const before = snapshot(id);
    const applyStyle = () => ENTRY_POINTS[entryName](id, shape, style, base);

    if (outline && isWholeLines(shape.rect)) {
      assert.throws(applyStyle, /outline border needs a bounded range/);
      assert.deepEqual(snapshot(id), before, "a refused pass changes nothing");
      assert.equal(sessionDirty(id), false);
      return;
    }

    const expected = expectedRegions(id, shape.rect, base, outline);
    const applied = applyStyle();
    assert.equal(sessionDirty(id), true);

    // Regions: one per distinct base per band, with no per-region cell count.
    assert.deepEqual(applied.regions.map(regionKey).sort(), expected.map(regionKey).sort());
    for (const region of applied.regions) {
      assert.deepEqual(Object.keys(region).sort(), ["baseXfIndex", "range", "xfIndex"]);
      assert.ok(region.xfIndex >= before.xfCount, "the patched xf is newly added");
    }
    assert.equal(new Set(applied.regions.map((region) => region.xfIndex)).size, expected.length);

    const after = snapshot(id);
    assert.equal(
      after.cellCount - before.cellCount,
      expectedCellGrowth(priorName, base, shape.rect),
    );
    assert.equal(after.otherSheetCells, before.otherSheetCells);
    assert.deepEqual(after.otherSheetStyle, before.otherSheetStyle);

    const { rect } = shape;
    const bounded = !isWholeLines(rect);
    for (const [index, cell] of after.grid.entries()) {
      const beforeCell = before.grid[index];
      const inside =
        cell.row >= rect.firstRow &&
        cell.row <= rect.lastRow &&
        cell.col >= rect.firstCol &&
        cell.col <= rect.lastCol;
      if (!inside) {
        assert.deepEqual(cell.style, beforeCell.style, `r${cell.row}c${cell.col} is untouched`);
        continue;
      }
      assert.equal(cell.style.fillForeground.argb, ORANGE, `r${cell.row}c${cell.col} fill`);
      assert.equal(engine(id, "getFont", cell.style.fontIndex).italic, true);
      // "existing" is a delta over what the cell showed; "default" starts from xf 0.
      assert.equal(
        cell.style.numFmtCode,
        base === "existing" ? beforeCell.style.numFmtCode : "General",
        `r${cell.row}c${cell.col} number format`,
      );
      if (bounded) {
        // Blank cells of a bounded range are materialized with a cell-level style.
        assert.equal(cell.style.source, 0, `r${cell.row}c${cell.col} is a styled cell`);
      }
      if (outline) {
        const border = engine(id, "getBorder", cell.style.borderIndex);
        const onEdge = {
          top: cell.row === rect.firstRow,
          bottom: cell.row === rect.lastRow,
          left: cell.col === rect.firstCol,
          right: cell.col === rect.lastCol,
        };
        for (const side of ["top", "bottom", "left", "right"]) {
          assert.equal(
            border[side].style,
            onEdge[side] ? MEDIUM : 0,
            `${side} of r${cell.row}c${cell.col}`,
          );
        }
      }
    }

    if (bounded) {
      assert.deepEqual(after.columns, before.columns, "a bounded pass adds no column entry");
      assert.deepEqual(after.rows, before.rows, "a bounded pass adds no row entry");
    } else {
      assertWholeLineMetadata(id, rect, before);
      // A value written later into the styled line inherits the new style.
      const [row, col] = wholeColumns(rect) ? [40, rect.firstCol] : [rect.firstRow, 30];
      const blankBefore = effectiveAt(id, row, col);
      engine(id, "setNumber", 0, row, col, 7);
      const written = effectiveAt(id, row, col);
      assert.equal(written.fillForeground.argb, ORANGE);
      assert.equal(written.numFmtCode, base === "existing" ? blankBefore.numFmtCode : "General");
      assert.equal(engine(id, "getFont", written.fontIndex).italic, true);
    }
  } finally {
    closeSession(id);
    await rm(dir, { recursive: true, force: true });
  }
}

for (const entryName of Object.keys(ENTRY_POINTS)) {
  for (const row of STYLE_CASES) {
    test(`style matrix via ${entryName}: ${row.join(" / ")}`, async () => {
      await runStyleCase(entryName, row);
    });
  }
}

const runStyle = {
  styleSessionRange: (id, ref, style, base) => styleSessionRange(id, ref, style, undefined, base),
  applyLayoutStyle: (id, ref, style, base) =>
    applySessionLayout(id, 0, [{ type: "style", range: ref, style, base }]).operations[0],
};

for (const [entryName, run] of Object.entries(runStyle)) {
  for (const ref of ["A1:A3", "A:A"]) {
    test(`keeps per-cell formats under a fill-only pass, ${entryName} ${ref}`, async () => {
      const id = `style-delta-cells-${entryName}-${ref.replace(":", "-")}`;
      await openSession(undefined, id);
      try {
        setSessionRange(id, "A1", [[0.5], [2000], ["text"]], undefined, false);
        styleSessionRange(id, "A1", { numberFormat: "0.0%", font: { bold: true } }, 0, "existing");
        styleSessionRange(id, "A2", { numberFormat: '"¥"#,##0' }, 0, "existing");
        const priorFormats = [0, 1, 2].map((row) => effectiveAt(id, row, 0).numFmtCode);
        assert.deepEqual(priorFormats, ["0.0%", '"¥"#,##0', "General"]);

        run(id, ref, { fill: { color: "#FFC000" } }, "existing");
        for (const row of [0, 1, 2]) {
          const style = effectiveAt(id, row, 0);
          assert.equal(style.fillForeground.argb, ORANGE);
          assert.equal(style.numFmtCode, priorFormats[row]);
          assert.equal(engine(id, "getFont", style.fontIndex).bold, row === 0);
        }

        run(id, ref, { fill: { color: "#00B050" } }, "default");
        for (const row of [0, 1, 2]) {
          const style = effectiveAt(id, row, 0);
          assert.equal(style.fillForeground.argb, parseColor("#00B050"));
          assert.equal(style.numFmtCode, "General");
          assert.equal(engine(id, "getFont", style.fontIndex).bold, false);
        }
      } finally {
        closeSession(id);
      }
    });
  }

  for (const ref of ["A:XFD", "1:1048576", "XFD:A", "1048576:1"]) {
    test(`styles a full-grid span without a cell cap, ${entryName} ${ref}`, async () => {
      const id = `style-full-${entryName}-${ref.replace(":", "-")}`;
      await openSession(undefined, id);
      try {
        const applied = run(id, ref, { font: { bold: true } }, "existing");
        assert.equal(applied.regions.length, 1);
        assert.equal(applied.regions[0].range, "A1:XFD1048576");
        assert.equal(engine(id, "cellCount", 0).value, 0);
        for (const [row, col] of [
          [0, 0],
          [500, 700],
          [MAX_ROW, MAX_COL],
        ]) {
          assert.equal(engine(id, "getFont", effectiveAt(id, row, col).fontIndex).bold, true);
        }
      } finally {
        closeSession(id);
      }
    });
  }

  test(`orders reversed spans before styling, ${entryName}`, async () => {
    const id = `style-reversed-${entryName}`;
    await openSession(undefined, id);
    try {
      run(id, "D4:B2", { fill: { color: "#FFC000" } }, "existing");
      run(id, "F:E", { fill: { color: "#FFC000" } }, "existing");
      run(id, "9:8", { fill: { color: "#FFC000" } }, "existing");
      const filled = (row, col) => effectiveAt(id, row, col).fillForeground.argb === ORANGE;
      assert.equal(filled(1, 1) && filled(3, 3) && filled(2, 2), true);
      assert.equal(filled(0, 0) || filled(4, 1), false);
      assert.equal(filled(50, 4) && filled(50, 5), true);
      assert.equal(filled(50, 3) || filled(50, 6), false);
      assert.equal(filled(7, 99) && filled(8, 99), true);
      assert.equal(filled(6, 99) || filled(9, 99), false);
    } finally {
      closeSession(id);
    }
  });
}

test("applies a qualified whole column to the named sheet only", async () => {
  await openSession(undefined, "style-qualified");
  try {
    engine("style-qualified", "addSheet", "Data Sheet");
    const applied = styleSessionRange(
      "style-qualified",
      "'Data Sheet'!B:C",
      { fill: { color: "#FFC000" } },
      0,
      "existing",
    );
    assert.equal(applied.range.sheet, 1);
    assert.equal(applied.range.sheetName, "Data Sheet");
    assert.equal(engine("style-qualified", "getSheetColumns", 1).columns.length, 1);
    assert.equal(engine("style-qualified", "getSheetColumns", 0).columns.length, 0);

    // The layout route only styles the selected sheet; a qualifier elsewhere is refused.
    assert.throws(
      () =>
        applySessionLayout("style-qualified", 0, [
          { type: "style", range: "'Data Sheet'!B:C", style: { font: { bold: true } } },
        ]),
      /must refer to the selected sheet/,
    );
    const viaLayout = applySessionLayout("style-qualified", "Data Sheet", [
      { type: "style", range: "'Data Sheet'!D:D", style: { font: { bold: true } } },
    ]);
    assert.equal(viaLayout.sheetName, "Data Sheet");
    assert.equal(engine("style-qualified", "getSheetColumns", 1).columns.length, 2);
    assert.equal(engine("style-qualified", "getSheetColumns", 0).columns.length, 0);
  } finally {
    closeSession("style-qualified");
  }
});

test("rejects malformed selections and unknown bases in both routes", async () => {
  await openSession(undefined, "style-invalid");
  try {
    for (const ref of [
      "0:1",
      "A:1",
      "1:A",
      ":",
      "A:",
      "",
      "A1:B2:C3",
      "XFE:XFE",
      "1048577:1048577",
    ]) {
      assert.throws(
        () => styleSessionRange("style-invalid", ref, { font: { bold: true } }, 0, "existing"),
        Error,
        `styleSessionRange accepted ${JSON.stringify(ref)}`,
      );
      assert.throws(
        () =>
          applySessionLayout("style-invalid", 0, [
            { type: "style", range: ref, style: { font: { bold: true } } },
          ]),
        Error,
        `apply_layout accepted ${JSON.stringify(ref)}`,
      );
    }
    assert.throws(
      () =>
        applySessionLayout("style-invalid", 0, [
          { type: "style", range: "A1", style: { font: { bold: true } }, base: "other" },
        ]),
      /base must be existing or default/,
    );
    assert.equal(sessionDirty("style-invalid"), false);
  } finally {
    closeSession("style-invalid");
  }
});

test("caps a bounded style pass at the cell limit and leaves the session clean", async () => {
  await openSession(undefined, "style-cap");
  try {
    // 100 columns x 1,001 rows is one row past the limit.
    assert.throws(
      () => styleSessionRange("style-cap", "A1:CV1001", { font: { bold: true } }, 0, "default"),
      /spans 100100 cells \(limit 100000\); style a smaller range/,
    );
    assert.throws(
      () => styleSessionRange("style-cap", "A1:A100001", { font: { bold: true } }, 0, "default"),
      /spans 100001 cells/,
    );
    assert.equal(sessionDirty("style-cap"), false);
    assert.equal(engine("style-cap", "cellCount", 0).value, 0);

    // Exactly at the limit is allowed and every cell is materialized.
    styleSessionRange("style-cap", "A1:CV1000", { font: { bold: true } }, 0, "default");
    assert.equal(engine("style-cap", "cellCount", 0).value, MAX_STYLED_CELLS);
    assert.equal(effectiveAt("style-cap", 999, 99).source, 0);
  } finally {
    closeSession("style-cap");
  }
});

test("counts bounded cells cumulatively across layout style operations", async () => {
  await openSession(undefined, "layout-cap");
  const bold = { font: { bold: true } };
  const half = (range) => ({ type: "style", range, style: bold });
  try {
    // 50,000 + 50,001 crosses the limit, so the whole batch is refused up front.
    assert.throws(
      () => applySessionLayout("layout-cap", 0, [half("A1:J5000"), half("L1:U5000"), half("W1")]),
      /style operations exceed the 100000-cell limit/,
    );
    assert.throws(
      () => applySessionLayout("layout-cap", 0, [half("A1:J10001")]),
      /style operations exceed the 100000-cell limit/,
    );
    assert.equal(sessionDirty("layout-cap"), false);
    assert.equal(engine("layout-cap", "cellCount", 0).value, 0);

    // Whole rows and columns count as zero, so they fit beside a full allotment.
    // The two blocks are adjacent because the engine stores a row densely
    // between its first and last stored column.
    const applied = applySessionLayout("layout-cap", 0, [
      half("A1:J5000"),
      half("K1:T5000"),
      half("A:XFD"),
      half("1:1048576"),
    ]);
    assert.equal(applied.operations.length, 4);
    assert.equal(engine("layout-cap", "cellCount", 0).value, MAX_STYLED_CELLS);
    assert.equal(sessionDirty("layout-cap"), true);
  } finally {
    closeSession("layout-cap");
  }
});

test("refuses an outline on whole lines in a batch before applying any operation", async () => {
  await openSession(undefined, "layout-outline-batch");
  try {
    assert.throws(
      () =>
        applySessionLayout("layout-outline-batch", 0, [
          { type: "style", range: "A1:B2", style: { fill: { color: "#FFC000" } } },
          { type: "style", range: "C:D", style: { border: { outline: "thin" } } },
        ]),
      /outline border needs a bounded range/,
    );
    // The first operation was valid, but preflight rejects the batch as a whole.
    assert.equal(engine("layout-outline-batch", "cellCount", 0).value, 0);
    assert.equal(sessionDirty("layout-outline-batch"), false);

    // "all" on whole lines is fine: only an outline needs edges.
    applySessionLayout("layout-outline-batch", 0, [
      { type: "style", range: "C:D", style: { border: { all: "thin" } } },
    ]);
    assert.equal(engine("layout-outline-batch", "getSheetColumns", 0).columns.length, 1);
  } finally {
    closeSession("layout-outline-batch");
  }
});

test("styles far-apart columns over many rows without dense row storage", async () => {
  await openSession(undefined, "style-far-columns");
  try {
    // A and XFD in the same rows: each row stores its two cells, not the gap.
    for (const ref of ["A1:A50000", "XFD1:XFD50000"]) {
      styleSessionRange(
        "style-far-columns",
        ref,
        { fill: { color: "#FFFF00" } },
        undefined,
        "existing",
      );
    }
    assert.equal(callWorkbookMethod("style-far-columns", "cellCount", [0]).result.value, 100_000);
    const last = callWorkbookMethod("style-far-columns", "cellAt", [0, 99_999]).result;
    assert.deepEqual([last.row, last.col], [49_999, 16_383]);

    // A fill-only pass over a cell carrying its own style reports only that base.
    const applied = styleSessionRange(
      "style-far-columns",
      "XFD7",
      { font: { bold: true } },
      undefined,
      "existing",
    );
    assert.equal(applied.regions.length, 1);
    assert.notEqual(applied.regions[0].baseXfIndex, 0);
  } finally {
    closeSession("style-far-columns");
  }
});
