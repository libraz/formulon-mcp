import assert from "node:assert/strict";
import { test } from "vitest";
import {
  callWorkbookMethod,
  closeSession,
  openSession,
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
    assert.equal(applied.regions[0].cellCount, 2);

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
    assert.deepEqual(
      applied.regions.map((region) => region.cellCount),
      [1, 1],
    );
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
