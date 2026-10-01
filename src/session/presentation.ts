import { cellToA1, normalizeRange, parseCellRef, parseRangeRef } from "../a1.js";
import { buildDocument, type DocumentSpec } from "../document.js";
import { findSheetIndex } from "../formulon.js";
import { type PrintSettingsInput, readPrintSettings, writePrintSettings } from "../print.js";
import {
  applyDefaultFont,
  applyRangeStyle,
  type FontStyleInput,
  readDefaultFont,
  type StyleBase,
  type StyleInput,
} from "../styles.js";
import { assertInGrid, MAX_RANGE_CELLS, rangeA1, sheetName } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

/**
 * Lays out and writes a block-composed document, then reports where every block
 * and named cell landed so the caller can refine it with the primitive tools.
 */
export function buildSessionDocument(
  id: string,
  spec: DocumentSpec,
  sheet: number | string | undefined,
) {
  const session = getSession(id);
  const anchor = parseCellRef(spec.start ?? "B2");
  const sheetIndex = findSheetIndex(session.workbook, anchor.sheetName ?? sheet);
  assertInGrid(anchor.row, anchor.col);
  touch(session, true);
  const document = buildDocument(
    session.workbook,
    sheetIndex,
    { row: anchor.row, col: anchor.col },
    spec,
  );
  touch(session, true);
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    sheetName: sheetName(session, sheetIndex),
    ...document,
  };
}

/**
 * Reads or writes a sheet's print settings.
 *
 * A read also reports the page count the settings produce, so the effect of a
 * write can be checked without saving and opening the file.
 */
export function sessionPrintSettings(
  id: string,
  sheet: number | string | undefined,
  input?: PrintSettingsInput,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  if (input && Object.values(input).every((value) => value === undefined)) {
    throw new Error("no print setting was given to set");
  }
  if (input) {
    touch(session, true);
  }
  const applied = input ? writePrintSettings(session.workbook, sheetIndex, input) : undefined;
  touch(session, applied !== undefined);
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    sheetName: sheetName(session, sheetIndex),
    applied,
    settings: readPrintSettings(session.workbook, sheetIndex),
  };
}

/**
 * Reads the workbook default font, or redeclares it from the stated deltas.
 *
 * The default is font 0, which every cell that was never styled resolves to,
 * so this is workbook-wide rather than a range operation: a new workbook is
 * seeded with Excel's Calibri, and nothing else can say what a plain cell of
 * Japanese text should be saved as.
 */
export function sessionDefaultFont(id: string, font?: FontStyleInput) {
  const session = getSession(id);
  if (font === undefined) {
    return { session: publicInfo(session), font: readDefaultFont(session.workbook) };
  }
  touch(session, true);
  const applied = applyDefaultFont(session.workbook, font);
  return { session: publicInfo(session), font: applied };
}

/**
 * Applies a declarative style across an A1 range, materializing blank cells so
 * an empty ruled box renders. Cells keep whatever the style did not state, so
 * ruling a table does not undo the number format already on its amount column.
 */
export function styleSessionRange(
  id: string,
  ref: string,
  style: StyleInput,
  fallbackSheet: number | string | undefined,
  base: StyleBase,
) {
  const session = getSession(id);
  const range = normalizeRange(parseRangeRef(ref));
  const sheetIndex = findSheetIndex(session.workbook, range.sheetName ?? fallbackSheet);
  assertInGrid(range.end.row, range.end.col);
  // Styling materializes every cell in the rectangle, so an over-wide range is
  // refused rather than turned into a sheet of millions of styled blanks.
  const cells = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
  if (cells > MAX_RANGE_CELLS) {
    throw new Error(
      `range ${ref} spans ${cells} cells (limit ${MAX_RANGE_CELLS}); style a smaller range`,
    );
  }
  touch(session, true);
  const applied = applyRangeStyle(
    session.workbook,
    sheetIndex,
    {
      firstRow: range.start.row,
      firstCol: range.start.col,
      lastRow: range.end.row,
      lastCol: range.end.col,
    },
    style,
    base,
  );
  return {
    session: publicInfo(session),
    range: {
      sheet: sheetIndex,
      sheetName: sheetName(session, sheetIndex),
      start: cellToA1(range.start.row, range.start.col),
      end: cellToA1(range.end.row, range.end.col),
    },
    regions: applied.regions.map((region) => ({
      range: rangeA1(region.firstRow, region.firstCol, region.lastRow, region.lastCol),
      cellCount: region.cellCount,
      baseXfIndex: region.baseXfIndex,
      xfIndex: region.xfIndex,
    })),
  };
}
