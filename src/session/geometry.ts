import { assertStatus, GeometryMode, requireNumberResult, type Workbook } from "../formulon.js";
import type { ColumnMetrics } from "../preview/geometry.js";

/** Reads the engine's display width model (Normal font calibration) for one sheet. */
export function resolveColumnMetrics(workbook: Workbook, sheet: number): ColumnMetrics {
  const model = workbook.getWidthModel(sheet, GeometryMode.Display);
  assertStatus(model.status, "read column width model");
  return {
    fontName: model.normalFontName,
    fontSize: model.normalFontSize,
    pointsPerChar: model.pointsPerChar,
    paddingPt: model.paddingPt,
    approximate: !model.calibrated,
  };
}

/** Converts a stored Excel column width (character units) to display points. */
export function columnCharsToPoints(workbook: Workbook, sheet: number, chars: number): number {
  return requireNumberResult(
    workbook.columnCharsToPt(sheet, GeometryMode.Display, chars),
    "convert column width to points",
  );
}

/** Converts display points to character units; widths at or below the cell padding are rejected. */
export function columnPointsToChars(workbook: Workbook, sheet: number, points: number): number {
  const chars = requireNumberResult(
    workbook.columnPtToChars(sheet, GeometryMode.Display, points),
    "convert column width to characters",
  );
  if (chars === 0 && points > 0) {
    const padding = resolveColumnMetrics(workbook, sheet).paddingPt;
    throw new Error(
      `column width in points must exceed the ${padding}pt visible-column padding, or be exactly zero: ${points}`,
    );
  }
  return chars;
}

export type { ColumnMetrics } from "../preview/geometry.js";
export {
  mmToPoints,
  pixelsToPoints,
  pointsToMm,
  pointsToPixels,
} from "../preview/geometry.js";
