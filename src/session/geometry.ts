import { GeometryMode, type Workbook } from "../formulon.js";
import type { ColumnMetrics } from "../preview/geometry.js";

/** Reads the engine's display width model (Normal font calibration) for one sheet. */
export function resolveColumnMetrics(workbook: Workbook, sheet: number): ColumnMetrics {
  const model = workbook.sheet(sheet).getWidthModel(GeometryMode.Display);
  return {
    fontName: model.normalFontName,
    fontSize: model.normalFontSizePt,
    pointsPerChar: model.pointsPerChar,
    paddingPt: model.paddingPt,
    approximate: !model.calibrated,
  };
}

/** Converts a stored Excel column width (character units) to display points. */
export function columnCharsToPoints(workbook: Workbook, sheet: number, chars: number): number {
  return workbook.sheet(sheet).convertColumnCharsToPt(chars, GeometryMode.Display);
}

/** Converts display points to Excel's stored column width (character units, padding included). */
export function columnPointsToChars(workbook: Workbook, sheet: number, points: number): number {
  return workbook.sheet(sheet).convertColumnPtToChars(points, GeometryMode.Display);
}

export type { ColumnMetrics } from "../preview/geometry.js";
export {
  mmToPoints,
  pixelsToPoints,
  pointsToMm,
  pointsToPixels,
} from "../preview/geometry.js";
