/** The engine's display width model for a sheet's Normal font. */
export type ColumnMetrics = {
  fontName: string;
  fontSize: number;
  pointsPerChar: number;
  paddingPt: number;
  approximate: boolean;
};

const POINTS_PER_PIXEL = 72 / 96;
const POINTS_PER_MM = 72 / 25.4;

function nonNegativeFinite(value: number, what: string): number {
  if (!Number.isFinite(value) || value < 0) {
    throw new Error(`${what} must be a finite nonnegative number: ${value}`);
  }
  return value;
}

/** Converts CSS/display pixels at Excel's 96-DPI reference scale to points. */
export function pixelsToPoints(px: number): number {
  return nonNegativeFinite(px, "pixels") * POINTS_PER_PIXEL;
}

/** Converts display points to CSS/display pixels at Excel's 96-DPI scale. */
export function pointsToPixels(pt: number): number {
  return nonNegativeFinite(pt, "points") / POINTS_PER_PIXEL;
}

/** Converts millimetres to display points. */
export function mmToPoints(mm: number): number {
  return nonNegativeFinite(mm, "millimetres") * POINTS_PER_MM;
}

/** Converts display points to millimetres. */
export function pointsToMm(pt: number): number {
  return nonNegativeFinite(pt, "points") / POINTS_PER_MM;
}
