/**
 * The display calibration used by Excel's column-width character units.
 *
 * These values come from Formulon's pagination calibration.  They describe
 * the 96-DPI display width (the value an author sees while laying out a
 * sheet), rather than the separate print-track calibration used for page
 * breaks.  Unknown font families and sizes intentionally use Calibri 11's
 * measured values and are marked approximate.
 */
export type ColumnMetrics = {
  fontName: string;
  fontSize: number;
  pointsPerChar: number;
  paddingPt: number;
  approximate: boolean;
};

type Calibration = {
  family: string;
  size: number;
  pointsPerChar: number;
  paddingPt: number;
};

const CALIBRATIONS: readonly Calibration[] = [
  { family: "Calibri", size: 8, pointsPerChar: 4.5, paddingPt: 3.75 },
  { family: "Calibri", size: 9, pointsPerChar: 4.5, paddingPt: 3.75 },
  { family: "Calibri", size: 10, pointsPerChar: 5.25, paddingPt: 3.75 },
  { family: "Calibri", size: 11, pointsPerChar: 5.25, paddingPt: 3.75 },
  { family: "Calibri", size: 12, pointsPerChar: 6, paddingPt: 3.75 },
  { family: "Calibri", size: 14, pointsPerChar: 7.5, paddingPt: 5.25 },
  { family: "Calibri", size: 16, pointsPerChar: 8.25, paddingPt: 5.25 },
  { family: "Calibri", size: 18, pointsPerChar: 9, paddingPt: 5.25 },
  { family: "ＭＳ Ｐゴシック", size: 8, pointsPerChar: 4.5, paddingPt: 3.75 },
  { family: "ＭＳ Ｐゴシック", size: 9, pointsPerChar: 4.5, paddingPt: 3.75 },
  { family: "ＭＳ Ｐゴシック", size: 10, pointsPerChar: 5.25, paddingPt: 3.75 },
  { family: "ＭＳ Ｐゴシック", size: 11, pointsPerChar: 6, paddingPt: 3.75 },
  { family: "ＭＳ Ｐゴシック", size: 12, pointsPerChar: 6, paddingPt: 3.75 },
  { family: "ＭＳ Ｐゴシック", size: 14, pointsPerChar: 7.5, paddingPt: 5.25 },
  { family: "ＭＳ Ｐゴシック", size: 16, pointsPerChar: 8.25, paddingPt: 5.25 },
  { family: "ＭＳ Ｐゴシック", size: 18, pointsPerChar: 9, paddingPt: 5.25 },
  { family: "游ゴシック", size: 8, pointsPerChar: 4.5, paddingPt: 3.75 },
  { family: "游ゴシック", size: 9, pointsPerChar: 5.25, paddingPt: 3.75 },
  { family: "游ゴシック", size: 10, pointsPerChar: 5.25, paddingPt: 3.75 },
  { family: "游ゴシック", size: 11, pointsPerChar: 6, paddingPt: 3.75 },
  { family: "游ゴシック", size: 12, pointsPerChar: 6.75, paddingPt: 5.25 },
  { family: "游ゴシック", size: 14, pointsPerChar: 8.25, paddingPt: 5.25 },
  { family: "游ゴシック", size: 16, pointsPerChar: 9, paddingPt: 5.25 },
  { family: "游ゴシック", size: 18, pointsPerChar: 9.75, paddingPt: 6.75 },
  { family: "Meiryo UI", size: 8, pointsPerChar: 5.25, paddingPt: 3.75 },
  { family: "Meiryo UI", size: 9, pointsPerChar: 5.25, paddingPt: 3.75 },
  { family: "Meiryo UI", size: 10, pointsPerChar: 6, paddingPt: 3.75 },
  { family: "Meiryo UI", size: 11, pointsPerChar: 6.75, paddingPt: 5.25 },
  { family: "Meiryo UI", size: 12, pointsPerChar: 7.5, paddingPt: 5.25 },
  { family: "Meiryo UI", size: 14, pointsPerChar: 9, paddingPt: 5.25 },
  { family: "Meiryo UI", size: 16, pointsPerChar: 9.75, paddingPt: 6.75 },
  { family: "Meiryo UI", size: 18, pointsPerChar: 11.25, paddingPt: 6.75 },
];

const FALLBACK_POINTS_PER_CHAR = 5.25;
const FALLBACK_PADDING_PT = 3.75;
const POINTS_PER_PIXEL = 72 / 96;
const POINTS_PER_MM = 72 / 25.4;

function nonNegativeFinite(value: number, what: string): number {
  if (!Number.isFinite(value) || value < 0) {
    throw new Error(`${what} must be a finite nonnegative number: ${value}`);
  }
  return value;
}

/** Returns the measured display calibration for a Normal font. */
export function getColumnMetrics(fontName: string, fontSize: number): ColumnMetrics {
  if (typeof fontName !== "string" || fontName.length === 0) {
    throw new Error(`fontName must be a non-empty string: ${fontName}`);
  }
  if (!Number.isFinite(fontSize) || fontSize <= 0) {
    throw new Error(`fontSize must be a finite positive number: ${fontSize}`);
  }
  const calibration = CALIBRATIONS.find(
    (entry) => entry.family === fontName && entry.size === fontSize,
  );
  return {
    fontName,
    fontSize,
    pointsPerChar: calibration?.pointsPerChar ?? FALLBACK_POINTS_PER_CHAR,
    paddingPt: calibration?.paddingPt ?? FALLBACK_PADDING_PT,
    approximate: calibration === undefined,
  };
}

/** Converts a stored Excel column width (character units) to display points. */
export function columnCharsToPoints(chars: number, metrics: ColumnMetrics): number {
  nonNegativeFinite(chars, "column width");
  if (chars === 0) {
    return 0;
  }
  return chars * metrics.pointsPerChar + metrics.paddingPt;
}

/** Converts display points to Excel's stored column-width character units. */
export function columnPointsToChars(points: number, metrics: ColumnMetrics): number {
  nonNegativeFinite(points, "column width in points");
  if (points === 0) {
    return 0;
  }
  if (points <= metrics.paddingPt) {
    throw new Error(
      `column width in points must exceed the ${metrics.paddingPt}pt visible-column padding, or be exactly zero: ${points}`,
    );
  }
  return (points - metrics.paddingPt) / metrics.pointsPerChar;
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
