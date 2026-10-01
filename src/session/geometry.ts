import type { Workbook } from "../formulon.js";
import { type ColumnMetrics, getColumnMetrics } from "../preview/geometry.js";

const FALLBACK_FONT_NAME = "Calibri";
const FALLBACK_FONT_SIZE = 11;

function usableFont(value: unknown): value is { name: string; size: number } {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const font = value as { name?: unknown; size?: unknown };
  return (
    typeof font.name === "string" &&
    font.name.length > 0 &&
    typeof font.size === "number" &&
    Number.isFinite(font.size) &&
    font.size > 0
  );
}

function resultFont(workbook: Workbook, index: number): { name: string; size: number } | undefined {
  try {
    const result = workbook.getFont(index);
    return result.status.ok && usableFont(result) ? result : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Resolves the font Excel uses for column character units.
 *
 * The named Normal style points to a cell-style XF, which in turn points to a
 * font.  A malformed or incomplete style table follows Formulon's pagination
 * behavior and falls back to font slot zero, then to Calibri 11's seeded
 * default.
 */
export function resolveColumnMetrics(workbook: Workbook): ColumnMetrics {
  let font: { name: string; size: number } | undefined;
  try {
    const count = workbook.cellStyleCount();
    if (count.status.ok && Number.isInteger(count.value) && count.value >= 0) {
      for (let index = 0; index < count.value; index += 1) {
        const style = workbook.getCellStyle(index);
        if (!style.status.ok || style.name !== "Normal") {
          continue;
        }
        const xf = workbook.getCellStyleXf(style.xfId);
        if (xf.status.ok && Number.isInteger(xf.fontIndex) && xf.fontIndex >= 0) {
          font = resultFont(workbook, xf.fontIndex);
        }
        break;
      }
    }
  } catch {
    // The font-zero fallback below also handles an absent style table.
  }
  font ??= resultFont(workbook, 0);
  if (!font) {
    return {
      ...getColumnMetrics(FALLBACK_FONT_NAME, FALLBACK_FONT_SIZE),
      approximate: true,
    };
  }
  return getColumnMetrics(font.name, font.size);
}

export type { ColumnMetrics } from "../preview/geometry.js";
export {
  columnCharsToPoints,
  columnPointsToChars,
  mmToPoints,
  pixelsToPoints,
  pointsToMm,
  pointsToPixels,
} from "../preview/geometry.js";
