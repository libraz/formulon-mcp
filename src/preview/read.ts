import type { ResolvedColor } from "@libraz/formulon";
import { cellToA1, normalizeRange, parseRangeRef } from "../a1.js";
import {
  assertStatus,
  ColorResolution,
  DisplayStatus,
  GeometryMode,
  requireCell,
  requireListResult,
  requireNumberResult,
  type Workbook,
} from "../formulon.js";
import type { ColumnMetrics } from "./geometry.js";
import type { PreviewCell, PreviewCellMetadata, PreviewStyle } from "./types.js";

export const MAX_METADATA_CELLS = 500;
export const MAX_METADATA_TEXT = 256;
const MAX_CELL_SCAN_CELLS = 1_000_000;
const MAX_PREVIEW_MERGES = 10_000;
const FORMULA_PAGE_SIZE = 4_096;

export const BASE_WARNINGS = [
  "Preview rendering is approximate and may differ from Excel.",
  "Font substitution is possible when the requested workbook font is unavailable.",
  "Conditional formatting is not applied.",
  "Headers and footers are not rendered.",
];

export type Rect = { firstRow: number; firstCol: number; lastRow: number; lastCol: number };
export type MergeRange = Rect;
export type WarningState = { values: Set<string>; truncated: boolean };

export function warning(state: WarningState, message: string): void {
  if (state.values.has(message)) {
    return;
  }
  if (state.values.size >= 100) {
    state.truncated = true;
    return;
  }
  state.values.add(message);
}

function argbColor(value: number): string {
  const argb = value >>> 0;
  const alpha = (argb >>> 24) / 255;
  const red = (argb >>> 16) & 0xff;
  const green = (argb >>> 8) & 0xff;
  const blue = argb & 0xff;
  if (alpha >= 0.999) {
    return `#${[red, green, blue].map((part) => part.toString(16).padStart(2, "0")).join("")}`;
  }
  return `rgba(${red},${green},${blue},${alpha.toFixed(3)})`;
}

function color(resolved: ResolvedColor, warnings: WarningState): string {
  if (
    resolved.resolution === ColorResolution.DefaultTheme ||
    resolved.resolution === ColorResolution.ThemeUnparseable
  ) {
    warning(
      warnings,
      "The workbook theme is missing or unreadable; theme colors use the default Office theme in preview.",
    );
  }
  return argbColor(resolved.argb);
}

export function rectArea(rect: Rect): number {
  return (rect.lastRow - rect.firstRow + 1) * (rect.lastCol - rect.firstCol + 1);
}

export function intersects(a: Rect, b: Rect): boolean {
  return (
    a.firstRow <= b.lastRow &&
    a.lastRow >= b.firstRow &&
    a.firstCol <= b.lastCol &&
    a.lastCol >= b.firstCol
  );
}

export function clipMerge(
  merge: MergeRange,
  rect: Rect,
  warnings: WarningState,
): MergeRange | undefined {
  if (!intersects(merge, rect)) {
    return undefined;
  }
  const clipped = {
    firstRow: Math.max(merge.firstRow, rect.firstRow),
    firstCol: Math.max(merge.firstCol, rect.firstCol),
    lastRow: Math.min(merge.lastRow, rect.lastRow),
    lastCol: Math.min(merge.lastCol, rect.lastCol),
  };
  if (JSON.stringify(clipped) !== JSON.stringify(merge)) {
    warning(warnings, "A merged cell crosses the requested range and is clipped to the preview.");
  }
  return clipped;
}

export function readMergesInRange(workbook: Workbook, sheet: number, rect: Rect): MergeRange[] {
  const merges = requireListResult(workbook.getMergesInRange(sheet, rect), "read merged ranges");
  if (merges.length > MAX_PREVIEW_MERGES) {
    throw new Error(
      `preview range has ${merges.length} merged cells (limit ${MAX_PREVIEW_MERGES})`,
    );
  }
  return merges.map((merge) => ({
    firstRow: merge.firstRow,
    firstCol: merge.firstCol,
    lastRow: merge.lastRow,
    lastCol: merge.lastCol,
  }));
}

export function usedBounds(workbook: Workbook, sheet: number): Rect | undefined {
  const count = requireNumberResult(workbook.cellCount(sheet), `read sheet ${sheet} cell count`);
  if (count === 0) {
    return undefined;
  }
  if (count > MAX_CELL_SCAN_CELLS) {
    throw new Error(
      `preview used-cell inventory spans ${count} cells (limit ${MAX_CELL_SCAN_CELLS})`,
    );
  }
  let firstRow = Number.POSITIVE_INFINITY;
  let firstCol = Number.POSITIVE_INFINITY;
  let lastRow = 0;
  let lastCol = 0;
  for (let index = 0; index < count; index += 1) {
    const cell = requireCell(workbook.cellAt(sheet, index), `read sheet ${sheet} cell ${index}`);
    firstRow = Math.min(firstRow, cell.row);
    firstCol = Math.min(firstCol, cell.col);
    lastRow = Math.max(lastRow, cell.row);
    lastCol = Math.max(lastCol, cell.col);
  }
  return { firstRow, firstCol, lastRow, lastCol };
}

export function chooseRect(
  workbook: Workbook,
  sheet: number,
  range: string | undefined,
): { rect: Rect; rangeLabel: string } {
  if (range) {
    const parsed = normalizeRange(parseRangeRef(range));
    const rect = {
      firstRow: parsed.start.row,
      firstCol: parsed.start.col,
      lastRow: parsed.end.row,
      lastCol: parsed.end.col,
    };
    return {
      rect,
      rangeLabel: `${cellToA1(rect.firstRow, rect.firstCol)}:${cellToA1(rect.lastRow, rect.lastCol)}`,
    };
  }
  const pagination = workbook.paginate(sheet);
  assertStatus(pagination.status, "paginate preview sheet");
  if (pagination.printArea.length > 1) {
    throw new Error("preview range is required when the sheet has multiple print areas");
  }
  if (pagination.printArea.length === 1) {
    const area = pagination.printArea[0];
    const rect = {
      firstRow: area.firstRow,
      firstCol: area.firstCol,
      lastRow: area.lastRow,
      lastCol: area.lastCol,
    };
    return {
      rect,
      rangeLabel: `${cellToA1(rect.firstRow, rect.firstCol)}:${cellToA1(rect.lastRow, rect.lastCol)}`,
    };
  }
  const used = usedBounds(workbook, sheet);
  const sheetMerges = requireListResult(workbook.getMerges(sheet), "read merged ranges");
  const merged = sheetMerges.reduce<Rect | undefined>((bounds, merge) => {
    if (!bounds) {
      return { ...merge };
    }
    return {
      firstRow: Math.min(bounds.firstRow, merge.firstRow),
      firstCol: Math.min(bounds.firstCol, merge.firstCol),
      lastRow: Math.max(bounds.lastRow, merge.lastRow),
      lastCol: Math.max(bounds.lastCol, merge.lastCol),
    };
  }, undefined);
  if (!(used || merged)) {
    throw new Error("cannot preview an empty sheet without an explicit range");
  }
  const rect = used
    ? merged
      ? {
          firstRow: Math.min(used.firstRow, merged.firstRow),
          firstCol: Math.min(used.firstCol, merged.firstCol),
          lastRow: Math.max(used.lastRow, merged.lastRow),
          lastCol: Math.max(used.lastCol, merged.lastCol),
        }
      : used
    : merged;
  if (!rect) {
    throw new Error("cannot preview an empty sheet without an explicit range");
  }
  return {
    rect,
    rangeLabel: `${cellToA1(rect.firstRow, rect.firstCol)}:${cellToA1(rect.lastRow, rect.lastCol)}`,
  };
}

export function dimensions(
  workbook: Workbook,
  sheet: number,
  rect: Rect,
): { columns: number[]; rows: number[] } {
  const columns = Array.from({ length: rect.lastCol - rect.firstCol + 1 }, (_, offset) =>
    requireNumberResult(
      workbook.getColumnWidthPt(sheet, rect.firstCol + offset, GeometryMode.Display),
      "read column width",
    ),
  );
  const rows = Array.from({ length: rect.lastRow - rect.firstRow + 1 }, (_, offset) =>
    requireNumberResult(workbook.getRowHeightPt(sheet, rect.firstRow + offset), "read row height"),
  );
  return { columns, rows };
}

/** Reads the effective style of a cell; xf-derived results are cached by the xf the engine chose. */
export function styleFor(
  workbook: Workbook,
  sheet: number,
  row: number,
  col: number,
  metrics: ColumnMetrics,
  warnings: WarningState,
  cache: Map<number, PreviewStyle>,
): PreviewStyle {
  const effective = workbook.getEffectiveStyle(sheet, row, col);
  assertStatus(effective.status, `read cell style at ${cellToA1(row, col)}`);
  const cached = cache.get(effective.xfIndex);
  if (cached) {
    return cached;
  }
  const xf = workbook.getCellXf(effective.xfIndex);
  assertStatus(xf.status, `read cell format ${effective.xfIndex}`);
  const font = workbook.getFont(effective.fontIndex);
  assertStatus(font.status, `read font ${effective.fontIndex}`);
  const fill = workbook.getFill(effective.fillIndex);
  assertStatus(fill.status, `read fill ${effective.fillIndex}`);
  const border = workbook.getBorder(effective.borderIndex);
  assertStatus(border.status, `read border ${effective.borderIndex}`);
  if (fill.pattern > 1) {
    warning(warnings, "Pattern fills are approximated as solid fills in preview.");
  }
  if (xf.textRotation !== undefined && xf.textRotation !== 0) {
    warning(warnings, "Rotated text is not rendered in preview.");
  }
  if (xf.shrinkToFit) {
    warning(warnings, "Shrink-to-fit text is not applied in preview.");
  }
  if (border.diagonalUp || border.diagonalDown) {
    warning(warnings, "Diagonal borders are not rendered in preview.");
  }
  const side = (edge: { style: number }, resolved: ResolvedColor) => ({
    style: edge.style,
    color: color(resolved, warnings),
  });
  const style: PreviewStyle = {
    fill: fill.pattern === 0 ? undefined : color(effective.fillForeground, warnings),
    fontFamily: font.name || metrics.fontName,
    fontSize: font.size || metrics.fontSize,
    fontColor: color(effective.font, warnings),
    bold: font.bold,
    italic: font.italic,
    underline: font.underline !== 0,
    horizontal: xf.horizontalAlign,
    vertical: xf.verticalAlign,
    wrapText: xf.wrapText,
    indent: xf.indent ?? 0,
    borders: {
      left: side(border.left, effective.borders.left),
      right: side(border.right, effective.borders.right),
      top: side(border.top, effective.borders.top),
      bottom: side(border.bottom, effective.borders.bottom),
    },
  };
  cache.set(effective.xfIndex, style);
  return style;
}

/** Renders the text Excel displays for one non-empty cell. */
export function cellText(
  workbook: Workbook,
  sheet: number,
  row: number,
  col: number,
  warnings: WarningState,
): string {
  const display = workbook.getDisplayText(sheet, row, col);
  assertStatus(display.status, `read display text at ${cellToA1(row, col)}`);
  if (display.displayStatus === DisplayStatus.InvalidFormat) {
    warning(
      warnings,
      "A cell number format is invalid; preview shows its General rendering instead.",
    );
  }
  return display.text;
}

export function formulaMap(workbook: Workbook, sheet: number, rect: Rect): Map<string, string> {
  const formulas = new Map<string, string>();
  let cursor: number | null = null;
  do {
    const page = workbook.getCellsInRange(sheet, rect, cursor, FORMULA_PAGE_SIZE);
    assertStatus(page.status, `read sheet ${sheet} cells`);
    for (const cell of page.cells) {
      if (cell.formula) {
        formulas.set(`${cell.row},${cell.col}`, cell.formula);
      }
    }
    cursor = page.nextCursor;
  } while (cursor !== null);
  return formulas;
}

export function buildMetadataCells(cells: PreviewCell[]): {
  cells: PreviewCellMetadata[];
  truncated: boolean;
  omitted: number;
} {
  const rows = new Map<number, PreviewCell>();
  const columns = new Map<number, PreviewCell>();
  for (const cell of cells) {
    rows.set(cell.row, cell);
    columns.set(cell.col, cell);
  }
  const nonEmpty = cells.filter(
    (cell) =>
      (cell.text || cell.formula) &&
      (!cell.merge || (cell.row === cell.merge.firstRow && cell.col === cell.merge.firstCol)),
  );
  let truncated = nonEmpty.length > MAX_METADATA_CELLS;
  const omitted = Math.max(0, nonEmpty.length - MAX_METADATA_CELLS);
  const result = nonEmpty.slice(0, MAX_METADATA_CELLS).map((cell) => {
    const text =
      cell.text.length > MAX_METADATA_TEXT ? cell.text.slice(0, MAX_METADATA_TEXT) : cell.text;
    const formula =
      cell.formula.length > MAX_METADATA_TEXT
        ? cell.formula.slice(0, MAX_METADATA_TEXT)
        : cell.formula;
    if (text.length !== cell.text.length || formula.length !== cell.formula.length) {
      truncated = true;
    }
    let width = cell.width;
    let height = cell.height;
    if (cell.merge) {
      width = 0;
      for (let col = cell.merge.firstCol; col <= cell.merge.lastCol; col += 1) {
        width += columns.get(col)?.width ?? 0;
      }
      height = 0;
      for (let row = cell.merge.firstRow; row <= cell.merge.lastRow; row += 1) {
        height += rows.get(row)?.height ?? 0;
      }
    }
    return {
      row: cell.row,
      col: cell.col,
      a1: cellToA1(cell.row, cell.col),
      text,
      ...(formula ? { formula } : {}),
      x: cell.x,
      y: cell.y,
      width,
      height,
    };
  });
  return { cells: result, truncated, omitted };
}
