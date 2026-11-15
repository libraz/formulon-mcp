import { HorizontalAlign, type ResolvedColor, VerticalAlign } from "@libraz/formulon";
import { cellToA1, normalizeRange, parseRangeRef } from "../a1.js";
import {
  ColorResolution,
  DisplayStatus,
  GeometryMode,
  StatusCode,
  statusError,
  type Workbook,
} from "../formulon.js";
import type { ColumnMetrics } from "./geometry.js";
import type { PreviewCell, PreviewCellMetadata, PreviewStyle } from "./types.js";

export const MAX_METADATA_CELLS = 500;
export const MAX_METADATA_TEXT = 256;
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
  const merges = workbook.sheet(sheet).getMergesInRange(rect);
  if (merges.length > MAX_PREVIEW_MERGES) {
    throw statusError(
      StatusCode.InvalidArgument,
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
  const used = workbook.sheet(sheet).getUsedRange();
  return used === null ? undefined : { ...used };
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
  const pagination = workbook.sheet(sheet).getPagination();
  if (pagination.printArea.length > 1) {
    throw statusError(
      StatusCode.InvalidArgument,
      "preview range is required when the sheet has multiple print areas",
    );
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
  const sheetMerges = workbook.sheet(sheet).listMerges();
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
    throw statusError(
      StatusCode.InvalidArgument,
      "cannot preview an empty sheet without an explicit range",
    );
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
    throw statusError(
      StatusCode.InvalidArgument,
      "cannot preview an empty sheet without an explicit range",
    );
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
  const handle = workbook.sheet(sheet);
  const columns = Array.from({ length: rect.lastCol - rect.firstCol + 1 }, (_, offset) =>
    handle.getColumnWidthPt(rect.firstCol + offset, GeometryMode.Display),
  );
  const rows = Array.from({ length: rect.lastRow - rect.firstRow + 1 }, (_, offset) =>
    handle.getRowHeightPt(rect.firstRow + offset),
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
  const effective = workbook.sheet(sheet).ooxml.getEffectiveStyle(row, col);
  const cached = cache.get(effective.xfIndex);
  if (cached) {
    return cached;
  }
  const { ooxml } = workbook;
  const xf = ooxml.getCellXf(effective.xfIndex);
  const font = ooxml.getFont(effective.fontIndex);
  const fill = ooxml.getFill(effective.fillIndex);
  const border = ooxml.getBorder(effective.borderIndex);
  if (fill.pattern > 1) {
    warning(warnings, "Pattern fills are approximated as solid fills in preview.");
  }
  if (xf.textRotation !== null && xf.textRotation !== 0) {
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
    fontSize: font.sizePt || metrics.fontSize,
    fontColor: color(effective.fontColor, warnings),
    bold: font.bold === true,
    italic: font.italic === true,
    underline: font.underline !== 0,
    horizontal: xf.horizontal ?? HorizontalAlign.General,
    vertical: xf.vertical ?? VerticalAlign.Bottom,
    wrapText: xf.wrapText === true,
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
  const display = workbook.sheet(sheet).getDisplayText(row, col);
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
  const handle = workbook.sheet(sheet);
  let cursor: number | null = null;
  do {
    const page = handle.getCellsInRange(rect, { cursor, limit: FORMULA_PAGE_SIZE });
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
