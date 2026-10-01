import type { Value } from "@libraz/formulon";
import { cellToA1, normalizeRange, parseRangeRef } from "../a1.js";
import {
  assertStatus,
  errorName,
  requireCell,
  requireNumberResult,
  type Workbook,
} from "../formulon.js";
import { formatPreviewNumber, formatPreviewText } from "./format.js";
import { type ColumnMetrics, columnCharsToPoints } from "./geometry.js";
import type { PreviewCell, PreviewCellMetadata, PreviewStyle } from "./types.js";

export const MAX_METADATA_CELLS = 500;
export const MAX_METADATA_TEXT = 256;
const MAX_FORMULA_SCAN_CELLS = 1_000_000;
const MAX_CELL_SCAN_CELLS = 1_000_000;
const MAX_PREVIEW_MERGES = 10_000;
export const DEFAULT_COLUMN_CHARS = 8.43;
export const DEFAULT_ROW_POINTS = 102 / 7;

export const BASE_WARNINGS = [
  "Preview rendering is approximate and may differ from Excel.",
  "Font substitution is possible when the requested workbook font is unavailable.",
  "Conditional formatting is not applied.",
  "Drawings, images, headers, and footers are not rendered.",
];

export type Rect = { firstRow: number; firstCol: number; lastRow: number; lastCol: number };
export type MergeRange = Rect;
type ColorSpec = { kind?: number; rgb?: number; theme?: number; indexed?: number };
export type WarningState = { values: Set<string>; truncated: boolean };
export type CellAppearance = {
  style: PreviewStyle;
  numFmtId: number;
  formatCode: string;
};

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

function color(spec: ColorSpec | undefined, fallback: number, warnings: WarningState): string {
  if (spec?.kind !== undefined && spec.kind !== 0 && spec.kind !== 1) {
    warning(warnings, "Theme, indexed, and automatic colors use an ARGB fallback in preview.");
  }
  return argbColor(spec?.kind === 1 && typeof spec.rgb === "number" ? spec.rgb : fallback);
}

function borderColor(
  spec: ColorSpec | undefined,
  fallback: number,
  warnings: WarningState,
): string {
  // A newly enabled Excel border commonly carries the engine's zero-valued
  // automatic colour selector. Zero ARGB means transparent when interpreted
  // literally, while Excel's automatic border colour is opaque black.
  return color(spec, fallback === 0 ? 0xff000000 : fallback, warnings);
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

export function readMerges(workbook: Workbook, sheet: number): MergeRange[] {
  const result = workbook.getMerges(sheet);
  assertStatus(result.status, "read merged ranges");
  if (result.length > MAX_PREVIEW_MERGES) {
    throw new Error(
      `preview sheet has ${result.length} merged cells (limit ${MAX_PREVIEW_MERGES})`,
    );
  }
  return Array.from(result).map((merge) => ({
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
  merges: MergeRange[],
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
  const merged = merges.reduce<Rect | undefined>((bounds, merge) => {
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
  metrics: ColumnMetrics,
  warnings: WarningState,
): { columns: number[]; rows: number[] } {
  const columnResult = workbook.getSheetColumns(sheet);
  assertStatus(columnResult.status, "read sheet columns");
  const rowResult = workbook.getSheetRowOverrides(sheet);
  assertStatus(rowResult.status, "read sheet rows");
  const columns = Array.from({ length: rect.lastCol - rect.firstCol + 1 }, (_, offset) => {
    const index = rect.firstCol + offset;
    const override = columnResult.columns.find(
      (entry) => entry.first <= index && index <= entry.last,
    );
    if (override?.hidden) {
      return 0;
    }
    const width =
      override && (override.hasWidth !== 0 || override.width !== 0)
        ? override.width
        : DEFAULT_COLUMN_CHARS;
    return columnCharsToPoints(width, metrics);
  });
  const rows = Array.from({ length: rect.lastRow - rect.firstRow + 1 }, (_, offset) => {
    const index = rect.firstRow + offset;
    const override = rowResult.rows.find((entry) => entry.row === index);
    if (override?.hidden) {
      return 0;
    }
    if (override && override.height === 0) {
      warning(
        warnings,
        "A row override has no explicit height; preview uses the default row height.",
      );
    }
    return override?.height || DEFAULT_ROW_POINTS;
  });
  return { columns, rows };
}

export function styleFor(
  workbook: Workbook,
  sheet: number,
  row: number,
  col: number,
  metrics: ColumnMetrics,
  warnings: WarningState,
  cache: Map<number, CellAppearance>,
): CellAppearance {
  const indexResult = workbook.getCellXfIndex(sheet, row, col);
  assertStatus(indexResult.status, `read cell style at ${cellToA1(row, col)}`);
  const cached = cache.get(indexResult.xfIndex);
  if (cached) {
    return cached;
  }
  const xf = workbook.getCellXf(indexResult.xfIndex);
  assertStatus(xf.status, `read cell format ${indexResult.xfIndex}`);
  const font = workbook.getFont(xf.fontIndex);
  assertStatus(font.status, `read font ${xf.fontIndex}`);
  const fill = workbook.getFill(xf.fillIndex);
  assertStatus(fill.status, `read fill ${xf.fillIndex}`);
  const border = workbook.getBorder(xf.borderIndex);
  assertStatus(border.status, `read border ${xf.borderIndex}`);
  const numFmt = workbook.getNumFmt(xf.numFmtId);
  assertStatus(numFmt.status, `read number format ${xf.numFmtId}`);
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
  const appearance: CellAppearance = {
    numFmtId: xf.numFmtId,
    formatCode: numFmt.formatCode,
    style: {
      fill: fill.pattern === 0 ? undefined : color(fill.fg, fill.fgArgb, warnings),
      fontFamily: font.name || metrics.fontName,
      fontSize: font.size || metrics.fontSize,
      fontColor: color(font.color, font.colorArgb, warnings),
      bold: font.bold,
      italic: font.italic,
      underline: font.underline !== 0,
      horizontal: xf.horizontalAlign,
      vertical: xf.verticalAlign,
      wrapText: xf.wrapText,
      indent: xf.indent ?? 0,
      borders: {
        left: {
          style: border.left.style,
          color: borderColor(border.left.color, border.left.colorArgb, warnings),
        },
        right: {
          style: border.right.style,
          color: borderColor(border.right.color, border.right.colorArgb, warnings),
        },
        top: {
          style: border.top.style,
          color: borderColor(border.top.color, border.top.colorArgb, warnings),
        },
        bottom: {
          style: border.bottom.style,
          color: borderColor(border.bottom.color, border.bottom.colorArgb, warnings),
        },
      },
    },
  };
  cache.set(indexResult.xfIndex, appearance);
  return appearance;
}

export function cellText(
  value: Value,
  appearance: CellAppearance,
  warnings: WarningState,
): { text: string; numeric: boolean } {
  if (value.kind === 1) {
    const formatted = formatPreviewNumber(value.number, appearance.numFmtId, appearance.formatCode);
    if (formatted.warning) {
      warning(warnings, formatted.warning);
    }
    return { text: formatted.text, numeric: true };
  }
  if (value.kind === 4) {
    return {
      text: formatPreviewText(value.kind, undefined, errorName(value.errorCode)),
      numeric: false,
    };
  }
  return {
    text: formatPreviewText(
      value.kind,
      value.kind === 2 ? Boolean(value.boolean) : value.kind === 3 ? value.text : undefined,
    ),
    numeric: false,
  };
}

export function formulaMap(workbook: Workbook, sheet: number, rect: Rect): Map<string, string> {
  const count = requireNumberResult(workbook.cellCount(sheet), `read sheet ${sheet} cell count`);
  if (count > MAX_FORMULA_SCAN_CELLS) {
    throw new Error(
      `preview formula inventory spans ${count} cells (limit ${MAX_FORMULA_SCAN_CELLS})`,
    );
  }
  const formulas = new Map<string, string>();
  for (let index = 0; index < count; index += 1) {
    const cell = requireCell(workbook.cellAt(sheet, index), `read sheet ${sheet} cell ${index}`);
    if (
      cell.formula &&
      cell.row >= rect.firstRow &&
      cell.row <= rect.lastRow &&
      cell.col >= rect.firstCol &&
      cell.col <= rect.lastCol
    ) {
      formulas.set(`${cell.row},${cell.col}`, cell.formula);
    }
  }
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
