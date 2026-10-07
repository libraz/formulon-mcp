import { colToIndex, normalizeRange, parseRangeRef } from "../a1.js";
import { findSheetIndex, type Workbook } from "../formulon.js";
import { mmToPoints, pixelsToPoints } from "../preview/geometry.js";
import type { PrintSettingsInput } from "../print.js";
import { columnPointsToChars } from "../session/geometry.js";
import { assertInGrid, MAX_COL, MAX_ROW } from "../session/shared.js";
import type { StyleInput } from "../styles.js";
import { nonEmptyPrintSettingsSchema, styleSchema } from "./schemas.js";
import {
  type LayoutOperation,
  type LayoutRect,
  MAX_EXPANDED_ROW_WRITES,
  MAX_LAYOUT_COLUMN_WIDTH_CHARS,
  MAX_LAYOUT_OPERATIONS,
  MAX_LAYOUT_ROW_HEIGHT_PT,
  MAX_STYLED_CELLS,
  type PreparedLayoutOperation,
  type ResolvedLayoutRange,
} from "./types.js";

function finiteNonnegative(value: unknown, name: string): asserts value is number {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    throw new Error(`${name} must be a finite nonnegative number: ${String(value)}`);
  }
}

function integerInGrid(value: unknown, name: string, max: number): asserts value is number {
  if (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value > max) {
    throw new Error(`${name} must be an integer in 0..${max}: ${String(value)}`);
  }
}

function rangeCellCount(rect: LayoutRect): number {
  return (rect.lastRow - rect.firstRow + 1) * (rect.lastCol - rect.firstCol + 1);
}

function resolvedRange(
  workbook: Workbook,
  selectedSheet: number,
  ref: string,
  label: string,
): ResolvedLayoutRange {
  if (typeof ref !== "string" || ref.trim().length === 0) {
    throw new Error(`${label} must be a non-empty A1 range`);
  }
  const parsed = normalizeRange(parseRangeRef(ref));
  const rangeSheet = findSheetIndex(workbook, parsed.sheetName ?? selectedSheet);
  if (rangeSheet !== selectedSheet) {
    throw new Error(`${label} must refer to the selected sheet; got ${ref}`);
  }
  assertInGrid(parsed.start.row, parsed.start.col);
  assertInGrid(parsed.end.row, parsed.end.col);
  return {
    ref,
    rect: {
      firstRow: parsed.start.row,
      firstCol: parsed.start.col,
      lastRow: parsed.end.row,
      lastCol: parsed.end.col,
    },
  };
}

function columnStoredSize(
  workbook: Workbook,
  selectedSheet: number,
  size: number,
  unit: "chars" | "pt" | "px" | "mm" | undefined,
): number {
  const resolvedUnit = unit ?? "chars";
  const points =
    resolvedUnit === "chars"
      ? undefined
      : resolvedUnit === "pt"
        ? size
        : resolvedUnit === "px"
          ? pixelsToPoints(size)
          : resolvedUnit === "mm"
            ? mmToPoints(size)
            : undefined;
  const stored = points === undefined ? size : columnPointsToChars(workbook, selectedSheet, points);
  if (stored > MAX_LAYOUT_COLUMN_WIDTH_CHARS) {
    throw new Error(
      `column width exceeds the ${MAX_LAYOUT_COLUMN_WIDTH_CHARS}-character Excel limit: ${stored}`,
    );
  }
  return stored;
}

function rowStoredSize(size: number, unit: "pt" | "px" | "mm" | undefined): number {
  const resolvedUnit = unit ?? "pt";
  const stored =
    resolvedUnit === "pt" ? size : resolvedUnit === "px" ? pixelsToPoints(size) : mmToPoints(size);
  if (stored > MAX_LAYOUT_ROW_HEIGHT_PT) {
    throw new Error(`row height exceeds the ${MAX_LAYOUT_ROW_HEIGHT_PT}pt Excel limit: ${stored}`);
  }
  return stored;
}

function printAreaRanges(workbook: Workbook, selectedSheet: number, printArea: string): void {
  if (printArea.trim() === "") {
    return;
  }
  for (const [index, ref] of printArea.split(",").entries()) {
    resolvedRange(workbook, selectedSheet, ref.trim(), `printArea[${index}]`);
  }
}

function validatePrintTitles(settings: PrintSettingsInput): void {
  const titles = settings.printTitles;
  if (!titles) {
    return;
  }
  if (titles.repeatRows !== undefined && titles.repeatRows !== "") {
    const match = /^(\d+):(\d+)$/.exec(titles.repeatRows.trim());
    if (
      !match ||
      Number(match[1]) < 1 ||
      Number(match[2]) < Number(match[1]) ||
      Number(match[2]) > MAX_ROW + 1
    ) {
      throw new Error(
        `printTitles.repeatRows must be a whole-row span such as 1:2: ${titles.repeatRows}`,
      );
    }
  }
  if (titles.repeatCols !== undefined && titles.repeatCols !== "") {
    const match = /^([A-Za-z]+):([A-Za-z]+)$/.exec(titles.repeatCols.trim());
    if (!match) {
      throw new Error(
        `printTitles.repeatCols must be a whole-column span such as A:A: ${titles.repeatCols}`,
      );
    }
    const first = colToIndex(match[1]);
    const last = colToIndex(match[2]);
    if (first < 0 || last < first || last > MAX_COL) {
      throw new Error(`printTitles.repeatCols exceeds Excel sheet limits: ${titles.repeatCols}`);
    }
  }
}

function preparedPrintSettings(
  workbook: Workbook,
  selectedSheet: number,
  settings: PrintSettingsInput,
): PrintSettingsInput {
  const parsed = nonEmptyPrintSettingsSchema.parse(settings) as PrintSettingsInput;
  if (parsed.printTitles) {
    parsed.printTitles = {
      ...parsed.printTitles,
      ...(parsed.printTitles.repeatRows !== undefined
        ? { repeatRows: parsed.printTitles.repeatRows.trim() }
        : {}),
      ...(parsed.printTitles.repeatCols !== undefined
        ? { repeatCols: parsed.printTitles.repeatCols.trim() }
        : {}),
    };
  }
  if (parsed.printArea !== undefined) {
    printAreaRanges(workbook, selectedSheet, parsed.printArea);
  }
  validatePrintTitles(parsed);
  for (const row of parsed.rowBreaks ?? []) {
    integerInGrid(row, "rowBreak", MAX_ROW);
  }
  for (const col of parsed.colBreaks ?? []) {
    integerInGrid(col, "colBreak", MAX_COL);
  }
  return parsed;
}

/** Validates every batch operation and computes native stored dimensions. */
export function preflightLayout(
  workbook: Workbook,
  selectedSheet: number,
  operations: readonly LayoutOperation[],
): PreparedLayoutOperation[] {
  if (operations.length === 0) {
    throw new Error("operations must contain at least one operation");
  }
  if (operations.length > MAX_LAYOUT_OPERATIONS) {
    throw new Error(`operations exceeds the ${MAX_LAYOUT_OPERATIONS}-operation limit`);
  }
  let expandedRows = 0;
  let styledCells = 0;
  return operations.map((operation, index) => {
    if (!operation || typeof operation !== "object" || typeof operation.type !== "string") {
      throw new Error(`operations[${index}] must have a recognized type`);
    }
    if (operation.type === "column") {
      if (operation.unit !== undefined && !["chars", "pt", "px", "mm"].includes(operation.unit)) {
        throw new Error(`operations[${index}].unit must be chars, pt, px, or mm`);
      }
      integerInGrid(operation.first, `operations[${index}].first`, MAX_COL);
      const last = operation.last ?? operation.first;
      integerInGrid(last, `operations[${index}].last`, MAX_COL);
      if (last < operation.first) {
        throw new Error(`operations[${index}] column span must be ordered`);
      }
      finiteNonnegative(operation.size, `operations[${index}].size`);
      const storedSize = columnStoredSize(workbook, selectedSheet, operation.size, operation.unit);
      return { ...operation, last, storedSize };
    }
    if (operation.type === "row") {
      if (operation.unit !== undefined && !["pt", "px", "mm"].includes(operation.unit)) {
        throw new Error(`operations[${index}].unit must be pt, px, or mm`);
      }
      integerInGrid(operation.first, `operations[${index}].first`, MAX_ROW);
      const last = operation.last ?? operation.first;
      integerInGrid(last, `operations[${index}].last`, MAX_ROW);
      if (last < operation.first) {
        throw new Error(`operations[${index}] row span must be ordered`);
      }
      finiteNonnegative(operation.size, `operations[${index}].size`);
      expandedRows += last - operation.first + 1;
      if (expandedRows > MAX_EXPANDED_ROW_WRITES) {
        throw new Error(`row operations exceed the ${MAX_EXPANDED_ROW_WRITES}-row write limit`);
      }
      return { ...operation, last, storedSize: rowStoredSize(operation.size, operation.unit) };
    }
    if (operation.type === "merge") {
      if (operation.operation !== "add" && operation.operation !== "remove") {
        throw new Error(`operations[${index}].operation must be add or remove`);
      }
      const resolved = resolvedRange(
        workbook,
        selectedSheet,
        operation.range,
        `operations[${index}].range`,
      );
      return { ...operation, resolved };
    }
    if (operation.type === "style") {
      const style = styleSchema.parse(operation.style) as StyleInput;
      const base = operation.base ?? "existing";
      if (base !== "existing" && base !== "default") {
        throw new Error(`operations[${index}].base must be existing or default`);
      }
      const resolved = resolvedRange(
        workbook,
        selectedSheet,
        operation.range,
        `operations[${index}].range`,
      );
      styledCells += rangeCellCount(resolved.rect);
      if (styledCells > MAX_STYLED_CELLS) {
        throw new Error(`style operations exceed the ${MAX_STYLED_CELLS}-cell limit`);
      }
      return { ...operation, style, resolved, base };
    }
    if (operation.type === "print") {
      const settings = preparedPrintSettings(workbook, selectedSheet, operation.settings);
      return { ...operation, settings };
    }
    throw new Error(`operations[${index}] has unknown type`);
  });
}
