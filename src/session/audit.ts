import type { Value } from "@libraz/formulon";
import { cellToA1, normalizeRange, parseRangeRef } from "../a1.js";
import {
  blankValue,
  collectIssues,
  findingFromIssues,
  key,
  populated,
  rangeArea,
  rangeText,
} from "../audit/detect.js";
import type {
  AuditCell,
  AuditRange,
  Axis,
  FormulaAuditOptions,
  FormulaAuditResult,
  RawIssue,
} from "../audit/types.js";
import {
  assertStatus,
  findSheetIndex,
  requireCell,
  requireNumberResult,
  valueToJson,
} from "../formulon.js";
import { assertInGrid, type PrimitiveCellValue, rectCells, sheetName } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

const MAX_AUDIT_INVENTORY_CELLS = 1_000_000;

export type {
  FormulaAuditDirection,
  FormulaAuditFinding,
  FormulaAuditOptions,
  FormulaAuditResult,
} from "../audit/types.js";

function validateOptions(options: FormulaAuditOptions) {
  const direction = options.direction ?? "vertical";
  if (direction !== "vertical" && direction !== "horizontal" && direction !== "both") {
    throw new Error(`direction must be vertical, horizontal, or both: ${direction}`);
  }
  const window = options.window ?? 5;
  if (!Number.isInteger(window) || window < 2 || window > 50) {
    throw new Error(`window must be an integer in 2..50: ${window}`);
  }
  const minPeers = options.minPeers ?? 3;
  if (!Number.isInteger(minPeers) || minPeers < 3 || minPeers > 20) {
    throw new Error(`minPeers must be an integer in 3..20: ${minPeers}`);
  }
  if (minPeers > 2 * window) {
    throw new Error(`minPeers must be no greater than twice window: ${minPeers} > ${2 * window}`);
  }
  const maxCells = options.maxCells ?? 50_000;
  if (!Number.isInteger(maxCells) || maxCells < 1 || maxCells > 50_000) {
    throw new Error(`maxCells must be an integer in 1..50000: ${maxCells}`);
  }
  const maxFindings = options.maxFindings ?? 100;
  if (!Number.isInteger(maxFindings) || maxFindings < 1 || maxFindings > 500) {
    throw new Error(`maxFindings must be an integer in 1..500: ${maxFindings}`);
  }
  return { direction, window, minPeers, maxCells, maxFindings };
}

function resolveSheetAndRange(
  session: ReturnType<typeof getSession>,
  options: FormulaAuditOptions,
): { sheet: number; range: AuditRange | undefined } {
  if (options.range !== undefined) {
    const parsed = normalizeRange(parseRangeRef(options.range));
    const qualifiedSheet =
      parsed.sheetName === undefined
        ? undefined
        : findSheetIndex(session.workbook, parsed.sheetName);
    const selectedSheet =
      options.sheet === undefined ? undefined : findSheetIndex(session.workbook, options.sheet);
    if (
      qualifiedSheet !== undefined &&
      selectedSheet !== undefined &&
      qualifiedSheet !== selectedSheet
    ) {
      throw new Error("sheet conflicts with the qualified range");
    }
    const sheet = qualifiedSheet ?? selectedSheet ?? findSheetIndex(session.workbook, undefined);
    assertInGrid(parsed.start.row, parsed.start.col);
    assertInGrid(parsed.end.row, parsed.end.col);
    return {
      sheet,
      range: {
        startRow: parsed.start.row,
        startCol: parsed.start.col,
        endRow: parsed.end.row,
        endCol: parsed.end.col,
      },
    };
  }
  return { sheet: findSheetIndex(session.workbook, options.sheet), range: undefined };
}

function readStoredCells(
  session: ReturnType<typeof getSession>,
  sheet: number,
  range?: AuditRange,
): AuditCell[] {
  const toCell = (cell: { row: number; col: number; formula: string; value: Value }) => ({
    row: cell.row,
    col: cell.col,
    formula: cell.formula,
    value: valueToJson(cell.value) as PrimitiveCellValue,
  });
  if (range) {
    const rect = {
      firstRow: range.startRow,
      lastRow: range.endRow,
      firstCol: range.startCol,
      lastCol: range.endCol,
    };
    return Array.from(rectCells(session, sheet, rect), toCell);
  }
  const count = requireNumberResult(
    session.workbook.cellCount(sheet),
    `read sheet ${sheet} cell count`,
  );
  if (count > MAX_AUDIT_INVENTORY_CELLS) {
    throw new Error(
      `formula audit cannot inspect sheet ${sheet}: ${count} stored cells exceeds the inventory limit of ${MAX_AUDIT_INVENTORY_CELLS}; pass an explicit range`,
    );
  }
  const cells: AuditCell[] = [];
  for (let index = 0; index < count; index += 1) {
    cells.push(
      toCell(
        requireCell(session.workbook.cellAt(sheet, index), `read sheet ${sheet} cell ${index}`),
      ),
    );
  }
  return cells;
}

function usedRange(cells: AuditCell[]): AuditRange | undefined {
  if (cells.length === 0) {
    return undefined;
  }
  let startRow = cells[0].row;
  let startCol = cells[0].col;
  let endRow = cells[0].row;
  let endCol = cells[0].col;
  for (const cell of cells.slice(1)) {
    startRow = Math.min(startRow, cell.row);
    startCol = Math.min(startCol, cell.col);
    endRow = Math.max(endRow, cell.row);
    endCol = Math.max(endCol, cell.col);
  }
  return { startRow, startCol, endRow, endCol };
}

function refreshValues(
  session: ReturnType<typeof getSession>,
  sheet: number,
  cells: AuditCell[],
  range: AuditRange,
  recalc: boolean,
): Map<string, AuditCell> {
  const byKey = new Map(
    cells
      .filter(
        (cell) =>
          cell.row >= range.startRow &&
          cell.row <= range.endRow &&
          cell.col >= range.startCol &&
          cell.col <= range.endCol,
      )
      .map((cell) => [key(cell.row, cell.col), cell]),
  );
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const cell = byKey.get(key(row, col));
      if (recalc) {
        const result = session.workbook.getValue(sheet, row, col);
        assertStatus(result.status, `read audit cell ${cellToA1(row, col)}`);
        if (cell) {
          cell.value = valueToJson(result.value) as PrimitiveCellValue;
        } else {
          byKey.set(key(row, col), {
            row,
            col,
            formula: "",
            value: valueToJson(result.value) as PrimitiveCellValue,
          });
        }
      } else if (!cell) {
        byKey.set(key(row, col), { row, col, formula: "", value: blankValue() });
      }
    }
  }
  return byKey;
}

/** Upper-cases a formula outside its double-quoted string literals, which stay case-significant. */
function foldCase(formula: string): string {
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, index) => (index % 2 === 1 ? part : part.toUpperCase()))
    .join("");
}

/**
 * Keys each formula by the engine's R1C1 text relative to its own cell, so
 * copied formulas compare equal. A formula the engine cannot render in R1C1 is
 * kept out of pattern comparisons.
 */
function formulaPatterns(
  session: ReturnType<typeof getSession>,
  sheet: number,
  cells: Map<string, AuditCell>,
) {
  for (const cell of cells.values()) {
    if (!cell.formula) {
      continue;
    }
    const result = session.workbook.getFormulaR1C1(sheet, cell.row, cell.col);
    const supported = result.status.ok && result.formula !== null;
    cell.pattern = supported ? foldCase(result.formula as string) : cell.formula;
    cell.patternSupported = supported;
  }
}

export function auditSessionFormulas(
  id: string,
  options: FormulaAuditOptions = {},
): FormulaAuditResult {
  const session = getSession(id);
  const validated = validateOptions(options);
  const target = resolveSheetAndRange(session, options);
  if (target.range && rangeArea(target.range) > validated.maxCells) {
    throw new Error(
      `audit range ${rangeText(target.range)} spans ${rangeArea(target.range)} cells (limit ${validated.maxCells}); narrow the range or raise maxCells`,
    );
  }
  const stored = readStoredCells(session, target.sheet, target.range);
  const range = target.range ?? usedRange(stored);
  const sheetNameValue = sheetName(session, target.sheet);

  if (!range) {
    touch(session, false);
    return {
      complete: true,
      session: publicInfo(session),
      sheet: target.sheet,
      sheetName: sheetNameValue,
      range: null,
      direction: validated.direction,
      window: validated.window,
      minPeers: validated.minPeers,
      scannedCells: 0,
      formulaCells: 0,
      unsupportedFormulaCells: 0,
      findingCount: 0,
      findings: [],
      findingsTruncated: false,
      warnings: [],
    };
  }
  const area = rangeArea(range);
  if (area > validated.maxCells) {
    throw new Error(
      `audit range ${rangeText(range)} spans ${area} cells (limit ${validated.maxCells}); narrow the range or raise maxCells`,
    );
  }

  if (options.recalc) {
    // Recalculation changes cached formula values and may partially mutate
    // native state before reporting a failure; record dirty before invoking it.
    touch(session, true);
    assertStatus(session.workbook.recalc(), "recalc workbook for formula audit");
  }

  const cells = refreshValues(session, target.sheet, stored, range, Boolean(options.recalc));
  formulaPatterns(session, target.sheet, cells);

  const populatedRows = new Set<number>();
  const populatedCols = new Set<number>();
  let formulaCells = 0;
  let unsupportedFormulaCells = 0;
  for (const cell of cells.values()) {
    if (populated(cell)) {
      populatedRows.add(cell.row);
      populatedCols.add(cell.col);
    }
    if (cell.formula) {
      formulaCells += 1;
      if (cell.patternSupported === false) {
        unsupportedFormulaCells += 1;
      }
    }
  }

  const axes: Axis[] =
    validated.direction === "both" ? ["vertical", "horizontal"] : [validated.direction];
  const rawIssues = axes.flatMap((axis) =>
    collectIssues(
      cells,
      range,
      axis,
      validated.window,
      validated.minPeers,
      populatedRows,
      populatedCols,
    ),
  );
  const grouped = new Map<string, RawIssue[]>();
  for (const issue of rawIssues) {
    const entry = grouped.get(key(issue.cell.row, issue.cell.col)) ?? [];
    entry.push(issue);
    grouped.set(key(issue.cell.row, issue.cell.col), entry);
  }
  const sortedGroups = Array.from(grouped.values()).sort((left, right) => {
    const leftIssue = left[0];
    const rightIssue = right[0];
    return (
      leftIssue.cell.row - rightIssue.cell.row ||
      leftIssue.cell.col - rightIssue.cell.col ||
      leftIssue.type.localeCompare(rightIssue.type) ||
      leftIssue.direction.localeCompare(rightIssue.direction)
    );
  });
  const findingsTruncated = sortedGroups.length > validated.maxFindings;
  const findings = sortedGroups
    .slice(0, validated.maxFindings)
    .map((issues) => findingFromIssues(issues, target.sheet, sheetNameValue));
  touch(session, false);
  return {
    complete: true,
    session: publicInfo(session),
    sheet: target.sheet,
    sheetName: sheetNameValue,
    range: rangeText(range),
    direction: validated.direction,
    window: validated.window,
    minPeers: validated.minPeers,
    scannedCells: area,
    formulaCells,
    unsupportedFormulaCells,
    findingCount: sortedGroups.length,
    findings,
    findingsTruncated,
    warnings:
      unsupportedFormulaCells > 0
        ? [
            `${unsupportedFormulaCells} formula(s) could not be rendered in R1C1 and were excluded from pattern comparisons`,
          ]
        : [],
  };
}
