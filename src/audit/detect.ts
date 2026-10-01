import { cellToA1 } from "../a1.js";
import type { PrimitiveCellValue } from "../session/shared.js";
import {
  type AuditCell,
  type AuditRange,
  type AuditValue,
  type Axis,
  type CellSample,
  type FindingType,
  type FormulaAuditFinding,
  type RawIssue,
  REVIEW_NOTE,
} from "./types.js";

export function key(row: number, col: number): string {
  return `${row},${col}`;
}

export function blankValue(): PrimitiveCellValue {
  return { kind: "blank" };
}

export function populated(cell: AuditCell): boolean {
  return cell.formula.length > 0 || cell.value.kind !== "blank";
}

export function truncateText(
  text: string | undefined,
  limit: number,
): { text: undefined; truncated: false } | { text: string; truncated: boolean } {
  if (text === undefined) {
    return { text: undefined, truncated: false };
  }
  return text.length > limit
    ? { text: text.slice(0, limit), truncated: true }
    : { text, truncated: false };
}

export function valueForOutput(value: PrimitiveCellValue): AuditValue {
  if (value.kind !== "text" || typeof value.value !== "string" || value.value.length <= 256) {
    return value;
  }
  return { ...value, value: value.value.slice(0, 256), textTruncated: true };
}

export function sampleFor(cell: AuditCell): CellSample {
  const formula = truncateText(cell.formula || undefined, 256);
  const pattern = truncateText(cell.pattern, 1024);
  const sample: CellSample = {
    a1: cellToA1(cell.row, cell.col),
    row: cell.row,
    col: cell.col,
    state: cell.formula ? "formula" : cell.value.kind === "blank" ? "blank" : "constant",
    value: valueForOutput(cell.value),
  };
  if (formula.text !== undefined) {
    sample.formula = formula.text;
    if (formula.truncated) {
      sample.formulaTextTruncated = true;
    }
  }
  if (pattern.text !== undefined) {
    sample.pattern = pattern.text;
    if (pattern.truncated) {
      sample.patternTextTruncated = true;
    }
  }
  return sample;
}

export function nearestSamples(cells: AuditCell[], expectedPattern: string | undefined) {
  const matching = expectedPattern
    ? cells.filter((cell) => cell.formula && cell.pattern === expectedPattern)
    : [];
  const selected = (matching.length > 0 ? matching : cells).slice(0, 3);
  return selected.map(sampleFor);
}

export function rangeArea(range: AuditRange): number {
  return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
}

export function rangeText(range: AuditRange): string {
  return `${cellToA1(range.startRow, range.startCol)}:${cellToA1(range.endRow, range.endCol)}`;
}

export function lineCells(
  map: Map<string, AuditCell>,
  range: AuditRange,
  target: AuditCell,
  axis: Axis,
  window: number,
  populatedRows: Set<number>,
  populatedCols: Set<number>,
): { before: AuditCell[]; after: AuditCell[]; considered: number } {
  const before: AuditCell[] = [];
  const after: AuditCell[] = [];
  let considered = 0;
  const isBoundary = (position: number) => {
    return axis === "vertical" ? !populatedRows.has(position) : !populatedCols.has(position);
  };
  const read = (offset: number, side: AuditCell[]) => {
    const position = axis === "vertical" ? target.row + offset : target.col + offset;
    const min = axis === "vertical" ? range.startRow : range.startCol;
    const max = axis === "vertical" ? range.endRow : range.endCol;
    if (position < min || position > max || isBoundary(position)) {
      return false;
    }
    const cell =
      map.get(axis === "vertical" ? key(position, target.col) : key(target.row, position)) ??
      (axis === "vertical"
        ? { row: position, col: target.col, formula: "", value: blankValue() }
        : { row: target.row, col: position, formula: "", value: blankValue() });
    side.push(cell);
    considered += 1;
    return true;
  };
  for (let distance = 1; distance <= window; distance += 1) {
    if (before.length < window && !read(-distance, before)) {
      break;
    }
  }
  for (let distance = 1; distance <= window; distance += 1) {
    if (after.length < window && !read(distance, after)) {
      break;
    }
  }
  return { before, after, considered };
}

export function expectedPattern(neighbors: AuditCell[], minPeers: number, considered: number) {
  const counts = new Map<string, number>();
  for (const cell of neighbors) {
    if (cell.formula && cell.patternSupported && cell.pattern !== undefined) {
      counts.set(cell.pattern, (counts.get(cell.pattern) ?? 0) + 1);
    }
  }
  const ordered = Array.from(counts.entries()).sort(
    (left, right) => right[1] - left[1] || left[0].localeCompare(right[0]),
  );
  const best = ordered[0];
  if (!best || best[1] < minPeers || best[1] * 2 <= considered) {
    return undefined;
  }
  return { pattern: best[0], count: best[1] };
}

export function issueFor(
  type: FindingType,
  direction: Axis,
  cell: AuditCell,
  before: AuditCell[],
  after: AuditCell[],
  considered: number,
  minPeers: number,
): RawIssue | undefined {
  if (type === "formula_error") {
    return {
      type,
      direction,
      cell,
      count: 0,
      considered: 0,
      before: [],
      after: [],
      confidence: 1,
    };
  }
  const neighbors = [...before, ...after];
  const expected = expectedPattern(neighbors, minPeers, considered);
  if (!expected) {
    return undefined;
  }
  const beforeMatches = before.filter(
    (peer) => peer.formula && peer.patternSupported && peer.pattern === expected.pattern,
  ).length;
  const afterMatches = after.filter(
    (peer) => peer.formula && peer.patternSupported && peer.pattern === expected.pattern,
  ).length;
  if (beforeMatches === 0 || afterMatches === 0) {
    return undefined;
  }
  if (type === "formula_pattern_outlier" && cell.pattern === expected.pattern) {
    return undefined;
  }
  if (type !== "formula_pattern_outlier" && cell.formula) {
    return undefined;
  }
  return {
    type,
    direction,
    cell,
    expectedPattern: expected.pattern,
    count: expected.count,
    considered,
    before,
    after,
    confidence: Number(Math.min(0.99, expected.count / Math.max(1, considered)).toFixed(2)),
  };
}

export function collectIssues(
  map: Map<string, AuditCell>,
  range: AuditRange,
  axis: Axis,
  window: number,
  minPeers: number,
  populatedRows: Set<number>,
  populatedCols: Set<number>,
): RawIssue[] {
  const issues: RawIssue[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const cell: AuditCell = map.get(key(row, col)) ?? {
        row,
        col,
        formula: "",
        value: blankValue(),
      };
      const neighbors = lineCells(map, range, cell, axis, window, populatedRows, populatedCols);
      if (cell.formula && cell.value.kind === "error") {
        const error = issueFor("formula_error", axis, cell, [], [], 0, minPeers);
        if (error) {
          issues.push(error);
        }
        continue;
      }
      if (cell.formula && cell.patternSupported) {
        const issue = issueFor(
          "formula_pattern_outlier",
          axis,
          cell,
          neighbors.before,
          neighbors.after,
          neighbors.considered,
          minPeers,
        );
        if (issue) {
          issues.push(issue);
        }
      } else if (!cell.formula) {
        const populatedLine = axis === "vertical" ? populatedRows.has(row) : populatedCols.has(col);
        if (populatedLine) {
          const type: FindingType =
            cell.value.kind === "blank" ? "blank_in_formula_run" : "constant_in_formula_run";
          const issue = issueFor(
            type,
            axis,
            cell,
            neighbors.before,
            neighbors.after,
            neighbors.considered,
            minPeers,
          );
          if (issue) {
            issues.push(issue);
          }
        }
      }
    }
  }
  return issues;
}

export function findingFromIssues(
  issues: RawIssue[],
  sheet: number,
  sheetNameValue: string,
): FormulaAuditFinding {
  const sorted = [...issues].sort(
    (left, right) =>
      [
        "formula_error",
        "constant_in_formula_run",
        "blank_in_formula_run",
        "formula_pattern_outlier",
      ].indexOf(left.type) -
        [
          "formula_error",
          "constant_in_formula_run",
          "blank_in_formula_run",
          "formula_pattern_outlier",
        ].indexOf(right.type) || left.direction.localeCompare(right.direction),
  );
  const primary = sorted[0];
  const formula = truncateText(primary.cell.formula || undefined, 256);
  const pattern = truncateText(primary.cell.pattern, 1024);
  const expected = truncateText(primary.expectedPattern, 1024);
  const directions = (["vertical", "horizontal"] as const).filter((direction) =>
    sorted.some((issue) => issue.direction === direction),
  );
  const finding: FormulaAuditFinding = {
    type: primary.type,
    direction: directions.length === 1 ? directions[0] : "both",
    directions,
    sheet,
    sheetName: sheetNameValue,
    row: primary.cell.row,
    col: primary.cell.col,
    a1: cellToA1(primary.cell.row, primary.cell.col),
    ref: `${sheetNameValue}!${cellToA1(primary.cell.row, primary.cell.col)}`,
    value: valueForOutput(primary.cell.value),
    count: primary.count,
    considered: primary.considered,
    nearestSamples: {
      before: nearestSamples(primary.before, primary.expectedPattern),
      after: nearestSamples(primary.after, primary.expectedPattern),
    },
    confidence: primary.confidence,
    issues: sorted.map((issue) => {
      const expected = truncateText(issue.expectedPattern, 1024);
      return {
        type: issue.type,
        direction: issue.direction,
        expectedPattern: expected.text,
        ...(expected.truncated ? { expectedPatternTextTruncated: true } : {}),
        count: issue.count,
        considered: issue.considered,
        confidence: issue.confidence,
      };
    }),
    reviewNote: REVIEW_NOTE,
  };
  if (formula.text !== undefined) {
    finding.formula = formula.text;
    if (formula.truncated) {
      finding.formulaTextTruncated = true;
    }
  }
  if (pattern.text !== undefined) {
    finding.pattern = pattern.text;
    if (pattern.truncated) {
      finding.patternTextTruncated = true;
    }
  }
  if (expected.text !== undefined) {
    finding.expectedPattern = expected.text;
    if (expected.truncated) {
      finding.expectedPatternTextTruncated = true;
    }
  }
  return finding;
}
