import type { PrimitiveCellValue } from "../session/shared.js";
import type { SessionInfo } from "../session/state.js";

export type AuditValue = PrimitiveCellValue & { textTruncated?: boolean };

export type FormulaAuditDirection = "vertical" | "horizontal" | "both";

export type FormulaAuditOptions = {
  sheet?: number | string;
  range?: string;
  direction?: FormulaAuditDirection;
  window?: number;
  minPeers?: number;
  maxCells?: number;
  maxFindings?: number;
  recalc?: boolean;
};

export type Axis = "vertical" | "horizontal";
export type FindingType =
  | "constant_in_formula_run"
  | "blank_in_formula_run"
  | "formula_pattern_outlier"
  | "formula_error";

export type AuditRange = {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

export type AuditCell = {
  row: number;
  col: number;
  formula: string;
  value: AuditValue;
  pattern?: string;
  patternSupported?: boolean;
};

export type CellSample = {
  a1: string;
  row: number;
  col: number;
  state: "formula" | "constant" | "blank";
  formula?: string;
  formulaTextTruncated?: boolean;
  value: AuditValue;
  pattern?: string;
  patternTextTruncated?: boolean;
};

export type RawIssue = {
  type: FindingType;
  direction: Axis;
  cell: AuditCell;
  expectedPattern?: string;
  count: number;
  considered: number;
  before: AuditCell[];
  after: AuditCell[];
  confidence: number;
};

export type FormulaAuditFinding = {
  type: FindingType;
  direction: Axis | "both";
  directions: Axis[];
  sheet: number;
  sheetName: string;
  row: number;
  col: number;
  a1: string;
  ref: string;
  formula?: string;
  formulaTextTruncated?: boolean;
  value: AuditValue;
  pattern?: string;
  patternTextTruncated?: boolean;
  expectedPattern?: string;
  expectedPatternTextTruncated?: boolean;
  count: number;
  considered: number;
  nearestSamples: {
    before: CellSample[];
    after: CellSample[];
  };
  confidence: number;
  issues: Array<{
    type: FindingType;
    direction: Axis;
    expectedPattern?: string;
    expectedPatternTextTruncated?: boolean;
    count: number;
    considered: number;
    confidence: number;
  }>;
  reviewNote: string;
};

export type FormulaAuditResult = {
  complete: true;
  session: SessionInfo;
  sheet: number;
  sheetName: string;
  range: string | null;
  direction: FormulaAuditDirection;
  window: number;
  minPeers: number;
  scannedCells: number;
  formulaCells: number;
  unsupportedFormulaCells: number;
  findingCount: number;
  findings: FormulaAuditFinding[];
  findingsTruncated: boolean;
  warnings: string[];
};

export const REVIEW_NOTE =
  "Review only: no automatic repair is performed, and this audit does not establish formula correctness.";
