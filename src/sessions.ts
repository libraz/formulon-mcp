export type {
  PreviewMetadata,
  PreviewOptions,
  PreviewResult,
} from "./preview/types.js";
export type { LayoutOptions } from "./session/analysis.js";
export {
  analyzeSessionWorkbook,
  detectSessionRegions,
  inspectSessionLayout,
} from "./session/analysis.js";
export type {
  FormulaAuditDirection,
  FormulaAuditFinding,
  FormulaAuditOptions,
  FormulaAuditResult,
} from "./session/audit.js";
export { auditSessionFormulas } from "./session/audit.js";
export type {
  FlexibleCellMutation,
  RangeReadOptions,
  RangeWriteCell,
} from "./session/cells.js";
export {
  applySessionMutations,
  getSessionCell,
  getSessionCellByA1,
  getSessionRange,
  setSessionRange,
} from "./session/cells.js";
export type { LayoutOperation } from "./session/layout.js";
export { applySessionLayout } from "./session/layout.js";
export type {
  CreatePivotOptions,
  PivotAggregationName,
  PivotGrandTotals,
  PivotValueOption,
} from "./session/pivots.js";
export { createSessionPivot } from "./session/pivots.js";
export {
  buildSessionDocument,
  sessionDefaultFont,
  sessionPrintSettings,
  styleSessionRange,
} from "./session/presentation.js";
export { previewSessionRange } from "./session/preview.js";
export type {
  ReplaceOptions,
  SearchOptions,
  SearchTarget,
} from "./session/search.js";
export {
  findSessionCells,
  replaceSessionCells,
} from "./session/search.js";
export type { SessionInfo, WorkbookSession } from "./session/state.js";
export {
  closeSession,
  getSession,
  inspectSession,
  listSessions,
  openSession,
  saveSession,
} from "./session/state.js";
export type {
  TableOperation,
  TableOperationOptions,
} from "./session/tables.js";
export { listSessionTables, tableOperation } from "./session/tables.js";
export type {
  DimensionAxis,
  DimensionOperation,
  SheetVisibilityName,
} from "./session/workbook.js";
export {
  applySheetOperation,
  editSessionStructure,
  getSessionMetadata,
  recalcSession,
  resolveSessionSheet,
  sessionDimension,
  setSessionDefinedName,
  setSessionSheetView,
  traceSession,
} from "./session/workbook.js";
export {
  callWorkbookMethod,
  WORKBOOK_METHODS,
} from "./session/workbook-call.js";
