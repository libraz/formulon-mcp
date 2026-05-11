import { randomUUID } from "node:crypto";
import { cellToA1, normalizeRange, parseCellRef, parseRangeRef } from "./a1.js";
import {
  applyMutation,
  assertStatus,
  type CellMutation,
  createOrLoadWorkbook,
  findSheetIndex,
  normalizeFormula,
  resultToJson,
  saveWorkbook,
  statusToJson,
  valueToJson,
  type Workbook,
  workbookSummary,
} from "./formulon.js";

export type SessionInfo = {
  id: string;
  sourcePath?: string;
  outputPath?: string;
  createdAt: string;
  updatedAt: string;
  dirty: boolean;
};

type WorkbookSession = SessionInfo & {
  workbook: Workbook;
};

type A1MutationBase = {
  sheet?: number | string;
  a1?: string;
  row?: number;
  col?: number;
};

export type FlexibleCellMutation =
  | (A1MutationBase & { type: "number"; value: number })
  | (A1MutationBase & { type: "bool"; value: boolean })
  | (A1MutationBase & { type: "text"; value: string })
  | (A1MutationBase & { type: "blank" })
  | (A1MutationBase & { type: "formula"; formula: string });

export type SearchTarget = "texts" | "formulas" | "both";

export type SearchOptions = {
  sheet?: number | string;
  target: SearchTarget;
  matchCase: boolean;
  wholeCell: boolean;
  regex: boolean;
  maxResults: number;
};

export type ReplaceOptions = SearchOptions & {
  replacement: string;
  maxReplacements: number;
  recalc: boolean;
};

export type LayoutOptions = {
  sheet?: number | string;
  includeCells: boolean;
  includeStyles: boolean;
  maxCells: number;
};

type PrimitiveCellValue = {
  kind: string;
  value?: number | boolean | string;
  errorCode?: number;
};

type CellSearchResult = {
  sheet: number;
  sheetName: string;
  row: number;
  col: number;
  a1: string;
  ref: string;
  target: "text" | "formula";
  text: string;
};

type CellReplaceResult = CellSearchResult & {
  before: string;
  after: string;
  status: unknown;
};

const sessions = new Map<string, WorkbookSession>();

const WORKBOOK_METHODS = new Set([
  "save",
  "addSheet",
  "removeSheet",
  "renameSheet",
  "moveSheet",
  "sheetCount",
  "sheetName",
  "setNumber",
  "setBool",
  "setText",
  "setBlank",
  "setFormula",
  "getValue",
  "getLambdaText",
  "recalc",
  "partialRecalc",
  "setIterative",
  "calcMode",
  "setCalcMode",
  "excelProfileId",
  "setExcelProfileId",
  "insertRows",
  "deleteRows",
  "insertCols",
  "deleteCols",
  "cellCount",
  "cellAt",
  "definedNameCount",
  "definedNameAt",
  "setDefinedName",
  "tableCount",
  "tableAt",
  "passthroughCount",
  "passthroughAt",
  "pivotCount",
  "pivotLayout",
  "pivotCacheCount",
  "pivotCacheIdAt",
  "pivotCacheCreate",
  "pivotCacheRemove",
  "pivotCacheFieldCount",
  "pivotCacheFieldName",
  "pivotCacheFieldAdd",
  "pivotCacheFieldClear",
  "pivotCacheFieldSharedItemCount",
  "pivotCacheFieldAddSharedItemNumber",
  "pivotCacheFieldAddSharedItemText",
  "pivotCacheFieldAddSharedItemBool",
  "pivotCacheFieldAddSharedItemBlank",
  "pivotCacheFieldClearSharedItems",
  "pivotCacheRecordCount",
  "pivotCacheRecordAdd",
  "pivotCacheRecordClear",
  "pivotCacheRecordSetNumber",
  "pivotCacheRecordSetText",
  "pivotCacheRecordSetBool",
  "pivotCacheRecordSetBlank",
  "pivotCacheRecordSetError",
  "pivotCreate",
  "pivotRemove",
  "pivotSetName",
  "pivotSetAnchor",
  "pivotSetGrandTotals",
  "pivotFieldCount",
  "pivotFieldAdd",
  "pivotFieldClear",
  "pivotFieldSetAxis",
  "pivotFieldSetSort",
  "pivotFieldSetSubtotalTop",
  "pivotFieldAddAggregation",
  "pivotFieldClearAggregations",
  "pivotFieldAddItem",
  "pivotFieldClearItems",
  "pivotFieldSetItemVisible",
  "pivotFieldAddSubtotalFn",
  "pivotFieldClearSubtotalFns",
  "pivotFieldSetDateGroup",
  "pivotFieldClearDateGroup",
  "pivotFieldSetNumberFormat",
  "pivotSetRowFieldOrder",
  "pivotSetColFieldOrder",
  "pivotDataFieldCount",
  "pivotDataFieldAdd",
  "pivotDataFieldClear",
  "pivotDataFieldSet",
  "pivotFilterCount",
  "pivotFilterAdd",
  "pivotFilterClear",
  "pivotFilterRemoveAt",
  "evaluateCfRange",
  "getSheetView",
  "setSheetZoom",
  "setSheetFreeze",
  "setSheetTabHidden",
  "getSheetProtection",
  "setSheetProtection",
  "getSheetColumns",
  "setColumnWidth",
  "setColumnHidden",
  "setColumnOutline",
  "getSheetRowOverrides",
  "setRowHeight",
  "setRowHidden",
  "setRowOutline",
  "getCellXfIndex",
  "setCellXfIndex",
  "getCellXf",
  "getFont",
  "getFill",
  "getBorder",
  "getNumFmt",
  "addFont",
  "addFill",
  "addBorder",
  "addNumFmt",
  "addXf",
  "fontCount",
  "fillCount",
  "borderCount",
  "xfCount",
  "cellStyleCount",
  "cellStyleXfCount",
  "getCellStyle",
  "getCellStyleXf",
  "getExternalLinks",
  "addMerge",
  "removeMerge",
  "removeMergeAt",
  "clearMerges",
  "getMerges",
  "getComment",
  "setComment",
  "addHyperlink",
  "removeHyperlink",
  "removeHyperlinkAt",
  "clearHyperlinks",
  "getHyperlinks",
  "getValidations",
  "addValidation",
  "removeValidationAt",
  "clearValidations",
  "getConditionalFormats",
  "addConditionalFormat",
  "removeConditionalFormatAt",
  "clearConditionalFormats",
  "precedents",
  "dependents",
  "functionMetadata",
  "functionNames",
  "localizeFunctionName",
  "canonicalizeFunctionName",
  "spillInfo",
]);

const MUTATING_METHOD_PREFIXES = [
  "add",
  "clear",
  "delete",
  "insert",
  "move",
  "pivot",
  "remove",
  "rename",
  "set",
];

function nowIso(): string {
  return new Date().toISOString();
}

function touch(session: WorkbookSession, dirty: boolean): void {
  session.updatedAt = nowIso();
  session.dirty = session.dirty || dirty;
}

function publicInfo(session: WorkbookSession): SessionInfo {
  return {
    id: session.id,
    sourcePath: session.sourcePath,
    outputPath: session.outputPath,
    createdAt: session.createdAt,
    updatedAt: session.updatedAt,
    dirty: session.dirty,
  };
}

function sheetName(session: WorkbookSession, sheet: number): string {
  const name = session.workbook.sheetName(sheet);
  assertStatus(name.status, `read sheet ${sheet} name`);
  return name.value;
}

function refA1(row: number, col: number): string {
  return cellToA1(row, col);
}

function rangeA1(firstRow: number, firstCol: number, lastRow: number, lastCol: number): string {
  return `${refA1(firstRow, firstCol)}:${refA1(lastRow, lastCol)}`;
}

function safeWorkbookCall(session: WorkbookSession, method: string, args: unknown[]): unknown {
  const callable = (session.workbook as unknown as Record<string, unknown>)[method];
  if (typeof callable !== "function") {
    return null;
  }
  try {
    return resultToJson(
      (callable as (...methodArgs: unknown[]) => unknown).apply(session.workbook, args),
    );
  } catch {
    return null;
  }
}

function jsonCellValue(value: unknown): PrimitiveCellValue {
  return valueToJson(value as Parameters<typeof valueToJson>[0]) as PrimitiveCellValue;
}

function primitiveText(value: PrimitiveCellValue): string {
  if (value.value === undefined) {
    return "";
  }
  return String(value.value);
}

function isNonBlank(value: PrimitiveCellValue, formula: string): boolean {
  return formula.length > 0 || value.kind !== "blank";
}

function cellKind(value: PrimitiveCellValue, formula: string): string {
  if (formula) {
    return "formula";
  }
  if (value.kind === "number") {
    return "number";
  }
  if (value.kind === "bool") {
    return "boolean";
  }
  if (value.kind === "text") {
    const text = primitiveText(value);
    if (/^\d{4}[-/年]\d{1,2}[-/月]\d{1,2}/.test(text)) {
      return "date";
    }
    if (/^[¥$€]?\s*-?\d[\d,]*(?:\.\d+)?$/.test(text)) {
      return "money";
    }
    return "text";
  }
  return value.kind;
}

function isNumericLikeKind(kind: string): boolean {
  return kind === "number" || kind === "money" || kind === "formula";
}

function stableCells(
  session: WorkbookSession,
  sheet: number,
  maxCells: number,
  includeStyles: boolean,
) {
  const count = session.workbook.cellCount(sheet);
  const limit = Math.max(0, Math.min(maxCells, count));
  const cells = [];
  for (let index = 0; index < limit; index += 1) {
    const cell = session.workbook.cellAt(sheet, index);
    assertStatus(cell.status, `read sheet ${sheet} cell ${index}`);
    const value = jsonCellValue(cell.value);
    const formula = cell.formula ?? "";
    const xfIndexResult = includeStyles
      ? (safeWorkbookCall(session, "getCellXfIndex", [sheet, cell.row, cell.col]) as {
          xfIndex?: number;
        } | null)
      : null;
    const xfIndex = typeof xfIndexResult?.xfIndex === "number" ? xfIndexResult.xfIndex : null;
    cells.push({
      row: cell.row,
      col: cell.col,
      a1: refA1(cell.row, cell.col),
      value,
      formula,
      kind: cellKind(value, formula),
      style: includeStyles
        ? {
            xfIndex,
            xf: xfIndex === null ? null : safeWorkbookCall(session, "getCellXf", [xfIndex]),
          }
        : undefined,
    });
  }
  cells.sort((left, right) => left.row - right.row || left.col - right.col);
  return { cells, cellCount: count, truncated: limit < count };
}

function usedRange(cells: { row: number; col: number }[]): string | null {
  if (cells.length === 0) {
    return null;
  }
  const rows = cells.map((cell) => cell.row);
  const cols = cells.map((cell) => cell.col);
  return rangeA1(Math.min(...rows), Math.min(...cols), Math.max(...rows), Math.max(...cols));
}

function searchSheets(session: WorkbookSession, sheet: number | string | undefined): number[] {
  if (sheet !== undefined) {
    return [findSheetIndex(session.workbook, sheet)];
  }
  return Array.from({ length: session.workbook.sheetCount() }, (_, index) => index);
}

function textValue(value: unknown): string | undefined {
  if (
    typeof value === "object" &&
    value !== null &&
    "kind" in value &&
    (value as { kind: unknown }).kind === 3 &&
    "text" in value &&
    typeof (value as { text: unknown }).text === "string"
  ) {
    return (value as { text: string }).text;
  }
  return undefined;
}

function makeMatcher(
  query: string,
  options: Pick<SearchOptions, "matchCase" | "wholeCell" | "regex">,
) {
  if (options.regex) {
    const flags = options.matchCase ? "g" : "gi";
    const pattern = new RegExp(query, flags);
    return (text: string) => {
      pattern.lastIndex = 0;
      const matched = pattern.test(text);
      pattern.lastIndex = 0;
      return matched;
    };
  }

  const needle = options.matchCase ? query : query.toLocaleLowerCase();
  return (text: string) => {
    const haystack = options.matchCase ? text : text.toLocaleLowerCase();
    return options.wholeCell ? haystack === needle : haystack.includes(needle);
  };
}

function replaceText(
  text: string,
  query: string,
  replacement: string,
  options: Pick<SearchOptions, "matchCase" | "wholeCell" | "regex">,
): string {
  if (options.regex) {
    return text.replace(new RegExp(query, options.matchCase ? "g" : "gi"), replacement);
  }
  if (options.wholeCell) {
    return options.matchCase
      ? text === query
        ? replacement
        : text
      : text.toLocaleLowerCase() === query.toLocaleLowerCase()
        ? replacement
        : text;
  }
  const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return text.replace(new RegExp(escaped, options.matchCase ? "g" : "gi"), replacement);
}

/** Opens a workbook session from disk, or creates a new default workbook when no path is given. */
export async function openSession(
  inputPath?: string,
  id: string = randomUUID(),
): Promise<SessionInfo> {
  if (sessions.has(id)) {
    throw new Error(`session already exists: ${id}`);
  }
  const workbook = await createOrLoadWorkbook(inputPath);
  const createdAt = nowIso();
  const session: WorkbookSession = {
    id,
    sourcePath: inputPath,
    createdAt,
    updatedAt: createdAt,
    dirty: false,
    workbook,
  };
  sessions.set(id, session);
  return publicInfo(session);
}

/** Lists open workbook sessions without exposing native handles. */
export function listSessions(): SessionInfo[] {
  return Array.from(sessions.values(), publicInfo);
}

/** Returns an open session or throws when the id is unknown. */
export function getSession(id: string): WorkbookSession {
  const session = sessions.get(id);
  if (!session) {
    throw new Error(`session not found: ${id}`);
  }
  return session;
}

/** Closes a session and releases the underlying Formulon workbook handle. */
export function closeSession(id: string): SessionInfo {
  const session = getSession(id);
  sessions.delete(id);
  session.workbook.delete();
  return publicInfo(session);
}

/** Returns a workbook summary for an open session. */
export function inspectSession(id: string, includeCells: boolean, maxCellsPerSheet: number) {
  const session = getSession(id);
  touch(session, false);
  return {
    session: publicInfo(session),
    workbook: workbookSummary(session.workbook, includeCells, maxCellsPerSheet),
  };
}

/** Searches text cells and/or formula text across an open workbook session. */
export function findSessionCells(id: string, query: string, options: SearchOptions) {
  const session = getSession(id);
  if (query.length === 0) {
    throw new Error("query must not be empty");
  }
  const matcher = makeMatcher(query, options);
  const results: CellSearchResult[] = [];
  let truncated = false;

  for (const sheet of searchSheets(session, options.sheet)) {
    const name = sheetName(session, sheet);
    const count = session.workbook.cellCount(sheet);
    for (let index = 0; index < count; index += 1) {
      const cell = session.workbook.cellAt(sheet, index);
      assertStatus(cell.status, `read sheet ${sheet} cell ${index}`);
      const a1 = cellToA1(cell.row, cell.col);
      const addResult = (target: "text" | "formula", text: string) => {
        if (!matcher(text)) {
          return;
        }
        if (results.length >= options.maxResults) {
          truncated = true;
          return;
        }
        results.push({
          sheet,
          sheetName: name,
          row: cell.row,
          col: cell.col,
          a1,
          ref: `${name}!${a1}`,
          target,
          text,
        });
      };

      if (cell.formula) {
        if (options.target === "formulas" || options.target === "both") {
          addResult("formula", cell.formula);
        }
      } else {
        const text = textValue(cell.value);
        if ((options.target === "texts" || options.target === "both") && text !== undefined) {
          addResult("text", text);
        }
      }
      if (truncated) {
        break;
      }
    }
    if (truncated) {
      break;
    }
  }

  touch(session, false);
  return {
    session: publicInfo(session),
    query,
    options,
    results,
    count: results.length,
    truncated,
  };
}

/** Replaces matching text cell values and/or formula text across an open workbook session. */
export function replaceSessionCells(id: string, query: string, options: ReplaceOptions) {
  const session = getSession(id);
  if (query.length === 0) {
    throw new Error("query must not be empty");
  }
  const matcher = makeMatcher(query, options);
  const replacements: CellReplaceResult[] = [];
  let truncated = false;

  for (const sheet of searchSheets(session, options.sheet)) {
    const name = sheetName(session, sheet);
    const count = session.workbook.cellCount(sheet);
    for (let index = 0; index < count; index += 1) {
      const cell = session.workbook.cellAt(sheet, index);
      assertStatus(cell.status, `read sheet ${sheet} cell ${index}`);
      const a1 = cellToA1(cell.row, cell.col);
      const replaceOne = (target: "text" | "formula", text: string) => {
        if (!matcher(text)) {
          return;
        }
        if (replacements.length >= options.maxReplacements) {
          truncated = true;
          return;
        }
        const next = replaceText(text, query, options.replacement, options);
        const status =
          target === "formula"
            ? session.workbook.setFormula(sheet, cell.row, cell.col, next)
            : session.workbook.setText(sheet, cell.row, cell.col, next);
        assertStatus(status, `replace ${target} ${name}!${a1}`);
        replacements.push({
          sheet,
          sheetName: name,
          row: cell.row,
          col: cell.col,
          a1,
          ref: `${name}!${a1}`,
          target,
          text,
          before: text,
          after: next,
          status: statusToJson(status),
        });
      };

      if (cell.formula) {
        if (options.target === "formulas" || options.target === "both") {
          replaceOne("formula", cell.formula);
        }
      } else {
        const text = textValue(cell.value);
        if ((options.target === "texts" || options.target === "both") && text !== undefined) {
          replaceOne("text", text);
        }
      }
      if (truncated) {
        break;
      }
    }
    if (truncated) {
      break;
    }
  }

  if (replacements.length > 0 && options.recalc) {
    assertStatus(session.workbook.recalc(), "recalc workbook");
  }
  touch(session, replacements.length > 0);
  return {
    session: publicInfo(session),
    query,
    replacement: options.replacement,
    options,
    replacements,
    count: replacements.length,
    truncated,
  };
}

/** Returns stable cell, layout, and optional style data without inference. */
export function inspectSessionLayout(id: string, options: LayoutOptions) {
  const session = getSession(id);
  const sheets = searchSheets(session, options.sheet).map((sheet) => {
    const { cells, cellCount, truncated } = stableCells(
      session,
      sheet,
      options.maxCells,
      options.includeStyles,
    );
    return {
      index: sheet,
      name: sheetName(session, sheet),
      usedRange: usedRange(cells),
      cellCount,
      mergedRanges: safeWorkbookCall(session, "getMerges", [sheet]),
      view: safeWorkbookCall(session, "getSheetView", [sheet]),
      columns: safeWorkbookCall(session, "getSheetColumns", [sheet]),
      rows: safeWorkbookCall(session, "getSheetRowOverrides", [sheet]),
      protection: safeWorkbookCall(session, "getSheetProtection", [sheet]),
      cells: options.includeCells ? cells : undefined,
      truncated,
    };
  });
  touch(session, false);
  return { session: publicInfo(session), sheets };
}

function inferColumnKind(values: { kind: string; value: PrimitiveCellValue }[]): string {
  const counts = new Map<string, number>();
  for (const value of values) {
    if (value.value.kind === "blank") {
      continue;
    }
    counts.set(value.kind, (counts.get(value.kind) ?? 0) + 1);
  }
  return Array.from(counts.entries()).sort((left, right) => right[1] - left[1])[0]?.[0] ?? "blank";
}

function labelKind(text: string): string | null {
  const normalized = text.toLocaleLowerCase();
  if (/日付|請求日|発行日|date/.test(normalized)) {
    return "date";
  }
  if (/合計|小計|税込|税|total|subtotal|tax/.test(normalized)) {
    return "total";
  }
  if (/氏名|名前|会社|宛先|customer|client|name|company/.test(normalized)) {
    return "party";
  }
  if (/請求|invoice/.test(normalized)) {
    return "invoice";
  }
  if (/見積|quote|quotation/.test(normalized)) {
    return "quote";
  }
  if (/発注|注文|purchase/.test(normalized)) {
    return "purchase_order";
  }
  if (/領収|receipt/.test(normalized)) {
    return "receipt";
  }
  return null;
}

function detectSheetRegions(session: WorkbookSession, sheet: number, maxCells: number) {
  const { cells, truncated } = stableCells(session, sheet, maxCells, false);
  const name = sheetName(session, sheet);
  const byRow = new Map<number, typeof cells>();
  const byAddress = new Map<string, (typeof cells)[number]>();
  for (const cell of cells) {
    if (!isNonBlank(cell.value, cell.formula)) {
      continue;
    }
    byAddress.set(`${cell.row}:${cell.col}`, cell);
    const row = byRow.get(cell.row) ?? [];
    row.push(cell);
    byRow.set(cell.row, row);
  }

  const regions = [];
  const denseRows = Array.from(byRow.entries())
    .filter(([, rowCells]) => rowCells.length >= 2)
    .sort(([left], [right]) => left - right);
  let idx = 0;
  while (idx < denseRows.length) {
    const group = [denseRows[idx]];
    idx += 1;
    while (idx < denseRows.length && denseRows[idx][0] === group[group.length - 1][0] + 1) {
      group.push(denseRows[idx]);
      idx += 1;
    }
    const groupCells = group.flatMap(([, rowCells]) => rowCells);
    const cols = Array.from(new Set(groupCells.map((cell) => cell.col))).sort((a, b) => a - b);
    if (group.length >= 2 && cols.length >= 2) {
      const firstRow = group[0][0];
      const lastRow = group[group.length - 1][0];
      const firstCol = cols[0];
      const lastCol = cols[cols.length - 1];
      const headerCells = group[0][1];
      const headerTextCount = headerCells.filter((cell) => cell.kind === "text").length;
      const confidence = Math.min(
        0.95,
        0.45 +
          Math.min(0.25, group.length / 20) +
          Math.min(0.25, cols.length / 20) +
          (headerTextCount >= 2 ? 0.15 : 0),
      );
      regions.push({
        type: "table",
        sheet,
        sheetName: name,
        range: rangeA1(firstRow, firstCol, lastRow, lastCol),
        confidence: Number(confidence.toFixed(2)),
        headerRow: firstRow,
        dataRange: firstRow < lastRow ? rangeA1(firstRow + 1, firstCol, lastRow, lastCol) : null,
        columns: cols.map((col) => {
          const header = headerCells.find((cell) => cell.col === col);
          const values = groupCells
            .filter((cell) => cell.col === col && cell.row > firstRow)
            .map((cell) => ({ kind: cell.kind, value: cell.value }));
          return {
            name: header ? primitiveText(header.value) : "",
            col,
            kind: inferColumnKind(values),
          };
        }),
        evidence: [
          "contiguous non-empty rows",
          "multiple occupied columns",
          headerTextCount >= 2 ? "header-like first row" : "weak header row",
        ],
      });
    }
  }

  for (const table of regions.filter((region) => region.type === "table")) {
    const range = parseSimpleRange(table.range);
    for (let row = range.lastRow + 1; row <= range.lastRow + 8; row += 1) {
      for (
        let col = Math.max(range.firstCol, range.lastCol - 2);
        col <= range.lastCol + 1;
        col += 1
      ) {
        const valueCell = byAddress.get(`${row}:${col}`);
        const labelCell = byAddress.get(`${row}:${col - 1}`);
        if (valueCell && labelCell?.kind === "text" && isNumericLikeKind(valueCell.kind)) {
          regions.push({
            type: "total",
            sheet,
            sheetName: name,
            range: rangeA1(labelCell.row, labelCell.col, valueCell.row, valueCell.col),
            confidence: 0.76,
            pairs: [
              {
                label: primitiveText(labelCell.value),
                labelCell: refA1(labelCell.row, labelCell.col),
                valueCell: refA1(valueCell.row, valueCell.col),
                valueKind: valueCell.kind,
                semanticKind: "total",
              },
            ],
            evidence: ["numeric or formula summary structure below table"],
          });
        }
      }
    }
  }

  for (const cell of cells) {
    const text = primitiveText(cell.value);
    if (!text || cell.kind !== "text") {
      continue;
    }
    const kind = labelKind(text);
    const right = byAddress.get(`${cell.row}:${cell.col + 1}`);
    const below = byAddress.get(`${cell.row + 1}:${cell.col}`);
    const valueCell = right ?? below;
    const genericPair = valueCell && cell.row <= 12 && (right || cell.col <= 3);
    if (valueCell && (kind || genericPair)) {
      regions.push({
        type: kind === "total" ? "total" : "labelValue",
        sheet,
        sheetName: name,
        range: rangeA1(
          Math.min(cell.row, valueCell.row),
          Math.min(cell.col, valueCell.col),
          Math.max(cell.row, valueCell.row),
          Math.max(cell.col, valueCell.col),
        ),
        confidence: kind === "total" ? 0.82 : 0.72,
        pairs: [
          {
            label: text,
            labelCell: refA1(cell.row, cell.col),
            valueCell: refA1(valueCell.row, valueCell.col),
            valueKind: valueCell.kind,
            semanticKind: kind ?? "field",
          },
        ],
        evidence: [
          kind ? "recognized label keyword" : "adjacent label-value structure",
          right ? "value cell to the right" : "value cell below",
        ],
      });
    }
  }

  return { sheet, sheetName: name, regions, truncated };
}

/** Detects tables, label-value pairs, and total-like regions with rule-based evidence. */
export function detectSessionRegions(
  id: string,
  sheet: number | string | undefined,
  maxCells: number,
) {
  const session = getSession(id);
  const sheets = searchSheets(session, sheet).map((sheetIndex) =>
    detectSheetRegions(session, sheetIndex, maxCells),
  );
  touch(session, false);
  return {
    session: publicInfo(session),
    sheets,
    regions: sheets.flatMap((entry) => entry.regions),
  };
}

type WorkbookKind =
  | "invoice"
  | "quote"
  | "purchase_order"
  | "receipt"
  | "list"
  | "report"
  | "schedule"
  | "form"
  | "ledger"
  | "unknown";

function rankedCandidates(score: Map<WorkbookKind, number>) {
  const entries = Array.from(score.entries())
    .map(([type, raw]) => ({ type, confidence: Number(Math.min(0.99, raw).toFixed(2)) }))
    .sort((left, right) => right.confidence - left.confidence);
  return entries.length ? entries : [{ type: "unknown" as const, confidence: 0 }];
}

function addScore(score: Map<WorkbookKind, number>, type: WorkbookKind, amount: number) {
  score.set(type, (score.get(type) ?? 0) + amount);
}

function parseSimpleRange(ref: string) {
  const [start, end] = ref.split(":");
  const startCell = parseCellRef(start);
  const endCell = parseCellRef(end ?? start);
  return {
    firstRow: Math.min(startCell.row, endCell.row),
    firstCol: Math.min(startCell.col, endCell.col),
    lastRow: Math.max(startCell.row, endCell.row),
    lastCol: Math.max(startCell.col, endCell.col),
  };
}

/** Classifies workbook shape using deterministic features and evidence. */
export function analyzeSessionWorkbook(
  id: string,
  includeEvidence: boolean,
  maxCellsPerSheet: number,
) {
  const session = getSession(id);
  const scores = new Map<WorkbookKind, number>();
  const evidence: string[] = [];
  const detected = detectSessionRegions(id, undefined, maxCellsPerSheet);
  const allCells = searchSheets(session, undefined).flatMap((sheet) => {
    const { cells } = stableCells(session, sheet, maxCellsPerSheet, false);
    return cells.map((cell) => ({ ...cell, sheet, sheetName: sheetName(session, sheet) }));
  });

  for (const cell of allCells) {
    const text = primitiveText(cell.value).toLocaleLowerCase();
    if (/請求書|invoice/.test(text)) {
      addScore(scores, "invoice", 0.15);
      evidence.push(`invoice-like keyword at ${cell.sheetName}!${cell.a1}`);
    }
    if (/見積書|quote|quotation/.test(text)) {
      addScore(scores, "quote", 0.15);
      evidence.push(`quote-like keyword at ${cell.sheetName}!${cell.a1}`);
    }
    if (/発注書|注文書|purchase order/.test(text)) {
      addScore(scores, "purchase_order", 0.15);
      evidence.push(`purchase-order-like keyword at ${cell.sheetName}!${cell.a1}`);
    }
    if (/領収書|receipt/.test(text)) {
      addScore(scores, "receipt", 0.15);
      evidence.push(`receipt-like keyword at ${cell.sheetName}!${cell.a1}`);
    }
    if (/予定|schedule/.test(text)) {
      addScore(scores, "schedule", 0.12);
    }
    if (/台帳|ledger/.test(text)) {
      addScore(scores, "ledger", 0.12);
    }
  }

  const tables = detected.regions.filter((region) => region.type === "table");
  const totals = detected.regions.filter((region) => region.type === "total");
  const labelValues = detected.regions.filter((region) => region.type === "labelValue");
  const cellsBySheetAddress = new Map<string, (typeof allCells)[number]>();
  for (const cell of allCells) {
    cellsBySheetAddress.set(`${cell.sheet}:${cell.row}:${cell.col}`, cell);
  }

  const upperLabelPairs = labelValues.filter((region) => {
    const range = parseSimpleRange(region.range);
    return range.firstRow <= 12;
  });
  const structuralTotals = tables.flatMap((region) => {
    const range = parseSimpleRange(region.range);
    const candidates = [];
    for (let row = range.lastRow + 1; row <= range.lastRow + 8; row += 1) {
      for (
        let col = Math.max(range.firstCol, range.lastCol - 2);
        col <= range.lastCol + 1;
        col += 1
      ) {
        const valueCell = cellsBySheetAddress.get(`${region.sheet}:${row}:${col}`);
        const leftCell = cellsBySheetAddress.get(`${region.sheet}:${row}:${col - 1}`);
        if (valueCell && leftCell?.kind === "text" && isNumericLikeKind(valueCell.kind)) {
          candidates.push({
            sheetName: region.sheetName,
            labelCell: leftCell.a1,
            valueCell: valueCell.a1,
            valueKind: valueCell.kind,
          });
        }
      }
    }
    return candidates;
  });

  if (tables.length > 0) {
    addScore(scores, "list", 0.35);
    addScore(scores, "report", 0.2);
    evidence.push("table-like region detected");
  }
  if (upperLabelPairs.length >= 2) {
    addScore(scores, "form", 0.3);
    evidence.push("multiple upper label-value structures detected");
  }
  if (totals.length > 0 || structuralTotals.length > 0) {
    addScore(scores, "invoice", 0.3);
    addScore(scores, "quote", 0.18);
    addScore(scores, "receipt", 0.15);
    evidence.push(
      structuralTotals.length > 0
        ? "numeric or formula summary cell detected below a table"
        : "total-like field detected",
    );
  }
  for (const table of tables) {
    const numericColumns =
      "columns" in table && Array.isArray(table.columns)
        ? table.columns.filter((column) => isNumericLikeKind(column.kind)).length
        : 0;
    const range = parseSimpleRange(table.range);
    const startsAfterMetadata = range.firstRow >= 3;
    const hasSeveralColumns = range.lastCol - range.firstCol + 1 >= 3;
    if (numericColumns >= 2 && hasSeveralColumns) {
      addScore(scores, "list", 0.15);
      evidence.push("table has multiple numeric or formula columns");
    }
    if (startsAfterMetadata && upperLabelPairs.length >= 1 && structuralTotals.length > 0) {
      addScore(scores, "invoice", 0.35);
      addScore(scores, "quote", 0.22);
      evidence.push("upper metadata, line-item table, and below-table summary detected together");
    }
  }
  if (tables.length === 1 && upperLabelPairs.length === 0 && structuralTotals.length === 0) {
    addScore(scores, "list", 0.2);
    evidence.push("single table dominates without document metadata");
  }

  const candidates = rankedCandidates(scores);
  const primary = candidates[0];
  const primaryType = primary.confidence >= 0.6 ? primary.type : "unknown";
  const likelyTitle = allCells
    .filter((cell) => cell.kind === "text")
    .sort((left, right) => left.row - right.row || left.col - right.col)[0];

  touch(session, false);
  return {
    session: publicInfo(session),
    classification: {
      primaryType,
      confidence: primaryType === "unknown" ? 0 : primary.confidence,
      candidates,
    },
    summary: {
      likelyTitle: likelyTitle
        ? {
            text: primitiveText(likelyTitle.value),
            cell: `${likelyTitle.sheetName}!${likelyTitle.a1}`,
          }
        : null,
      tables: tables.map((region) => `${region.sheetName}!${region.range}`),
      totals: totals.flatMap((region) => ("pairs" in region ? region.pairs : [])),
      keyFields: labelValues.flatMap((region) =>
        "pairs" in region && region.pairs
          ? region.pairs.map((pair) => ({
              name: pair.semanticKind,
              label: pair.label,
              valueCell: `${region.sheetName}!${pair.valueCell}`,
              confidence: region.confidence,
            }))
          : [],
      ),
    },
    evidence: includeEvidence ? Array.from(new Set(evidence)) : undefined,
    warnings: detected.sheets.some((entry) => entry.truncated)
      ? ["analysis truncated because maxCellsPerSheet was reached"]
      : [],
  };
}

/** Recalculates a session workbook. */
export function recalcSession(id: string) {
  const session = getSession(id);
  const status = session.workbook.recalc();
  assertStatus(status, "recalc workbook");
  touch(session, true);
  return { session: publicInfo(session), status: statusToJson(status) };
}

/** Applies a sheet add/remove/rename/move operation. */
export function applySheetOperation(
  id: string,
  operation: "add" | "remove" | "rename" | "move",
  args: {
    name?: string;
    index?: number;
    newName?: string;
    fromIndex?: number;
    toIndex?: number;
  },
) {
  const session = getSession(id);
  const workbook = session.workbook;
  const status =
    operation === "add"
      ? workbook.addSheet(requiredString(args.name, "name"))
      : operation === "remove"
        ? workbook.removeSheet(requiredNumber(args.index, "index"))
        : operation === "rename"
          ? workbook.renameSheet(
              requiredNumber(args.index, "index"),
              requiredString(args.newName, "newName"),
            )
          : workbook.moveSheet(
              requiredNumber(args.fromIndex, "fromIndex"),
              requiredNumber(args.toIndex, "toIndex"),
            );
  assertStatus(status, `${operation} sheet`);
  touch(session, true);
  return { session: publicInfo(session), status: statusToJson(status) };
}

/** Adds, replaces, or removes a workbook-scoped defined name. */
export function setSessionDefinedName(id: string, name: string, formula: string) {
  const session = getSession(id);
  const status = session.workbook.setDefinedName(name, formula ? normalizeFormula(formula) : "");
  assertStatus(status, "set defined name");
  touch(session, true);
  return { session: publicInfo(session), status: statusToJson(status) };
}

/** Inserts or deletes rows or columns while letting Formulon rewrite affected references. */
export function editSessionStructure(
  id: string,
  operation: "insertRows" | "deleteRows" | "insertCols" | "deleteCols",
  sheet: number | string | undefined,
  start: number,
  count: number,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const status =
    operation === "insertRows"
      ? session.workbook.insertRows(sheetIndex, start, count)
      : operation === "deleteRows"
        ? session.workbook.deleteRows(sheetIndex, start, count)
        : operation === "insertCols"
          ? session.workbook.insertCols(sheetIndex, start, count)
          : session.workbook.deleteCols(sheetIndex, start, count);
  assertStatus(status, operation);
  touch(session, true);
  return { session: publicInfo(session), sheet: sheetIndex, status: statusToJson(status) };
}

/** Updates sheet view settings such as zoom, frozen panes, and tab hidden state. */
export function setSessionSheetView(
  id: string,
  sheet: number | string | undefined,
  options: { zoom?: number; freezeRows?: number; freezeCols?: number; hidden?: boolean },
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const statuses = [];
  if (options.zoom !== undefined) {
    const status = session.workbook.setSheetZoom(sheetIndex, options.zoom);
    assertStatus(status, "set sheet zoom");
    statuses.push(statusToJson(status));
  }
  if (options.freezeRows !== undefined || options.freezeCols !== undefined) {
    const status = session.workbook.setSheetFreeze(
      sheetIndex,
      options.freezeRows ?? 0,
      options.freezeCols ?? 0,
    );
    assertStatus(status, "set sheet freeze");
    statuses.push(statusToJson(status));
  }
  if (options.hidden !== undefined) {
    const status = session.workbook.setSheetTabHidden(sheetIndex, options.hidden);
    assertStatus(status, "set sheet tab hidden");
    statuses.push(statusToJson(status));
  }
  touch(session, statuses.length > 0);
  return { session: publicInfo(session), sheet: sheetIndex, statuses };
}

/** Reads broad workbook metadata that does not require mutation. */
export function getSessionMetadata(id: string, kind: "functions" | "externalLinks") {
  const session = getSession(id);
  touch(session, false);
  return {
    session: publicInfo(session),
    kind,
    value:
      kind === "functions"
        ? resultToJson(session.workbook.functionNames())
        : resultToJson(session.workbook.getExternalLinks()),
  };
}

function resolveCell(session: WorkbookSession, mutation: FlexibleCellMutation) {
  if (mutation.a1) {
    const parsed = parseCellRef(mutation.a1);
    return {
      sheet: findSheetIndex(session.workbook, parsed.sheetName ?? mutation.sheet),
      row: parsed.row,
      col: parsed.col,
    };
  }
  if (mutation.row === undefined || mutation.col === undefined) {
    throw new Error("mutation requires either a1 or row/col");
  }
  return {
    sheet: findSheetIndex(session.workbook, mutation.sheet),
    row: mutation.row,
    col: mutation.col,
  };
}

/** Applies flexible cell mutations using either A1 references or zero-based coordinates. */
export function applySessionMutations(
  id: string,
  mutations: FlexibleCellMutation[],
  recalc: boolean,
) {
  const session = getSession(id);
  const applied = [];
  for (const [idx, mutation] of mutations.entries()) {
    const address = resolveCell(session, mutation);
    const concrete = { ...mutation, ...address } as CellMutation;
    const status = applyMutation(session.workbook, concrete);
    assertStatus(status, `apply mutation ${idx}`);
    applied.push({
      index: idx,
      a1: cellToA1(address.row, address.col),
      sheet: address.sheet,
      status: statusToJson(status),
    });
  }
  if (recalc) {
    assertStatus(session.workbook.recalc(), "recalc workbook");
  }
  touch(session, true);
  return { session: publicInfo(session), applied };
}

/** Reads one cell from an open session using zero-based coordinates. */
export function getSessionCell(
  id: string,
  sheet: number | string | undefined,
  row: number,
  col: number,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const result = session.workbook.getValue(sheetIndex, row, col);
  assertStatus(result.status, "get cell");
  touch(session, false);
  return {
    session: publicInfo(session),
    address: { sheet: sheetIndex, row, col, a1: cellToA1(row, col) },
    status: statusToJson(result.status),
    value: valueToJson(result.value),
  };
}

/** Reads one cell from an open session using an A1 reference. */
export function getSessionCellByA1(id: string, ref: string) {
  const parsed = parseCellRef(ref);
  return getSessionCell(id, parsed.sheetName, parsed.row, parsed.col);
}

/** Reads a rectangular A1 range from an open session. */
export function getSessionRange(id: string, ref: string) {
  const session = getSession(id);
  const range = normalizeRange(parseRangeRef(ref));
  const sheetIndex = findSheetIndex(session.workbook, range.sheetName);
  const rows = [];
  for (let row = range.start.row; row <= range.end.row; row += 1) {
    const values = [];
    for (let col = range.start.col; col <= range.end.col; col += 1) {
      const result = session.workbook.getValue(sheetIndex, row, col);
      assertStatus(result.status, `get cell ${cellToA1(row, col)}`);
      values.push({
        a1: cellToA1(row, col),
        value: valueToJson(result.value),
      });
    }
    rows.push(values);
  }
  touch(session, false);
  return {
    session: publicInfo(session),
    range: {
      sheet: sheetIndex,
      ref,
      start: cellToA1(range.start.row, range.start.col),
      end: cellToA1(range.end.row, range.end.col),
    },
    rows,
  };
}

/** Calls an allowlisted low-level Formulon Workbook method with positional JSON arguments. */
export function callWorkbookMethod(id: string, method: string, args: unknown[]) {
  if (!WORKBOOK_METHODS.has(method)) {
    throw new Error(`workbook method is not allowlisted: ${method}`);
  }
  const session = getSession(id);
  const callable = (session.workbook as unknown as Record<string, unknown>)[method];
  if (typeof callable !== "function") {
    throw new Error(`workbook method is not callable: ${method}`);
  }
  const result = (callable as (...methodArgs: unknown[]) => unknown).apply(session.workbook, args);
  const dirty = MUTATING_METHOD_PREFIXES.some((prefix) => method.startsWith(prefix));
  touch(session, dirty);
  return {
    session: publicInfo(session),
    method,
    result: resultToJson(result),
  };
}

/** Saves an open session to disk. */
export async function saveSession(id: string, outputPath?: string) {
  const session = getSession(id);
  const destination = outputPath ?? session.outputPath ?? session.sourcePath;
  if (!destination) {
    throw new Error("outputPath is required for a new workbook session");
  }
  const bytes = await saveWorkbook(session.workbook, destination);
  session.outputPath = destination;
  session.dirty = false;
  touch(session, false);
  return { session: publicInfo(session), outputPath: destination, bytes };
}

function requiredString(value: string | undefined, name: string): string {
  if (value === undefined || value.length === 0) {
    throw new Error(`${name} is required`);
  }
  return value;
}

function requiredNumber(value: number | undefined, name: string): number {
  if (value === undefined) {
    throw new Error(`${name} is required`);
  }
  return value;
}
