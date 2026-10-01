import { parseCellRef } from "../a1.js";
import {
  isNonBlank,
  isNumericLikeKind,
  type PrimitiveCellValue,
  primitiveText,
  rangeA1,
  refA1,
  safeWorkbookCall,
  searchSheets,
  sheetName,
  stableCells,
  usedRange,
} from "./shared.js";
import { getSession, publicInfo, touch, type WorkbookSession } from "./state.js";

export type LayoutOptions = {
  sheet?: number | string;
  includeCells: boolean;
  includeStyles: boolean;
  maxCells: number;
};
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
