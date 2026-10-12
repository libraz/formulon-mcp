import { assertStatus, findSheetIndex, statusToJson, type Workbook } from "../formulon.js";
import { preflightLayout } from "../layout/preflight.js";
import type { LayoutOperation, PreparedLayoutOperation } from "../layout/types.js";
import { writePrintSettings } from "../print.js";
import { applyRangeStyle } from "../styles.js";
import { columnCharsToPoints, pointsToMm, pointsToPixels } from "./geometry.js";
import { rangeA1, sheetName } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

export type { LayoutOperation } from "../layout/types.js";

type DisplayGeometry = {
  points: number;
  pixels: number;
  mm: number;
};

function displayGeometry(points: number): DisplayGeometry {
  return { points, pixels: pointsToPixels(points), mm: pointsToMm(points) };
}

function applyMerge(
  workbook: Workbook,
  sheet: number,
  operation: Extract<PreparedLayoutOperation, { type: "merge" }>,
) {
  return workbook[operation.operation === "add" ? "addMerge" : "removeMerge"](sheet, {
    firstRow: operation.resolved.rect.firstRow,
    firstCol: operation.resolved.rect.firstCol,
    lastRow: operation.resolved.rect.lastRow,
    lastCol: operation.resolved.rect.lastCol,
  });
}

/** Applies a bounded sequence of presentation mutations after complete preflight. */
export function applySessionLayout(
  id: string,
  sheet: number | string | undefined,
  operations: readonly LayoutOperation[],
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const prepared = preflightLayout(session.workbook, sheetIndex, operations);
  const results: unknown[] = [];

  for (const operation of prepared) {
    if (operation.type === "column") {
      touch(session, true);
      const status = session.workbook.setColumnWidth(
        sheetIndex,
        operation.first,
        operation.last,
        operation.storedSize,
      );
      assertStatus(status, "set column width");
      results.push({
        type: operation.type,
        first: operation.first,
        last: operation.last,
        storedUnits: "chars",
        storedSize: operation.storedSize,
        displayGeometry: displayGeometry(
          columnCharsToPoints(session.workbook, sheetIndex, operation.storedSize),
        ),
      });
      continue;
    }
    if (operation.type === "row") {
      for (let row = operation.first; row <= operation.last; row += 1) {
        touch(session, true);
        const status = session.workbook.setRowHeight(sheetIndex, row, operation.storedSize);
        assertStatus(status, "set row height");
      }
      results.push({
        type: operation.type,
        first: operation.first,
        last: operation.last,
        storedUnits: "pt",
        storedSize: operation.storedSize,
        displayGeometry: displayGeometry(operation.storedSize),
      });
      continue;
    }
    if (operation.type === "merge") {
      touch(session, true);
      const status = applyMerge(session.workbook, sheetIndex, operation);
      assertStatus(status, `${operation.operation} merge`);
      results.push({
        type: operation.type,
        operation: operation.operation,
        range: operation.resolved.ref,
        status: statusToJson(status),
      });
      continue;
    }
    if (operation.type === "style") {
      touch(session, true);
      const applied = applyRangeStyle(
        session.workbook,
        sheetIndex,
        operation.resolved.rect,
        operation.style,
        operation.base,
      );
      results.push({
        type: operation.type,
        range: operation.resolved.ref,
        base: operation.base,
        regions: applied.regions.map((region) => ({
          range: rangeA1(region.firstRow, region.firstCol, region.lastRow, region.lastCol),
          baseXfIndex: region.baseXfIndex,
          xfIndex: region.xfIndex,
        })),
      });
      continue;
    }
    touch(session, true);
    const applied = writePrintSettings(session.workbook, sheetIndex, operation.settings);
    results.push({ type: operation.type, applied });
  }

  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    sheetName: sheetName(session, sheetIndex),
    operations: results,
  };
}
