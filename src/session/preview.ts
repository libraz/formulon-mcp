import { writeFileSync } from "node:fs";
import { ValueKind } from "@libraz/formulon";
import { cellToA1, normalizeRange, parseRangeRef } from "../a1.js";
import { assertStatus, findSheetIndex, type Workbook } from "../formulon.js";
import { pointsToPixels } from "../preview/geometry.js";
import { readObjects } from "../preview/objects.js";
import {
  BASE_WARNINGS,
  buildMetadataCells,
  cellText,
  chooseRect,
  clipMerge,
  dimensions,
  formulaMap,
  intersects,
  type MergeRange,
  type Rect,
  readMergesInRange,
  rectArea,
  styleFor,
  type WarningState,
} from "../preview/read.js";
import { renderPreviewPng, renderPreviewSvg } from "../preview/svg.js";
import type {
  PreviewCell,
  PreviewMetadata,
  PreviewModel,
  PreviewOptions,
  PreviewPages,
  PreviewResult,
  PreviewStyle,
} from "../preview/types.js";
import { resolveColumnMetrics } from "./geometry.js";
import { assertInGrid } from "./shared.js";
import { getSession, touch } from "./state.js";

const MAX_PREVIEW_CELLS = 10_000;
const MAX_PREVIEW_DIMENSION = 4_096;
const MAX_PREVIEW_PIXELS = 8_000_000;
const MAX_SVG_BYTES = 4 * 1024 * 1024;
const MAX_PAGE_RANGES = 50;

function validateDimensions(
  logicalWidth: number,
  logicalHeight: number,
  scale: number,
): { width: number; height: number } {
  const width = Math.max(1, Math.ceil(Math.ceil(pointsToPixels(logicalWidth)) * scale));
  const height = Math.max(1, Math.ceil(Math.ceil(pointsToPixels(logicalHeight)) * scale));
  if (width > MAX_PREVIEW_DIMENSION || height > MAX_PREVIEW_DIMENSION) {
    throw new Error(
      `preview dimensions ${width}x${height} exceed the ${MAX_PREVIEW_DIMENSION}px limit`,
    );
  }
  if (width * height > MAX_PREVIEW_PIXELS) {
    throw new Error(
      `preview image ${width}x${height} exceeds the ${MAX_PREVIEW_PIXELS}-pixel limit`,
    );
  }
  return { width, height };
}

function pngDimensions(png: Buffer): { width: number; height: number } {
  if (png.length < 24 || png.subarray(0, 8).toString("hex") !== "89504e470d0a1a0a") {
    throw new Error("preview renderer returned an invalid PNG");
  }
  return { width: png.readUInt32BE(16), height: png.readUInt32BE(20) };
}

function rangeCutsMerge(rect: Rect, merges: MergeRange[]): MergeRange | undefined {
  for (const merge of merges) {
    if (
      intersects(merge, rect) &&
      !(
        merge.firstRow >= rect.firstRow &&
        merge.firstCol >= rect.firstCol &&
        merge.lastRow <= rect.lastRow &&
        merge.lastCol <= rect.lastCol
      )
    ) {
      return merge;
    }
  }
  return undefined;
}

function pageSummary(pagination: ReturnType<Workbook["paginate"]>): PreviewPages {
  return {
    count: pagination.pageCount,
    scale: pagination.scale,
    paperWidthPt: pagination.paper.widthPt,
    paperHeightPt: pagination.paper.heightPt,
    landscape: pagination.paper.landscape,
    ranges: pagination.pages
      .slice(0, MAX_PAGE_RANGES)
      .map(
        (page) =>
          `${cellToA1(page.firstRow, page.firstCol)}:${cellToA1(page.lastRow, page.lastCol)}`,
      ),
    rangesTruncated: pagination.pages.length > MAX_PAGE_RANGES,
  };
}

/** Reads one bounded range, builds SVG, and rasterizes it as a PNG. */
export function previewSessionRange(id: string, options: PreviewOptions = {}): PreviewResult {
  const session = getSession(id);
  const workbook = session.workbook;
  const scale = options.scale ?? 1;
  if (!Number.isFinite(scale) || scale < 0.25 || scale > 2) {
    throw new Error("scale must be between 0.25 and 2");
  }

  // Resolve and validate the request before any optional recalc can dirty the
  // session. A qualified range is authoritative when no sheet argument was
  // supplied; an explicit conflicting sheet is rejected.
  const parsedRange = options.range ? normalizeRange(parseRangeRef(options.range)) : undefined;
  const qualifiedSheet = parsedRange?.sheetName;
  const requestedSheet =
    options.sheet === undefined ? undefined : findSheetIndex(workbook, options.sheet);
  const qualifiedSheetIndex =
    qualifiedSheet === undefined ? undefined : findSheetIndex(workbook, qualifiedSheet);
  if (
    requestedSheet !== undefined &&
    qualifiedSheetIndex !== undefined &&
    requestedSheet !== qualifiedSheetIndex
  ) {
    throw new Error("range sheet and sheet argument refer to different worksheets");
  }
  const sheet = qualifiedSheetIndex ?? requestedSheet ?? 0;
  const sheetNameResult = workbook.sheetName(sheet);
  assertStatus(sheetNameResult.status, "read preview sheet name");

  const warnings: WarningState = { values: new Set(BASE_WARNINGS), truncated: false };
  const chosen = chooseRect(workbook, sheet, options.range);
  assertInGrid(chosen.rect.firstRow, chosen.rect.firstCol);
  assertInGrid(chosen.rect.lastRow, chosen.rect.lastCol);
  const merges = readMergesInRange(workbook, sheet, chosen.rect);
  if (options.range) {
    const cutMerge = rangeCutsMerge(chosen.rect, merges);
    if (cutMerge) {
      throw new Error(
        `range ${options.range} cuts merged cell ${cellToA1(cutMerge.firstRow, cutMerge.firstCol)}:${cellToA1(cutMerge.lastRow, cutMerge.lastCol)}`,
      );
    }
  }
  const area = rectArea(chosen.rect);
  if (area > MAX_PREVIEW_CELLS) {
    throw new Error(`preview range spans ${area} cells (limit ${MAX_PREVIEW_CELLS})`);
  }

  const metrics = resolveColumnMetrics(workbook, sheet);
  const size = dimensions(workbook, sheet, chosen.rect);
  const logicalWidth = size.columns.reduce((sum, value) => sum + value, 0);
  const logicalHeight = size.rows.reduce((sum, value) => sum + value, 0);
  if (logicalWidth <= 0 || logicalHeight <= 0) {
    throw new Error("preview range has no visible rows or columns");
  }
  validateDimensions(logicalWidth, logicalHeight, scale);

  let outputPath: string | undefined;
  let outputExtension: string | undefined;
  if (options.outputPath !== undefined) {
    outputExtension = options.outputPath.toLowerCase().slice(-4);
    if (outputExtension !== ".png" && outputExtension !== ".svg") {
      throw new Error("outputPath must end with .png or .svg");
    }
    outputPath = options.outputPath;
  }

  // Inventory formulas before the optional recalc so an oversized sparse
  // workbook is rejected while the session is still clean.
  const formulas = formulaMap(workbook, sheet, chosen.rect);
  const pagination = workbook.paginate(sheet);
  assertStatus(pagination.status, "paginate preview sheet");
  if (options.recalc) {
    touch(session, true);
    assertStatus(workbook.recalc(), "recalc preview workbook");
  }

  const rowOffsets = [0];
  for (const height of size.rows) {
    rowOffsets.push(rowOffsets[rowOffsets.length - 1] + height);
  }
  const colOffsets = [0];
  for (const width of size.columns) {
    colOffsets.push(colOffsets[colOffsets.length - 1] + width);
  }
  const clippedMerges = merges
    .map((merge) => clipMerge(merge, chosen.rect, warnings))
    .filter((merge): merge is MergeRange => merge !== undefined);
  const mergeByCell = new Map<string, MergeRange>();
  for (const merge of clippedMerges) {
    for (let row = merge.firstRow; row <= merge.lastRow; row += 1) {
      for (let col = merge.firstCol; col <= merge.lastCol; col += 1) {
        mergeByCell.set(`${row},${col}`, merge);
      }
    }
  }

  const styleCache = new Map<number, PreviewStyle>();
  const cells: PreviewCell[] = [];
  const fonts = new Set<string>();
  for (let row = chosen.rect.firstRow; row <= chosen.rect.lastRow; row += 1) {
    for (let col = chosen.rect.firstCol; col <= chosen.rect.lastCol; col += 1) {
      const style = styleFor(workbook, sheet, row, col, metrics, warnings, styleCache);
      const value = workbook.getValue(sheet, row, col);
      assertStatus(value.status, `read preview cell ${cellToA1(row, col)}`);
      const text =
        value.value.kind === ValueKind.Blank ? "" : cellText(workbook, sheet, row, col, warnings);
      const formula = formulas.get(`${row},${col}`) ?? "";
      fonts.add(style.fontFamily);
      cells.push({
        row,
        col,
        x: colOffsets[col - chosen.rect.firstCol],
        y: rowOffsets[row - chosen.rect.firstRow],
        width: size.columns[col - chosen.rect.firstCol],
        height: size.rows[row - chosen.rect.firstRow],
        text,
        formula,
        numeric: value.value.kind === ValueKind.Number,
        style,
        merge: mergeByCell.get(`${row},${col}`),
      });
    }
  }
  const drawn = readObjects(workbook, sheet, chosen.rect, warnings);
  const model: PreviewModel = {
    firstRow: chosen.rect.firstRow,
    firstCol: chosen.rect.firstCol,
    lastRow: chosen.rect.lastRow,
    lastCol: chosen.rect.lastCol,
    logicalWidth,
    logicalHeight,
    logicalUnit: "pt",
    columns: size.columns,
    rows: size.rows,
    cells,
    objects: drawn.objects,
    rowBreaks: pagination.horizontalBreaks,
    colBreaks: pagination.verticalBreaks,
    showGridLines: options.showGridLines ?? false,
    showPageBreaks: options.showPageBreaks ?? true,
  };
  const svg = renderPreviewSvg(model);
  if (Buffer.byteLength(svg, "utf8") > MAX_SVG_BYTES) {
    throw new Error(`preview SVG exceeds the ${MAX_SVG_BYTES}-byte limit`);
  }
  const png = renderPreviewPng(svg, scale);
  const renderedDimensions = pngDimensions(png);
  const metadataCells = buildMetadataCells(cells);
  const pages = pageSummary(pagination);
  const warningList = Array.from(warnings.values);
  const metadata: PreviewMetadata = {
    sheet,
    sheetName: sheetNameResult.value,
    range: `${sheetNameResult.value}!${chosen.rangeLabel}`,
    logicalWidth,
    logicalHeight,
    logicalUnit: "pt",
    scaledWidth: renderedDimensions.width,
    scaledHeight: renderedDimensions.height,
    scale,
    fontList: Array.from(fonts),
    pageCount: pagination.pageCount,
    pages,
    objects: drawn.metadata,
    rowBreaks: pagination.horizontalBreaks,
    colBreaks: pagination.verticalBreaks,
    warnings: warningList,
    cells: metadataCells.cells,
    cellGeometryUnit: "pt",
    metadataTruncated: metadataCells.truncated,
    omittedCellCount: metadataCells.omitted,
    warningsTruncated: warnings.truncated,
  };
  if (outputPath !== undefined && outputExtension !== undefined) {
    writeFileSync(outputPath, outputExtension === ".svg" ? svg : png);
  }
  touch(session, false);
  return { svg, png, metadata, ...(outputPath ? { outputPath } : {}) };
}
