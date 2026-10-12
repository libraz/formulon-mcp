import {
  assertStatus,
  type ExcelProfileId,
  findSheetIndex,
  GeometryMode,
  requireListResult,
  requireNumberResult,
  resultToJson,
  SheetVisibility,
  type Status,
  statusToJson,
  WORKBOOK_SCOPE,
  type Workbook,
} from "../formulon.js";
import { MAX_LAYOUT_COLUMN_WIDTH_CHARS, MAX_LAYOUT_ROW_HEIGHT_PT } from "../layout/types.js";
import {
  columnCharsToPoints,
  columnPointsToChars,
  mmToPoints,
  pixelsToPoints,
  pointsToMm,
  pointsToPixels,
} from "./geometry.js";
import {
  assertInGrid,
  rangeA1,
  refA1,
  requiredBoolean,
  requiredNumber,
  requiredString,
  sheetName,
} from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

/** Recalculates a session workbook. */
export function recalcSession(id: string) {
  const session = getSession(id);
  touch(session, true);
  const status = session.workbook.recalc();
  assertStatus(status, "recalc workbook");
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
  let status: Status;
  if (operation === "add") {
    const name = requiredString(args.name, "name");
    touch(session, true);
    status = workbook.addSheet(name);
  } else if (operation === "remove") {
    const index = requiredNumber(args.index, "index");
    touch(session, true);
    status = workbook.removeSheet(index);
  } else if (operation === "rename") {
    const index = requiredNumber(args.index, "index");
    const newName = requiredString(args.newName, "newName");
    touch(session, true);
    status = workbook.renameSheet(index, newName);
  } else {
    const fromIndex = requiredNumber(args.fromIndex, "fromIndex");
    const toIndex = requiredNumber(args.toIndex, "toIndex");
    touch(session, true);
    status = workbook.moveSheet(fromIndex, toIndex);
  }
  assertStatus(status, `${operation} sheet`);
  touch(session, true);
  return { session: publicInfo(session), status: statusToJson(status) };
}

/**
 * Adds, replaces, or removes a defined name, scoped to one sheet when `sheet`
 * is given and to the whole workbook otherwise. Scope is not cosmetic: Excel
 * only honours `_xlnm.Print_Area` and `_xlnm.Print_Titles` when they are
 * sheet-local, and writing them in workbook scope fails silently — the file
 * opens, the name is there, and the print range simply does not apply.
 */
export function setSessionDefinedName(
  id: string,
  name: string,
  formula: string,
  sheet?: number | string,
) {
  const session = getSession(id);
  // OOXML <definedName> content is a bare expression with no leading '='.
  // Strip one if the caller supplied it so the saved file is Excel-conformant.
  const trimmed = formula.trim();
  const refersTo = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const localSheetId =
    sheet === undefined ? WORKBOOK_SCOPE : findSheetIndex(session.workbook, sheet);
  touch(session, true);
  const status =
    localSheetId === WORKBOOK_SCOPE
      ? session.workbook.setDefinedName(name, refersTo)
      : session.workbook.setDefinedNameScoped(name, refersTo, localSheetId);
  assertStatus(status, "set defined name");
  touch(session, true);
  return { session: publicInfo(session), status: statusToJson(status), localSheetId };
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
  touch(session, true);
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

export type SheetVisibilityName = "visible" | "hidden" | "veryHidden";

const VISIBILITY_CODES: Record<SheetVisibilityName, SheetVisibility> = {
  visible: SheetVisibility.Visible,
  hidden: SheetVisibility.Hidden,
  veryHidden: SheetVisibility.VeryHidden,
};

/**
 * Updates sheet view settings such as zoom, frozen panes, and tab visibility.
 *
 * `hidden` is the two-state view of `visibility`: it cannot state very-hidden,
 * and `false` on a very-hidden sheet shows it. Pass `visibility` to reach the
 * third state, which is how a workbook keeps a lookup or settings sheet out of
 * Excel's "Unhide" dialog.
 */
export function setSessionSheetView(
  id: string,
  sheet: number | string | undefined,
  options: {
    zoom?: number;
    freezeRows?: number;
    freezeCols?: number;
    hidden?: boolean;
    visibility?: SheetVisibilityName;
  },
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const statuses = [];
  if (
    options.zoom !== undefined ||
    options.freezeRows !== undefined ||
    options.freezeCols !== undefined ||
    options.hidden !== undefined ||
    options.visibility !== undefined
  ) {
    touch(session, true);
  }
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
  if (options.visibility !== undefined) {
    const status = session.workbook.setSheetVisibility(
      sheetIndex,
      VISIBILITY_CODES[options.visibility],
    );
    assertStatus(status, "set sheet visibility");
    statuses.push(statusToJson(status));
  }
  touch(session, statuses.length > 0);
  return { session: publicInfo(session), sheet: sheetIndex, statuses };
}

/** Reads broad workbook metadata that does not require mutation. */
export function getSessionMetadata(id: string, kind: "functions" | "externalLinks") {
  const session = getSession(id);
  touch(session, false);
  const value =
    kind === "functions"
      ? requireListResult(session.workbook.functionNames(), "read function names")
      : requireListResult(session.workbook.getExternalLinks(), "read external links");
  return {
    session: publicInfo(session),
    kind,
    value: resultToJson(value),
  };
}

/** Reads the Excel profile a session evaluates with. */
export function sessionProfileId(id: string): ExcelProfileId {
  const result = getSession(id).workbook.excelProfileId();
  assertStatus(result.status, "read excel profile");
  return result.value;
}

/** Resolves a sheet reference (zero-based index or name) to an index for a session. */
export function resolveSessionSheet(id: string, sheet: number | string | undefined): number {
  return findSheetIndex(getSession(id).workbook, sheet);
}

/**
 * Traces cell precedents, dependents, or spill region, returning A1 references
 * (with sheet names) instead of raw zero-based indices so results are readable.
 */
export function traceSession(
  id: string,
  operation: "precedents" | "dependents" | "spillInfo",
  sheet: number | string | undefined,
  row: number,
  col: number,
  depth: number,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  assertInGrid(row, col);
  if (!Number.isInteger(depth) || depth < 1 || depth > 32) {
    throw new Error(`trace depth must be an integer in 1..32: ${depth}`);
  }
  const cellRef = (s: number, r: number, c: number) => ({
    sheet: s,
    sheetName: sheetName(session, s),
    row: r,
    col: c,
    ref: `${sheetName(session, s)}!${refA1(r, c)}`,
  });
  touch(session, false);
  if (operation === "spillInfo") {
    const info = session.workbook.spillInfo(sheetIndex, row, col);
    assertStatus(info.status, "trace spill info");
    return {
      session: publicInfo(session),
      operation,
      cell: cellRef(sheetIndex, row, col),
      spill: {
        engaged: info.engaged,
        anchor: info.engaged ? cellRef(sheetIndex, info.anchorRow, info.anchorCol) : null,
        rows: info.rows,
        cols: info.cols,
        range: info.engaged
          ? `${sheetName(session, sheetIndex)}!${rangeA1(
              info.anchorRow,
              info.anchorCol,
              info.anchorRow + info.rows - 1,
              info.anchorCol + info.cols - 1,
            )}`
          : null,
      },
    };
  }
  const nodes =
    operation === "precedents"
      ? requireListResult(
          session.workbook.precedents(sheetIndex, row, col, depth),
          "trace precedents",
        )
      : requireListResult(
          session.workbook.dependents(sheetIndex, row, col, depth),
          "trace dependents",
        );
  const cells = Array.from({ length: nodes.length }, (_, index) => {
    const node = nodes[index];
    return cellRef(node.sheet, node.row, node.col);
  });
  return {
    session: publicInfo(session),
    operation,
    cell: cellRef(sheetIndex, row, col),
    depth,
    count: cells.length,
    cells,
  };
}

export type DimensionAxis = "column" | "row";
export type DimensionOperation = "list" | "size" | "resetSize" | "hidden" | "outline";
export type DimensionUnit = "chars" | "pt" | "px" | "mm";

type DisplayGeometry = {
  points: number;
  pixels: number;
  mm: number;
};

function dimensionSize(value: number | undefined, name: string): number {
  const size = requiredNumber(value, name);
  if (size < 0) {
    throw new Error(`${name} must be nonnegative`);
  }
  return size;
}

function displayGeometry(points: number): DisplayGeometry {
  return {
    points,
    pixels: pointsToPixels(points),
    mm: pointsToMm(points),
  };
}

function columnStoredSize(workbook: Workbook, sheet: number, size: number, unit: DimensionUnit) {
  if (!["chars", "pt", "px", "mm"].includes(unit)) {
    throw new Error("column size unit must be chars, pt, px, or mm");
  }
  if (unit === "chars") {
    if (size > MAX_LAYOUT_COLUMN_WIDTH_CHARS) {
      throw new Error(
        `column width exceeds the ${MAX_LAYOUT_COLUMN_WIDTH_CHARS}-character Excel limit: ${size}`,
      );
    }
    return size;
  }
  const points = unit === "pt" ? size : unit === "px" ? pixelsToPoints(size) : mmToPoints(size);
  const stored = columnPointsToChars(workbook, sheet, points);
  if (stored > MAX_LAYOUT_COLUMN_WIDTH_CHARS) {
    throw new Error(
      `column width exceeds the ${MAX_LAYOUT_COLUMN_WIDTH_CHARS}-character Excel limit: ${stored}`,
    );
  }
  return stored;
}

function dimensionList(
  workbook: Workbook,
  sheet: number,
  axis: DimensionAxis,
  includeGeometry: boolean,
) {
  if (axis === "column") {
    const result = workbook.getSheetColumns(sheet);
    assertStatus(result.status, "list column dimensions");
    const columns = result.columns.map((column) => {
      if (!includeGeometry) {
        return { ...column };
      }
      const points = requireNumberResult(
        workbook.getColumnWidthPt(sheet, column.first, GeometryMode.Display),
        "read column width",
      );
      return { ...column, displayGeometry: displayGeometry(points) };
    });
    return { status: statusToJson(result.status), columns };
  }
  const result = workbook.getSheetRowOverrides(sheet);
  assertStatus(result.status, "list row dimensions");
  const rows = result.rows.map((row) => {
    if (!includeGeometry) {
      return { ...row };
    }
    const points = requireNumberResult(workbook.getRowHeightPt(sheet, row.row), "read row height");
    return { ...row, displayGeometry: displayGeometry(points) };
  });
  return { status: statusToJson(result.status), rows };
}

/**
 * Reads or sets column-width / row-height, hidden, and outline-level overrides.
 * Columns act on an inclusive `[first, last]` span; rows act on a single index.
 */
export function sessionDimension(
  id: string,
  axis: DimensionAxis,
  operation: DimensionOperation,
  params: {
    sheet?: number | string;
    first?: number;
    last?: number;
    row?: number;
    size?: number;
    hidden?: boolean;
    level?: number;
    unit?: DimensionUnit;
    includeGeometry?: boolean;
  },
) {
  const session = getSession(id);
  const workbook = session.workbook;
  const sheetIndex = findSheetIndex(workbook, params.sheet);
  if (operation === "list") {
    const value = dimensionList(workbook, sheetIndex, axis, params.includeGeometry ?? false);
    touch(session, false);
    return { session: publicInfo(session), axis, operation, sheet: sheetIndex, value };
  }

  let status: Status;
  if (axis === "column") {
    const first = requiredNumber(params.first, "first");
    const last = params.last ?? first;
    assertInGrid(0, first);
    assertInGrid(0, last);
    if (last < first) {
      throw new Error(`last (${last}) must be >= first (${first})`);
    }
    if (operation === "size") {
      const size = dimensionSize(params.size, "size");
      const unit = params.unit ?? "chars";
      const storedSize = columnStoredSize(workbook, sheetIndex, size, unit);
      touch(session, true);
      status = workbook.setColumnWidth(sheetIndex, first, last, storedSize);
      assertStatus(status, `${axis} ${operation}`);
      touch(session, true);
      return {
        session: publicInfo(session),
        axis,
        operation,
        sheet: sheetIndex,
        status: statusToJson(status),
        storedUnits: "chars",
        storedSize,
        displayGeometry: displayGeometry(columnCharsToPoints(workbook, sheetIndex, storedSize)),
      };
    } else if (operation === "resetSize") {
      touch(session, true);
      status = workbook.clearColumnWidth(sheetIndex, first, last);
    } else if (operation === "hidden") {
      const hidden = requiredBoolean(params.hidden);
      touch(session, true);
      status = workbook.setColumnHidden(sheetIndex, first, last, hidden);
    } else {
      const level = requiredNumber(params.level, "level");
      touch(session, true);
      status = workbook.setColumnOutline(sheetIndex, first, last, level);
    }
  } else {
    const rowIndex = requiredNumber(params.row, "row");
    assertInGrid(rowIndex, 0);
    if (operation === "size") {
      const size = dimensionSize(params.size, "size");
      const unit = params.unit ?? "pt";
      if (!["pt", "px", "mm"].includes(unit)) {
        throw new Error("row size unit must be pt, px, or mm; chars is unsupported");
      }
      const storedSize =
        unit === "pt" ? size : unit === "px" ? pixelsToPoints(size) : mmToPoints(size);
      if (storedSize > MAX_LAYOUT_ROW_HEIGHT_PT) {
        throw new Error(
          `row height exceeds the ${MAX_LAYOUT_ROW_HEIGHT_PT}pt Excel limit: ${storedSize}`,
        );
      }
      touch(session, true);
      status = workbook.setRowHeight(sheetIndex, rowIndex, storedSize);
      assertStatus(status, `${axis} ${operation}`);
      touch(session, true);
      return {
        session: publicInfo(session),
        axis,
        operation,
        sheet: sheetIndex,
        status: statusToJson(status),
        storedUnits: "pt",
        storedSize,
        displayGeometry: displayGeometry(storedSize),
      };
    } else if (operation === "resetSize") {
      touch(session, true);
      status = workbook.clearRowHeight(sheetIndex, rowIndex);
    } else if (operation === "hidden") {
      const hidden = requiredBoolean(params.hidden);
      touch(session, true);
      status = workbook.setRowHidden(sheetIndex, rowIndex, hidden);
    } else {
      const level = requiredNumber(params.level, "level");
      touch(session, true);
      status = workbook.setRowOutline(sheetIndex, rowIndex, level);
    }
  }
  assertStatus(status, `${axis} ${operation}`);
  touch(session, true);
  return {
    session: publicInfo(session),
    axis,
    operation,
    sheet: sheetIndex,
    status: statusToJson(status),
  };
}
