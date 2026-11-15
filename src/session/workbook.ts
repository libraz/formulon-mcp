import { parseCellRef } from "../a1.js";
import {
  findSheetIndex,
  formulonModule,
  GeometryMode,
  okStatus,
  type ProfileId,
  resultToJson,
  SheetVisibility,
  StatusCode,
  statusError,
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
  refA1,
  requiredBoolean,
  requiredNumber,
  requiredString,
  sheetName,
  spillRecord,
} from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

/** Recalculates a session workbook. */
export function recalcSession(id: string) {
  const session = getSession(id);
  touch(session, true);
  session.workbook.recalc();
  const circularReferences = session.workbook
    .listSheets()
    .flatMap((handle) =>
      handle.listCircularReferences().map((cell) => `${handle.name}!${refA1(cell.row, cell.col)}`),
    );
  return { session: publicInfo(session), status: okStatus(), circularReferences };
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
  if (operation === "add") {
    const name = requiredString(args.name, "name");
    touch(session, true);
    workbook.addSheet(name);
  } else if (operation === "remove") {
    const index = requiredNumber(args.index, "index");
    touch(session, true);
    workbook.removeSheet(index);
  } else if (operation === "rename") {
    const index = requiredNumber(args.index, "index");
    const newName = requiredString(args.newName, "newName");
    touch(session, true);
    workbook.sheet(index).rename(newName);
  } else {
    const fromIndex = requiredNumber(args.fromIndex, "fromIndex");
    const toIndex = requiredNumber(args.toIndex, "toIndex");
    touch(session, true);
    // `toIndex` is the final position; `move` takes the position in the list as it is now.
    workbook.sheet(fromIndex).move(toIndex > fromIndex ? toIndex + 1 : toIndex);
  }
  touch(session, true);
  return { session: publicInfo(session), status: okStatus() };
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
  attrs: { comment?: string; hidden?: boolean } = {},
) {
  const session = getSession(id);
  // OOXML <definedName> content is a bare expression with no leading '='.
  // Strip one if the caller supplied it so the saved file is Excel-conformant.
  const trimmed = formula.trim();
  const refersTo = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const hasAttrs = attrs.comment !== undefined || attrs.hidden !== undefined;
  if (refersTo === "" && hasAttrs) {
    throw statusError(
      StatusCode.InvalidArgument,
      "comment and hidden cannot be set while removing a defined name",
    );
  }
  const localSheetId =
    sheet === undefined ? WORKBOOK_SCOPE : findSheetIndex(session.workbook, sheet);
  touch(session, true);
  const scope = localSheetId === WORKBOOK_SCOPE ? null : localSheetId;
  if (refersTo === "") {
    session.workbook.removeDefinedName(name, scope);
  } else {
    // The engine replaces every attribute of the name, so an omitted one keeps its stored value.
    const existing = session.workbook.getDefinedName(name, scope);
    const comment = attrs.comment ?? existing?.comment ?? undefined;
    const hidden = attrs.hidden ?? existing?.hidden;
    session.workbook.setDefinedName({
      name,
      formula: refersTo,
      scope,
      ...(comment !== undefined ? { comment } : {}),
      ...(hidden !== undefined ? { hidden } : {}),
    });
  }
  touch(session, true);
  return { session: publicInfo(session), status: okStatus(), localSheetId };
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
  const handle = session.workbook.sheet(sheetIndex);
  touch(session, true);
  if (operation === "insertRows") {
    handle.insertRows(start, start + count - 1);
  } else if (operation === "deleteRows") {
    handle.deleteRows(start, start + count - 1);
  } else if (operation === "insertCols") {
    handle.insertColumns(start, start + count - 1);
  } else {
    handle.deleteColumns(start, start + count - 1);
  }
  touch(session, true);
  return { session: publicInfo(session), sheet: sheetIndex, status: okStatus() };
}

export type SheetVisibilityName = "visible" | "hidden" | "veryHidden";

/** `<pane xSplit ySplit>` is stored in twips. */
const TWIPS_PER_POINT = 20;

const VISIBILITY_CODES: Record<SheetVisibilityName, SheetVisibility> = {
  visible: SheetVisibility.Visible,
  hidden: SheetVisibility.Hidden,
  veryHidden: SheetVisibility.VeryHidden,
};

/** The name of an engine sheet visibility code. */
export function sheetVisibilityName(code: SheetVisibility): SheetVisibilityName {
  return (
    (Object.keys(VISIBILITY_CODES) as SheetVisibilityName[]).find(
      (name) => VISIBILITY_CODES[name] === code,
    ) ?? "visible"
  );
}

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
    splitX?: number;
    splitY?: number;
    topLeft?: string;
    hidden?: boolean;
    visibility?: SheetVisibilityName;
    showFormulas?: boolean;
    showOutlineSymbols?: boolean;
    active?: true;
  },
) {
  const session = getSession(id);
  const { workbook } = session;
  const sheetIndex = findSheetIndex(workbook, sheet);
  if (
    (options.freezeRows !== undefined || options.freezeCols !== undefined) &&
    (options.splitX !== undefined || options.splitY !== undefined)
  ) {
    throw statusError(
      StatusCode.InvalidArgument,
      "a view is either frozen or split; pass freezeRows/freezeCols or splitX/splitY",
    );
  }
  const topLeft = options.topLeft === undefined ? undefined : parseCellRef(options.topLeft);
  if (topLeft?.sheetName !== undefined) {
    throw statusError(
      StatusCode.InvalidArgument,
      `topLeft is a cell on the selected sheet, without a sheet name: ${options.topLeft}`,
    );
  }
  if (topLeft) {
    assertInGrid(topLeft.row, topLeft.col);
  }
  const handle = workbook.sheet(sheetIndex);
  // Visibility applies before activation, so one call can unhide and activate a sheet.
  const steps: [boolean, () => void][] = [
    [options.zoom !== undefined, () => handle.updateView({ zoomPct: options.zoom ?? 100 })],
    [
      options.freezeRows !== undefined || options.freezeCols !== undefined,
      () =>
        handle.updateView({
          freezeRows: options.freezeRows ?? 0,
          freezeCols: options.freezeCols ?? 0,
        }),
    ],
    [
      options.splitX !== undefined || options.splitY !== undefined,
      () =>
        handle.updateView({
          splitXTwips: Math.round((options.splitX ?? 0) * TWIPS_PER_POINT),
          splitYTwips: Math.round((options.splitY ?? 0) * TWIPS_PER_POINT),
        }),
    ],
    [
      topLeft !== undefined,
      () => handle.updateView({ topLeft: { row: topLeft?.row ?? 0, col: topLeft?.col ?? 0 } }),
    ],
    [
      options.hidden !== undefined,
      () => handle.setVisibility(options.hidden ? SheetVisibility.Hidden : SheetVisibility.Visible),
    ],
    [
      options.visibility !== undefined,
      () => handle.setVisibility(VISIBILITY_CODES[options.visibility ?? "visible"]),
    ],
    [
      options.showFormulas !== undefined,
      () => handle.updateView({ showFormulas: options.showFormulas ?? false }),
    ],
    [
      options.showOutlineSymbols !== undefined,
      () => handle.updateView({ showOutlineSymbols: options.showOutlineSymbols ?? true }),
    ],
    [options.active === true, () => workbook.setActiveSheet(sheetIndex)],
  ];
  const pending = steps.filter(([wanted]) => wanted);
  if (pending.length > 0) {
    touch(session, true);
  }
  const statuses = pending.map(([, apply]) => {
    apply();
    return okStatus();
  });
  touch(session, statuses.length > 0);
  return { session: publicInfo(session), sheet: sheetIndex, statuses };
}

/** Reads broad workbook metadata that does not require mutation. */
export async function getSessionMetadata(id: string, kind: "functions" | "externalLinks") {
  const session = getSession(id);
  touch(session, false);
  const value =
    kind === "functions"
      ? (await formulonModule()).listFunctionNames()
      : session.workbook.listExternalLinks();
  return {
    session: publicInfo(session),
    kind,
    value: resultToJson(value),
  };
}

/** Reads the Excel profile a session evaluates with. */
export function sessionProfileId(id: string): ProfileId {
  return getSession(id).workbook.getProfileId();
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
    throw statusError(
      StatusCode.InvalidArgument,
      `trace depth must be an integer in 1..32: ${depth}`,
    );
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
    return {
      session: publicInfo(session),
      operation,
      cell: cellRef(sheetIndex, row, col),
      spill: spillRecord(session, sheetIndex, row, col),
    };
  }
  const handle = session.workbook.sheet(sheetIndex);
  const nodes =
    operation === "precedents"
      ? handle.getPrecedents(row, col, depth)
      : handle.getDependents(row, col, depth);
  const cells = nodes.map((node) => cellRef(node.sheet, node.row, node.col));
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
export type DimensionOperation = "list" | "size" | "resetSize" | "hidden" | "outline" | "collapsed";
export type DimensionUnit = "chars" | "pt" | "px" | "mm";

type DisplayGeometry = {
  points: number;
  pixels: number;
  mm: number;
};

function dimensionSize(value: number | undefined, name: string): number {
  const size = requiredNumber(value, name);
  if (size < 0) {
    throw statusError(StatusCode.InvalidArgument, `${name} must be nonnegative`);
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
    throw statusError(StatusCode.InvalidArgument, "column size unit must be chars, pt, px, or mm");
  }
  if (unit === "chars") {
    if (size > MAX_LAYOUT_COLUMN_WIDTH_CHARS) {
      throw statusError(
        StatusCode.InvalidArgument,
        `column width exceeds the ${MAX_LAYOUT_COLUMN_WIDTH_CHARS}-character Excel limit: ${size}`,
      );
    }
    return size;
  }
  const points = unit === "pt" ? size : unit === "px" ? pixelsToPoints(size) : mmToPoints(size);
  const stored = columnPointsToChars(workbook, sheet, points);
  if (stored > MAX_LAYOUT_COLUMN_WIDTH_CHARS) {
    throw statusError(
      StatusCode.InvalidArgument,
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
  const handle = workbook.sheet(sheet);
  if (axis === "column") {
    const columns = handle.listColumnOverrides().map((column) => {
      if (!includeGeometry) {
        return { ...column };
      }
      const points = handle.getColumnWidthPt(column.first, GeometryMode.Display);
      return { ...column, displayGeometry: displayGeometry(points) };
    });
    return { status: okStatus(), columns };
  }
  const rows = handle.listRowOverrides().map((row) => {
    if (!includeGeometry) {
      return { ...row };
    }
    return { ...row, displayGeometry: displayGeometry(handle.getRowHeightPt(row.row)) };
  });
  return { status: okStatus(), rows };
}

/**
 * Reads or sets column-width / row-height, hidden, outline-level and collapsed overrides.
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
    collapsed?: boolean;
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

  const handle = workbook.sheet(sheetIndex);
  if (axis === "column") {
    const first = requiredNumber(params.first, "first");
    const last = params.last ?? first;
    assertInGrid(0, first);
    assertInGrid(0, last);
    if (last < first) {
      throw statusError(StatusCode.InvalidArgument, `last (${last}) must be >= first (${first})`);
    }
    if (operation === "size") {
      const size = dimensionSize(params.size, "size");
      const unit = params.unit ?? "chars";
      const storedSize = columnStoredSize(workbook, sheetIndex, size, unit);
      touch(session, true);
      handle.setColumnWidth(first, last, storedSize);
      touch(session, true);
      return {
        session: publicInfo(session),
        axis,
        operation,
        sheet: sheetIndex,
        status: okStatus(),
        storedUnits: "chars",
        storedSize,
        displayGeometry: displayGeometry(columnCharsToPoints(workbook, sheetIndex, storedSize)),
      };
    } else if (operation === "resetSize") {
      touch(session, true);
      handle.resetColumnWidth(first, last);
    } else if (operation === "hidden") {
      const hidden = requiredBoolean(params.hidden, "hidden");
      touch(session, true);
      handle.setColumnHidden(first, last, hidden);
    } else if (operation === "collapsed") {
      const collapsed = requiredBoolean(params.collapsed, "collapsed");
      touch(session, true);
      handle.setColumnCollapsed(first, last, collapsed);
    } else {
      const level = requiredNumber(params.level, "level");
      touch(session, true);
      handle.setColumnOutline(first, last, level);
    }
  } else {
    const rowIndex = requiredNumber(params.row, "row");
    assertInGrid(rowIndex, 0);
    if (operation === "size") {
      const size = dimensionSize(params.size, "size");
      const unit = params.unit ?? "pt";
      if (!["pt", "px", "mm"].includes(unit)) {
        throw statusError(
          StatusCode.InvalidArgument,
          "row size unit must be pt, px, or mm; chars is unsupported",
        );
      }
      const storedSize =
        unit === "pt" ? size : unit === "px" ? pixelsToPoints(size) : mmToPoints(size);
      if (storedSize > MAX_LAYOUT_ROW_HEIGHT_PT) {
        throw statusError(
          StatusCode.InvalidArgument,
          `row height exceeds the ${MAX_LAYOUT_ROW_HEIGHT_PT}pt Excel limit: ${storedSize}`,
        );
      }
      touch(session, true);
      handle.setRowHeight(rowIndex, rowIndex, storedSize);
      touch(session, true);
      return {
        session: publicInfo(session),
        axis,
        operation,
        sheet: sheetIndex,
        status: okStatus(),
        storedUnits: "pt",
        storedSize,
        displayGeometry: displayGeometry(storedSize),
      };
    } else if (operation === "resetSize") {
      touch(session, true);
      handle.resetRowHeight(rowIndex, rowIndex);
    } else if (operation === "hidden") {
      const hidden = requiredBoolean(params.hidden, "hidden");
      touch(session, true);
      handle.setRowHidden(rowIndex, rowIndex, hidden);
    } else if (operation === "collapsed") {
      const collapsed = requiredBoolean(params.collapsed, "collapsed");
      touch(session, true);
      handle.setRowCollapsed(rowIndex, rowIndex, collapsed);
    } else {
      const level = requiredNumber(params.level, "level");
      touch(session, true);
      handle.setRowOutline(rowIndex, rowIndex, level);
    }
  }
  touch(session, true);
  return {
    session: publicInfo(session),
    axis,
    operation,
    sheet: sheetIndex,
    status: okStatus(),
  };
}
