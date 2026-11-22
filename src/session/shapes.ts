import type {
  ChartFill,
  ChartFillPatch,
  ChartLine,
  ChartLinePatch,
  DrawingAnchor,
  DrawingAnchorInput,
  DrawingPatch,
  Shape,
  ShapePatch,
} from "@libraz/formulon";
import { parseCellRef } from "../a1.js";
import {
  AnchorKind,
  ColorKind,
  findSheetIndex,
  StatusCode,
  sheetAt,
  statusError,
} from "../formulon.js";
import { formatColor, parseColor } from "../styles/input.js";
import { describePlacement, emuSize, type ImageLength, KIND_NAMES, lengthToEmu } from "./images.js";
import { assertInGrid, sheetName } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

const EMU_PER_POINT = 12_700;

/** A solid `#RRGGBB` colour or `none`; shapes carry no alpha. */
export type ShapePaintInput = string;

export type ShapeFields = {
  geometry?: string;
  text?: string;
  textBox?: boolean;
  name?: string;
  fill?: ShapePaintInput;
  line?: ShapePaintInput;
  lineWidth?: ImageLength;
  rotation?: number;
  hidden?: boolean;
};

export type ShapePlacementInput = {
  anchor?: string;
  to?: string;
  offsetX?: ImageLength;
  offsetY?: ImageLength;
  width?: ImageLength;
  height?: ImageLength;
};

export type ShapeInsertInput = ShapeFields & ShapePlacementInput & { anchor: string };

type Session = ReturnType<typeof getSession>;
type ShapeFieldsSpec = Pick<
  ShapePatch,
  "presetGeometry" | "text" | "textBox" | "fill" | "line" | "rotationDeg"
> &
  Pick<DrawingPatch, "name" | "hidden">;

function sheetCell(ref: string, label: string) {
  const cell = parseCellRef(ref);
  if (cell.sheetName !== undefined) {
    throw statusError(
      StatusCode.InvalidArgument,
      `${label} must be a cell on the selected sheet, without a sheet name`,
    );
  }
  assertInGrid(cell.row, cell.col);
  return cell;
}

function paint(spec: ShapePaintInput, label: string): ChartFillPatch {
  if (spec.trim().toLowerCase() === "none") {
    return { none: true };
  }
  const argb = parseColor(spec);
  if (argb >>> 24 !== 0xff) {
    throw statusError(
      StatusCode.InvalidArgument,
      `${label} must be an opaque #RRGGBB colour or none: ${spec}`,
    );
  }
  return { color: formatColor(argb), none: false };
}

function describePaint(paintState: ChartFill | ChartLine | null): string | undefined {
  if (paintState === null) {
    return undefined;
  }
  if (paintState.none) {
    return "none";
  }
  if (paintState.color?.base.kind === ColorKind.Rgb) {
    return formatColor(paintState.color.base.argb);
  }
  return paintState.other || paintState.color ? "other" : undefined;
}

function describeShape(shape: Shape) {
  const properties = shape.getProperties();
  const base = shape.getDrawingProperties();
  const fill = describePaint(properties.fill);
  const line = describePaint(properties.line);
  const widthPt = properties.line?.widthPt;
  return {
    objectId: base.objectId,
    kind: KIND_NAMES[base.kind] ?? "other",
    name: base.name,
    geometry: properties.presetGeometry,
    ...(properties.text ? { text: properties.text } : {}),
    ...(properties.textBox ? { textBox: true } : {}),
    ...describePlacement(base.anchor),
    ...(fill ? { fill } : {}),
    ...(line ? { line } : {}),
    ...(widthPt ? { lineWidth: emuSize(widthPt * EMU_PER_POINT) } : {}),
    ...(properties.rotationDeg ? { rotation: properties.rotationDeg } : {}),
    ...(base.hidden ? { hidden: true } : {}),
  };
}

function fieldsSpec(fields: ShapeFields): ShapeFieldsSpec {
  const spec: { -readonly [K in keyof ShapeFieldsSpec]: ShapeFieldsSpec[K] } = {};
  if (fields.geometry !== undefined) {
    spec.presetGeometry = fields.geometry;
  }
  if (fields.text !== undefined) {
    spec.text = fields.text;
  }
  if (fields.textBox !== undefined) {
    spec.textBox = fields.textBox;
  }
  if (fields.name !== undefined) {
    spec.name = fields.name;
  }
  if (fields.fill !== undefined) {
    spec.fill = paint(fields.fill, "fill");
  }
  if (fields.line !== undefined || fields.lineWidth !== undefined) {
    const line: { -readonly [K in keyof ChartLinePatch]: ChartLinePatch[K] } =
      fields.line === undefined ? {} : paint(fields.line, "line");
    if (fields.lineWidth !== undefined) {
      line.widthPt = lengthToEmu(fields.lineWidth, "lineWidth") / EMU_PER_POINT;
    }
    spec.line = line;
  }
  if (fields.rotation !== undefined) {
    if (!Number.isFinite(fields.rotation)) {
      throw statusError(
        StatusCode.InvalidArgument,
        `rotation must be a finite number of degrees: ${fields.rotation}`,
      );
    }
    spec.rotationDeg = ((fields.rotation % 360) + 360) % 360;
  }
  if (fields.hidden !== undefined) {
    spec.hidden = fields.hidden;
  }
  return spec;
}

function hasPlacement(input: ShapePlacementInput): boolean {
  return (
    input.anchor !== undefined ||
    input.to !== undefined ||
    input.offsetX !== undefined ||
    input.offsetY !== undefined ||
    input.width !== undefined ||
    input.height !== undefined
  );
}

/**
 * Builds the whole anchor the engine needs from the stated placement, falling
 * back to `current` for what is left out. A `to` cell makes a two-cell box;
 * otherwise the shape is one-cell with an explicit size.
 */
function anchorSpec(input: ShapePlacementInput, current?: DrawingAnchor): DrawingAnchorInput {
  const kind = current
    ? current.kind
    : input.to === undefined
      ? AnchorKind.OneCell
      : AnchorKind.TwoCell;
  if (kind === AnchorKind.Absolute) {
    throw statusError(
      StatusCode.InvalidArgument,
      "an absolutely anchored shape cannot be moved by cell",
    );
  }
  const from = input.anchor === undefined ? undefined : sheetCell(input.anchor, "anchor");
  const offsets = {
    fromRow: from?.row ?? current?.fromRow ?? 0,
    fromCol: from?.col ?? current?.fromCol ?? 0,
    fromColOffsetEmu:
      input.offsetX === undefined
        ? (current?.fromColOffsetEmu ?? 0)
        : lengthToEmu(input.offsetX, "offsetX"),
    fromRowOffsetEmu:
      input.offsetY === undefined
        ? (current?.fromRowOffsetEmu ?? 0)
        : lengthToEmu(input.offsetY, "offsetY"),
  };
  if (kind === AnchorKind.TwoCell) {
    if (input.width !== undefined || input.height !== undefined) {
      throw statusError(
        StatusCode.InvalidArgument,
        "a two-cell shape is sized by its to cell; width and height do not apply",
      );
    }
    const to = input.to === undefined ? undefined : sheetCell(input.to, "to");
    const toRow = to?.row ?? current?.toRow ?? 0;
    const toCol = to?.col ?? current?.toCol ?? 0;
    if (toRow < offsets.fromRow || toCol < offsets.fromCol) {
      throw statusError(
        StatusCode.InvalidArgument,
        "to must not lie above or left of the anchor cell",
      );
    }
    return {
      kind: AnchorKind.TwoCell,
      ...offsets,
      toRow,
      toCol,
      toRowOffsetEmu: to ? 0 : (current?.toRowOffsetEmu ?? 0),
      toColOffsetEmu: to ? 0 : (current?.toColOffsetEmu ?? 0),
    };
  }
  if (input.to !== undefined) {
    throw statusError(
      StatusCode.InvalidArgument,
      "a one-cell shape keeps its anchor kind; use width and height instead of to",
    );
  }
  const widthEmu =
    input.width === undefined ? current?.widthEmu : lengthToEmu(input.width, "width");
  const heightEmu =
    input.height === undefined ? current?.heightEmu : lengthToEmu(input.height, "height");
  if (!(widthEmu && heightEmu)) {
    throw statusError(
      StatusCode.InvalidArgument,
      "a one-cell shape needs a positive width and height, or a to cell",
    );
  }
  return { kind: AnchorKind.OneCell, ...offsets, widthEmu, heightEmu };
}

function currentShape(session: Session, sheetIndex: number, objectId: number): Shape | undefined {
  return sheetAt(session.workbook, sheetIndex)
    .listShapes()
    .find((entry) => entry.objectId === objectId);
}

/** Lists a sheet's shapes, text boxes, connectors and groups in z-order, back first. */
export function listSessionShapes(id: string, sheet: number | string | undefined) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const shapes = sheetAt(session.workbook, sheetIndex).listShapes().map(describeShape);
  touch(session, false);
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    sheetName: sheetName(session, sheetIndex),
    count: shapes.length,
    shapes,
  };
}

/** Inserts a shape on top of the sheet's drawing, keeping everything already there. */
export function insertSessionShape(
  id: string,
  sheet: number | string | undefined,
  input: ShapeInsertInput,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const spec = { ...fieldsSpec(input), anchor: anchorSpec(input) };
  touch(session, true);
  const shape = sheetAt(session.workbook, sheetIndex).addShape(spec);
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    objectId: shape.objectId,
    placed: describeShape(shape),
  };
}

/** Changes only the stated fields of a shape; a placement change keeps the anchor kind. */
export function updateSessionShape(
  id: string,
  sheet: number | string | undefined,
  objectId: number,
  input: ShapeFields & ShapePlacementInput,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const current = currentShape(session, sheetIndex, objectId);
  if (!current) {
    throw statusError(
      StatusCode.NotFound,
      "no such shape",
      `object_id=${objectId} sheet_index=${sheetIndex}`,
    );
  }
  const spec = {
    ...fieldsSpec(input),
    ...(hasPlacement(input)
      ? { anchor: anchorSpec(input, current.getDrawingProperties().anchor) }
      : {}),
  };
  if (Object.keys(spec).length === 0) {
    throw statusError(
      StatusCode.InvalidArgument,
      "shape update requires at least one field or placement change",
    );
  }
  touch(session, true);
  current.update(spec);
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    objectId,
    placed: describeShape(current),
  };
}

/** Removes a shape, connector or group together with its children. */
export function removeSessionShape(
  id: string,
  sheet: number | string | undefined,
  objectId: number,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const sheetHandle = sheetAt(session.workbook, sheetIndex);
  const shape = sheetHandle.shape(objectId);
  touch(session, true);
  sheetHandle.removeDrawing(shape);
  touch(session, true);
  return { session: publicInfo(session), sheet: sheetIndex, objectId, removed: true };
}
