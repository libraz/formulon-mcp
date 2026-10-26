import type { ShapeInput } from "@libraz/formulon";
import { parseCellRef } from "../a1.js";
import {
  AnchorEditAs,
  AnchorKind,
  assertStatus,
  findSheetIndex,
  requireListResult,
  ShapePaint,
} from "../formulon.js";
import { formatColor, parseColor } from "../styles/input.js";
import { describePlacement, emuSize, type ImageLength, KIND_NAMES, lengthToEmu } from "./images.js";
import { assertInGrid, sheetName } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

/** DrawingML stores rotation in 60000ths of a degree. */
const ROTATION_UNITS_PER_DEGREE = 60_000;

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
type CurrentShape = ReturnType<Session["workbook"]["listShapes"]>[number];

function sheetCell(ref: string, label: string) {
  const cell = parseCellRef(ref);
  if (cell.sheetName !== undefined) {
    throw new Error(`${label} must be a cell on the selected sheet, without a sheet name`);
  }
  assertInGrid(cell.row, cell.col);
  return cell;
}

function paint(spec: ShapePaintInput, label: string): { kind: ShapePaint; rgb: number } {
  if (spec.trim().toLowerCase() === "none") {
    return { kind: ShapePaint.None, rgb: 0 };
  }
  const argb = parseColor(spec);
  if (argb >>> 24 !== 0xff) {
    throw new Error(`${label} must be an opaque #RRGGBB colour or none: ${spec}`);
  }
  return { kind: ShapePaint.Solid, rgb: argb & 0xffffff };
}

function describePaint(kind: ShapePaint, rgb: number): string | undefined {
  if (kind === ShapePaint.Solid) {
    return formatColor((rgb | 0xff000000) >>> 0);
  }
  return kind === ShapePaint.None ? "none" : kind === ShapePaint.Other ? "other" : undefined;
}

function describeShape(shape: CurrentShape) {
  const fill = describePaint(shape.fillKind, shape.fillRgb);
  const line = describePaint(shape.lineKind, shape.lineRgb);
  return {
    objectId: shape.objectId,
    kind: KIND_NAMES[shape.kind] ?? "other",
    name: shape.name,
    geometry: shape.prst,
    ...(shape.text ? { text: shape.text } : {}),
    ...(shape.textBox ? { textBox: true } : {}),
    ...describePlacement(shape),
    ...(fill ? { fill } : {}),
    ...(line ? { line } : {}),
    ...(shape.lineWidthEmu ? { lineWidth: emuSize(shape.lineWidthEmu) } : {}),
    ...(shape.rotation ? { rotation: shape.rotation / ROTATION_UNITS_PER_DEGREE } : {}),
    ...(shape.hidden ? { hidden: true } : {}),
  };
}

function currentShapes(session: Session, sheetIndex: number): CurrentShape[] {
  return requireListResult(session.workbook.listShapes(sheetIndex), "list shapes");
}

function fieldsSpec(fields: ShapeFields): ShapeInput {
  const spec: ShapeInput = {};
  if (fields.geometry !== undefined) {
    spec.prst = fields.geometry;
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
    const { kind, rgb } = paint(fields.fill, "fill");
    spec.fillKind = kind;
    spec.fillRgb = rgb;
  }
  if (fields.line !== undefined) {
    const { kind, rgb } = paint(fields.line, "line");
    spec.lineKind = kind;
    spec.lineRgb = rgb;
  }
  if (fields.lineWidth !== undefined) {
    spec.lineWidthEmu = lengthToEmu(fields.lineWidth, "lineWidth");
  }
  if (fields.rotation !== undefined) {
    if (!Number.isFinite(fields.rotation)) {
      throw new Error(`rotation must be a finite number of degrees: ${fields.rotation}`);
    }
    const degrees = ((fields.rotation % 360) + 360) % 360;
    spec.rotation = Math.round(degrees * ROTATION_UNITS_PER_DEGREE);
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
function anchorSpec(input: ShapePlacementInput, current?: CurrentShape): ShapeInput {
  const kind = current
    ? current.anchorKind
    : input.to === undefined
      ? AnchorKind.OneCell
      : AnchorKind.TwoCell;
  if (kind === AnchorKind.Absolute) {
    throw new Error("an absolutely anchored shape cannot be moved by cell");
  }
  const from = input.anchor === undefined ? undefined : sheetCell(input.anchor, "anchor");
  const offsets = {
    fromRow: from?.row ?? current?.fromRow ?? 0,
    fromCol: from?.col ?? current?.fromCol ?? 0,
    fromColOff:
      input.offsetX === undefined
        ? (current?.fromColOff ?? 0)
        : lengthToEmu(input.offsetX, "offsetX"),
    fromRowOff:
      input.offsetY === undefined
        ? (current?.fromRowOff ?? 0)
        : lengthToEmu(input.offsetY, "offsetY"),
  };
  if (kind === AnchorKind.TwoCell) {
    if (input.width !== undefined || input.height !== undefined) {
      throw new Error("a two-cell shape is sized by its to cell; width and height do not apply");
    }
    const to = input.to === undefined ? undefined : sheetCell(input.to, "to");
    const toRow = to?.row ?? current?.toRow ?? 0;
    const toCol = to?.col ?? current?.toCol ?? 0;
    if (toRow < offsets.fromRow || toCol < offsets.fromCol) {
      throw new Error("to must not lie above or left of the anchor cell");
    }
    return {
      anchorKind: AnchorKind.TwoCell,
      editAs: current?.editAs ?? AnchorEditAs.TwoCell,
      ...offsets,
      toRow,
      toCol,
      toRowOff: to ? 0 : (current?.toRowOff ?? 0),
      toColOff: to ? 0 : (current?.toColOff ?? 0),
    };
  }
  if (input.to !== undefined) {
    throw new Error("a one-cell shape keeps its anchor kind; use width and height instead of to");
  }
  const cx = input.width === undefined ? current?.cx : lengthToEmu(input.width, "width");
  const cy = input.height === undefined ? current?.cy : lengthToEmu(input.height, "height");
  if (!(cx && cy)) {
    throw new Error("a one-cell shape needs a positive width and height, or a to cell");
  }
  return {
    anchorKind: AnchorKind.OneCell,
    editAs: current?.editAs ?? AnchorEditAs.OneCell,
    ...offsets,
    cx,
    cy,
  };
}

function placedShape(session: Session, sheetIndex: number, objectId: number) {
  const shape = currentShapes(session, sheetIndex).find((entry) => entry.objectId === objectId);
  return shape ? describeShape(shape) : undefined;
}

/** Lists a sheet's shapes, text boxes, connectors and groups in z-order, back first. */
export function listSessionShapes(id: string, sheet: number | string | undefined) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const shapes = currentShapes(session, sheetIndex).map(describeShape);
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
  const spec = { ...fieldsSpec(input), ...anchorSpec(input) };
  touch(session, true);
  const result = session.workbook.insertShape(sheetIndex, spec);
  assertStatus(result.status, "insert shape");
  const placed = placedShape(session, sheetIndex, result.objectId);
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    objectId: result.objectId,
    ...(placed ? { placed } : {}),
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
  const current = currentShapes(session, sheetIndex).find((entry) => entry.objectId === objectId);
  if (!current) {
    throw new Error(`no shape ${objectId} on sheet ${sheetIndex}`);
  }
  const spec = {
    ...fieldsSpec(input),
    ...(hasPlacement(input) ? anchorSpec(input, current) : {}),
  };
  if (Object.keys(spec).length === 0) {
    throw new Error("shape update requires at least one field or placement change");
  }
  touch(session, true);
  assertStatus(session.workbook.updateShape(sheetIndex, objectId, spec), "update shape");
  const placed = placedShape(session, sheetIndex, objectId);
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    objectId,
    ...(placed ? { placed } : {}),
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
  touch(session, true);
  assertStatus(session.workbook.removeShape(sheetIndex, objectId), "remove shape");
  touch(session, true);
  return { session: publicInfo(session), sheet: sheetIndex, objectId, removed: true };
}
