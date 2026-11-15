import { readFileSync, writeFileSync } from "node:fs";
import { type Drawing, type DrawingAnchor, DrawingObjectKind } from "@libraz/formulon";
import { parseCellRef } from "../a1.js";
import {
  AnchorEditAs,
  AnchorKind,
  findSheetIndex,
  formulonModule,
  ImageFormat,
  resolveUserPath,
  StatusCode,
  sheetAt,
  statusError,
} from "../formulon.js";
import { mmToPoints, pixelsToPoints, pointsToPixels } from "./geometry.js";
import { assertInGrid, refA1, sheetName } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

const EMU_PER_POINT = 12_700;
const MAX_IMAGE_BYTES = 20 * 1024 * 1024;

export type ImageLength = { size: number; unit?: "pt" | "px" | "mm" };

export type ImageInsertInput = {
  path: string;
  anchor: string;
  offsetX?: ImageLength;
  offsetY?: ImageLength;
  width?: ImageLength;
  height?: ImageLength;
  anchorKind?: "oneCell" | "twoCell";
  name?: string;
  description?: string;
};

const FORMAT_NAMES: Record<number, string> = {
  [ImageFormat.Unknown]: "unknown",
  [ImageFormat.Png]: "png",
  [ImageFormat.Jpeg]: "jpeg",
  [ImageFormat.Gif]: "gif",
  [ImageFormat.Bmp]: "bmp",
};

export const KIND_NAMES: Record<number, string> = {
  [DrawingObjectKind.Picture]: "picture",
  [DrawingObjectKind.Shape]: "shape",
  [DrawingObjectKind.Chart]: "chart",
  [DrawingObjectKind.Group]: "group",
  [DrawingObjectKind.Connector]: "connector",
  [DrawingObjectKind.GraphicFrame]: "graphicFrame",
  [DrawingObjectKind.Other]: "other",
};

const ANCHOR_NAMES: Record<number, string> = {
  [AnchorKind.OneCell]: "oneCell",
  [AnchorKind.TwoCell]: "twoCell",
  [AnchorKind.Absolute]: "absolute",
};

export function imageFormatName(format: number): string {
  return FORMAT_NAMES[format] ?? "unknown";
}

export function imageMimeType(format: number): string | undefined {
  const name = FORMAT_NAMES[format];
  return name && name !== "unknown" ? `image/${name}` : undefined;
}

export function lengthToEmu(length: ImageLength | undefined, label: string): number {
  if (length === undefined) {
    return 0;
  }
  const unit = length.unit ?? "pt";
  const points =
    unit === "px"
      ? pixelsToPoints(length.size)
      : unit === "mm"
        ? mmToPoints(length.size)
        : length.size;
  const emu = Math.round(points * EMU_PER_POINT);
  if (!(Number.isFinite(emu) && emu >= 0)) {
    throw statusError(StatusCode.InvalidArgument, `${label} must be a finite non-negative length`);
  }
  return emu;
}

export function emuSize(emu: number) {
  const points = emu / EMU_PER_POINT;
  return { pt: Math.round(points * 100) / 100, px: Math.round(pointsToPixels(points)) };
}

/** Describes a drawing object's anchor with A1 cells and sizes in points and 96-DPI pixels. */
export function describePlacement(anchor: DrawingAnchor) {
  return {
    anchorKind: ANCHOR_NAMES[anchor.kind] ?? "oneCell",
    ...(anchor.kind === AnchorKind.Absolute
      ? {
          position: {
            x: emuSize(anchor.fromColOffsetEmu),
            y: emuSize(anchor.fromRowOffsetEmu),
          },
        }
      : {
          from: refA1(anchor.fromRow, anchor.fromCol),
          offset: { x: emuSize(anchor.fromColOffsetEmu), y: emuSize(anchor.fromRowOffsetEmu) },
        }),
    ...(anchor.kind === AnchorKind.TwoCell
      ? {
          to: refA1(anchor.toRow, anchor.toCol),
          toOffset: { x: emuSize(anchor.toColOffsetEmu), y: emuSize(anchor.toRowOffsetEmu) },
        }
      : {}),
    width: emuSize(anchor.widthEmu),
    height: emuSize(anchor.heightEmu),
  };
}

type StackableDrawing = Drawing & { setZOrder(zOrderIndex: number): void };

function isStackable(drawing: Drawing): drawing is StackableDrawing {
  return "setZOrder" in drawing;
}

/** Lists a sheet's drawing objects with A1 anchors and sizes in points and 96-DPI pixels. */
export function listSessionImages(id: string, sheet: number | string | undefined) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const sheetHandle = sheetAt(session.workbook, sheetIndex);
  const objects = sheetHandle.listDrawings().map((drawing) => {
    const properties = drawing.getDrawingProperties();
    return {
      objectId: properties.objectId,
      kind: KIND_NAMES[properties.kind] ?? "other",
      name: properties.name,
      description: properties.description ?? "",
      ...describePlacement(properties.anchor),
      ...(properties.kind === DrawingObjectKind.Picture
        ? { format: imageFormatName(sheetHandle.picture(drawing).getProperties().format) }
        : {}),
    };
  });
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    sheetName: sheetName(session, sheetIndex),
    count: objects.length,
    objects,
  };
}

/** Reads a picture's media bytes, optionally writing them to `outputPath`. */
export function getSessionImage(
  id: string,
  sheet: number | string | undefined,
  objectId: number,
  outputPath?: string,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const result = sheetAt(session.workbook, sheetIndex).picture(objectId).getProperties();
  const bytes = Buffer.from(result.bytes);
  const written = outputPath === undefined ? undefined : resolveUserPath(outputPath);
  if (written !== undefined) {
    writeFileSync(written, bytes);
  }
  return {
    bytes,
    mimeType: imageMimeType(result.format),
    metadata: {
      session: publicInfo(session),
      sheet: sheetIndex,
      objectId,
      format: imageFormatName(result.format),
      pxWidth: result.widthPx,
      pxHeight: result.heightPx,
      byteLength: bytes.byteLength,
      ...(written ? { outputPath: written } : {}),
    },
  };
}

/** A cell on the selected sheet named by `anchor`, which must not carry a sheet name. */
function anchorCell(anchor: string) {
  const cell = parseCellRef(anchor);
  if (cell.sheetName !== undefined) {
    throw statusError(
      StatusCode.InvalidArgument,
      "anchor must be a cell on the selected sheet, without a sheet name",
    );
  }
  assertInGrid(cell.row, cell.col);
  return cell;
}

/** Reads an image file the caller names, refusing one past the size cap. */
function readImageFile(input: string) {
  const path = resolveUserPath(input);
  const bytes = new Uint8Array(readFileSync(path));
  if (bytes.byteLength > MAX_IMAGE_BYTES) {
    throw statusError(
      StatusCode.InvalidArgument,
      `image exceeds ${MAX_IMAGE_BYTES} bytes: ${path}`,
    );
  }
  return { path, bytes };
}

/**
 * Inserts a PNG/JPEG/GIF/BMP file as a picture anchored at a cell. An omitted
 * width or height keeps the image's aspect ratio; omitting both uses its
 * pixel size at 96 DPI.
 */
export async function insertSessionImage(
  id: string,
  sheet: number | string | undefined,
  input: ImageInsertInput,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const anchor = anchorCell(input.anchor);
  const { path, bytes } = readImageFile(input.path);
  const info = (await formulonModule()).getImageInfo(bytes);

  let widthEmu = lengthToEmu(input.width, "width");
  let heightEmu = lengthToEmu(input.height, "height");
  if (widthEmu > 0 && heightEmu === 0) {
    heightEmu = Math.round((widthEmu * info.heightPx) / info.widthPx);
  } else if (heightEmu > 0 && widthEmu === 0) {
    widthEmu = Math.round((heightEmu * info.widthPx) / info.heightPx);
  }

  touch(session, true);
  const twoCell = input.anchorKind === "twoCell";
  const picture = sheetAt(session.workbook, sheetIndex).addPicture(bytes, {
    name: input.name ?? "",
    description: input.description ?? "",
    anchor: {
      kind: twoCell ? AnchorKind.TwoCell : AnchorKind.OneCell,
      ...(twoCell ? { editAs: AnchorEditAs.OneCell } : {}),
      fromRow: anchor.row,
      fromCol: anchor.col,
      fromColOffsetEmu: lengthToEmu(input.offsetX, "offsetX"),
      fromRowOffsetEmu: lengthToEmu(input.offsetY, "offsetY"),
      widthEmu,
      heightEmu,
    },
  });
  const { objectId } = picture;
  const placed = listSessionImages(id, sheetIndex).objects.find(
    (object) => object.objectId === objectId,
  );
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    objectId,
    source: {
      path,
      format: imageFormatName(info.format),
      pxWidth: info.widthPx,
      pxHeight: info.heightPx,
    },
    ...(placed ? { placed } : {}),
  };
}

/** Removes a picture or chart; other drawing objects, values and formatting are kept. */
export function removeSessionImage(
  id: string,
  sheet: number | string | undefined,
  objectId: number,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const sheetHandle = sheetAt(session.workbook, sheetIndex);
  const { kind } = sheetHandle.drawing(objectId);
  if (kind !== DrawingObjectKind.Picture && kind !== DrawingObjectKind.Chart) {
    throw statusError(
      StatusCode.InvalidArgument,
      `drawing object ${objectId} is not a picture or chart`,
    );
  }
  touch(session, true);
  sheetHandle.removeDrawing(objectId);
  touch(session, true);
  return { session: publicInfo(session), sheet: sheetIndex, objectId, removed: true };
}

export type ImageMoveInput = {
  anchor?: string;
  offsetX?: ImageLength;
  offsetY?: ImageLength;
  width?: ImageLength;
  height?: ImageLength;
  anchorKind?: "oneCell" | "twoCell";
};

/**
 * Moves or resizes a picture or chart in place, keeping its id (and a picture's crop and rotation).
 * Anything the input omits keeps its current value; a lone width or height
 * keeps the object's current aspect ratio.
 */
export function moveSessionImage(
  id: string,
  sheet: number | string | undefined,
  objectId: number,
  input: ImageMoveInput,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const drawing = sheetAt(session.workbook, sheetIndex).drawing(objectId);
  const current = drawing.getDrawingProperties().anchor;
  const anchor = input.anchor === undefined ? undefined : parseCellRef(input.anchor);
  if (anchor?.sheetName !== undefined) {
    throw statusError(
      StatusCode.InvalidArgument,
      "anchor must be a cell on the selected sheet, without a sheet name",
    );
  }
  if (anchor) {
    assertInGrid(anchor.row, anchor.col);
  }

  let widthEmu = lengthToEmu(input.width, "width");
  let heightEmu = lengthToEmu(input.height, "height");
  if (widthEmu > 0 && heightEmu === 0) {
    heightEmu = Math.round((widthEmu * current.heightEmu) / current.widthEmu);
  } else if (heightEmu > 0 && widthEmu === 0) {
    widthEmu = Math.round((heightEmu * current.widthEmu) / current.heightEmu);
  }
  const placement = {
    ...(anchor ? { fromRow: anchor.row, fromCol: anchor.col } : {}),
    ...(input.offsetX === undefined
      ? {}
      : { fromColOffsetEmu: lengthToEmu(input.offsetX, "offsetX") }),
    ...(input.offsetY === undefined
      ? {}
      : { fromRowOffsetEmu: lengthToEmu(input.offsetY, "offsetY") }),
    ...(widthEmu > 0 ? { widthEmu } : {}),
    ...(heightEmu > 0 ? { heightEmu } : {}),
  };
  const finalKind =
    input.anchorKind === undefined
      ? current.kind
      : input.anchorKind === "twoCell"
        ? AnchorKind.TwoCell
        : AnchorKind.OneCell;
  touch(session, true);
  if (finalKind === AnchorKind.TwoCell && Object.keys(placement).length > 0) {
    // A two-cell anchor takes no size and keeps its `to` corner when only `from` moves, so the
    // move is made on a one-cell anchor (which keeps the box size) and converted back; the
    // conversion derives the `to` corner from the size.
    drawing.update({ anchor: { kind: AnchorKind.OneCell, ...placement } });
    drawing.update({
      anchor: {
        kind: AnchorKind.TwoCell,
        ...(current.kind === AnchorKind.TwoCell ? { editAs: current.editAs } : {}),
      },
    });
  } else {
    drawing.update({ anchor: { kind: finalKind, ...placement } });
  }
  const placed = listSessionImages(id, sheetIndex).objects.find(
    (object) => object.objectId === objectId,
  );
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    objectId,
    ...(placed ? { placed } : {}),
  };
}

/** Moves a drawing object to position `index` in the back-to-front stacking order. */
export function reorderSessionImage(
  id: string,
  sheet: number | string | undefined,
  objectId: number,
  index: number,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const drawing = sheetAt(session.workbook, sheetIndex).drawing(objectId);
  if (!isStackable(drawing)) {
    throw statusError(StatusCode.InvalidArgument, `drawing object ${objectId} cannot be reordered`);
  }
  touch(session, true);
  drawing.setZOrder(index);
  const order = listSessionImages(id, sheetIndex).objects.map((object) => object.objectId);
  return { session: publicInfo(session), sheet: sheetIndex, objectId, index, order };
}

/** Places a PNG/JPEG file in a cell, replacing whatever the cell held, a formula or another picture. */
export async function insertSessionCellPicture(
  id: string,
  sheet: number | string | undefined,
  input: { path: string; anchor: string; description?: string },
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const cell = anchorCell(input.anchor);
  const { path, bytes } = readImageFile(input.path);
  const info = (await formulonModule()).getImageInfo(bytes);
  touch(session, true);
  sheetAt(session.workbook, sheetIndex).setCellPicture(cell.row, cell.col, {
    bytes,
    altText: input.description ?? null,
  });
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    anchor: refA1(cell.row, cell.col),
    source: {
      path,
      format: imageFormatName(info.format),
      pxWidth: info.widthPx,
      pxHeight: info.heightPx,
    },
  };
}

/** Reads the picture placed in a cell, optionally writing its bytes to `outputPath`. */
export function getSessionCellPicture(
  id: string,
  sheet: number | string | undefined,
  anchor: string,
  outputPath?: string,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const cell = anchorCell(anchor);
  const picture = sheetAt(session.workbook, sheetIndex).getCellPicture(cell.row, cell.col);
  if (picture === null) {
    throw statusError(StatusCode.NotFound, `no picture in cell ${refA1(cell.row, cell.col)}`);
  }
  const bytes = Buffer.from(picture.bytes);
  const written = outputPath === undefined ? undefined : resolveUserPath(outputPath);
  if (written !== undefined) {
    writeFileSync(written, bytes);
  }
  return {
    bytes,
    mimeType: imageMimeType(picture.format),
    metadata: {
      session: publicInfo(session),
      sheet: sheetIndex,
      anchor: refA1(cell.row, cell.col),
      format: imageFormatName(picture.format),
      pxWidth: picture.widthPx,
      pxHeight: picture.heightPx,
      altText: picture.altText,
      byteLength: bytes.byteLength,
      ...(written ? { outputPath: written } : {}),
    },
  };
}

/** Removes the picture placed in a cell, which empties it; `removed` is false when it held none. */
export function removeSessionCellPicture(
  id: string,
  sheet: number | string | undefined,
  anchor: string,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const cell = anchorCell(anchor);
  const handle = sheetAt(session.workbook, sheetIndex);
  const removed = handle.getCellPicture(cell.row, cell.col) !== null;
  if (removed) {
    touch(session, true);
    handle.setCellPicture(cell.row, cell.col, null);
  }
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    anchor: refA1(cell.row, cell.col),
    removed,
  };
}
