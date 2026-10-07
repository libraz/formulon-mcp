import { readFileSync, writeFileSync } from "node:fs";
import { parseCellRef } from "../a1.js";
import {
  AnchorEditAs,
  AnchorKind,
  assertStatus,
  DrawingObjectKind,
  findSheetIndex,
  ImageFormat,
  resolveUserPath,
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

const KIND_NAMES: Record<number, string> = {
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

function lengthToEmu(length: ImageLength | undefined, label: string): number {
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
    throw new Error(`${label} must be a finite non-negative length`);
  }
  return emu;
}

function emuSize(emu: number) {
  const points = emu / EMU_PER_POINT;
  return { pt: Math.round(points * 100) / 100, px: Math.round(pointsToPixels(points)) };
}

/** Lists a sheet's drawing objects with A1 anchors and sizes in points and 96-DPI pixels. */
export function listSessionImages(id: string, sheet: number | string | undefined) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const result = session.workbook.listDrawingObjects(sheetIndex);
  assertStatus(result.status, "list drawing objects");
  const objects = Array.from(result, (object) => ({
    objectId: object.objectId,
    kind: KIND_NAMES[object.kind] ?? "other",
    name: object.name,
    description: object.descr,
    anchorKind: ANCHOR_NAMES[object.anchorKind] ?? "oneCell",
    ...(object.anchorKind === AnchorKind.Absolute
      ? { position: { x: emuSize(object.fromColOff), y: emuSize(object.fromRowOff) } }
      : {
          from: refA1(object.fromRow, object.fromCol),
          offset: { x: emuSize(object.fromColOff), y: emuSize(object.fromRowOff) },
        }),
    ...(object.anchorKind === AnchorKind.TwoCell
      ? {
          to: refA1(object.toRow, object.toCol),
          toOffset: { x: emuSize(object.toColOff), y: emuSize(object.toRowOff) },
        }
      : {}),
    width: emuSize(object.cx),
    height: emuSize(object.cy),
    ...(object.kind === DrawingObjectKind.Picture
      ? { format: imageFormatName(object.imageFormat) }
      : {}),
  }));
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
  const result = session.workbook.getImage(sheetIndex, objectId);
  assertStatus(result.status, "get image");
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
      pxWidth: result.pxWidth,
      pxHeight: result.pxHeight,
      byteLength: bytes.byteLength,
      ...(written ? { outputPath: written } : {}),
    },
  };
}

/**
 * Inserts a PNG/JPEG/GIF/BMP file as a picture anchored at a cell. An omitted
 * width or height keeps the image's aspect ratio; omitting both uses its
 * pixel size at 96 DPI.
 */
export function insertSessionImage(
  id: string,
  sheet: number | string | undefined,
  input: ImageInsertInput,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const anchor = parseCellRef(input.anchor);
  if (anchor.sheetName !== undefined) {
    throw new Error("anchor must be a cell on the selected sheet, without a sheet name");
  }
  assertInGrid(anchor.row, anchor.col);
  const path = resolveUserPath(input.path);
  const bytes = new Uint8Array(readFileSync(path));
  if (bytes.byteLength > MAX_IMAGE_BYTES) {
    throw new Error(`image exceeds ${MAX_IMAGE_BYTES} bytes: ${path}`);
  }
  const info = session.workbook.probeImage(bytes);
  assertStatus(info.status, "probe image");

  let widthEmu = lengthToEmu(input.width, "width");
  let heightEmu = lengthToEmu(input.height, "height");
  if (widthEmu > 0 && heightEmu === 0) {
    heightEmu = Math.round((widthEmu * info.pxHeight) / info.pxWidth);
  } else if (heightEmu > 0 && widthEmu === 0) {
    widthEmu = Math.round((heightEmu * info.pxWidth) / info.pxHeight);
  }

  touch(session, true);
  const result = session.workbook.insertImage(sheetIndex, bytes, {
    name: input.name ?? "",
    descr: input.description ?? "",
    anchorKind: input.anchorKind === "twoCell" ? AnchorKind.TwoCell : AnchorKind.OneCell,
    editAs: AnchorEditAs.OneCell,
    row: anchor.row,
    col: anchor.col,
    colOffEmu: lengthToEmu(input.offsetX, "offsetX"),
    rowOffEmu: lengthToEmu(input.offsetY, "offsetY"),
    widthEmu,
    heightEmu,
  });
  assertStatus(result.status, "insert image");
  const placed = listSessionImages(id, sheetIndex).objects.find(
    (object) => object.objectId === result.objectId,
  );
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    objectId: result.objectId,
    source: {
      path,
      format: imageFormatName(info.format),
      pxWidth: info.pxWidth,
      pxHeight: info.pxHeight,
    },
    ...(placed ? { placed } : {}),
  };
}

/** Removes a picture; other drawing objects, values and formatting are kept. */
export function removeSessionImage(
  id: string,
  sheet: number | string | undefined,
  objectId: number,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  touch(session, true);
  assertStatus(session.workbook.removeImage(sheetIndex, objectId), "remove image");
  touch(session, true);
  return { session: publicInfo(session), sheet: sheetIndex, objectId, removed: true };
}
