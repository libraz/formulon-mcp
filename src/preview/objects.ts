import { DrawingObjectKind } from "@libraz/formulon";
import { AnchorKind, GeometryMode, ImageFormat, type Workbook } from "../formulon.js";
import { type Rect, type WarningState, warning } from "./read.js";
import type { PreviewObject, PreviewObjectMetadata } from "./types.js";

const EMU_PER_POINT = 12_700;
const MAX_PREVIEW_OBJECTS = 100;
const MAX_EMBEDDED_IMAGE_BYTES = 2 * 1024 * 1024;

const KIND_NAMES: Record<number, string> = {
  [DrawingObjectKind.Picture]: "picture",
  [DrawingObjectKind.Shape]: "shape",
  [DrawingObjectKind.Chart]: "chart",
  [DrawingObjectKind.Group]: "group",
  [DrawingObjectKind.Connector]: "connector",
  [DrawingObjectKind.GraphicFrame]: "graphic frame",
  [DrawingObjectKind.Other]: "object",
};

const IMAGE_MIME: Partial<Record<number, string>> = {
  [ImageFormat.Png]: "image/png",
  [ImageFormat.Jpeg]: "image/jpeg",
  [ImageFormat.Gif]: "image/gif",
};

/** Lists the sheet's drawing objects that intersect `rect`, positioned relative to its top-left corner. */
export function readObjects(
  workbook: Workbook,
  sheet: number,
  rect: Rect,
  warnings: WarningState,
): { objects: PreviewObject[]; metadata: PreviewObjectMetadata[] } {
  const sheetHandle = workbook.sheet(sheet);
  const listed = sheetHandle.listDrawings();
  const objects: PreviewObject[] = [];
  const metadata: PreviewObjectMetadata[] = [];
  if (listed.length === 0) {
    return { objects, metadata };
  }
  const origin = sheetHandle.getCellRectPt(rect, GeometryMode.Display);
  const pointAt = (row: number, col: number, rowOff: number, colOff: number) => {
    const cell = sheetHandle.getCellRectPt(
      { firstRow: row, firstCol: col, lastRow: row, lastCol: col },
      GeometryMode.Display,
    );
    return { x: cell.x + colOff / EMU_PER_POINT, y: cell.y + rowOff / EMU_PER_POINT };
  };
  let embeddedBytes = 0;
  let placeholders = 0;
  for (const drawing of listed) {
    const object = drawing.getDrawingProperties();
    const { anchor } = object;
    let start: { x: number; y: number };
    let width = anchor.widthEmu / EMU_PER_POINT;
    let height = anchor.heightEmu / EMU_PER_POINT;
    if (anchor.kind === AnchorKind.Absolute) {
      start = {
        x: anchor.fromColOffsetEmu / EMU_PER_POINT,
        y: anchor.fromRowOffsetEmu / EMU_PER_POINT,
      };
    } else {
      start = pointAt(
        anchor.fromRow,
        anchor.fromCol,
        anchor.fromRowOffsetEmu,
        anchor.fromColOffsetEmu,
      );
      if (anchor.kind === AnchorKind.TwoCell) {
        const end = pointAt(
          anchor.toRow,
          anchor.toCol,
          anchor.toRowOffsetEmu,
          anchor.toColOffsetEmu,
        );
        width = end.x - start.x;
        height = end.y - start.y;
      }
    }
    const intersectsRect =
      width > 0 &&
      height > 0 &&
      start.x < origin.x + origin.widthPt &&
      start.x + width > origin.x &&
      start.y < origin.y + origin.heightPt &&
      start.y + height > origin.y;
    if (!intersectsRect) {
      continue;
    }
    if (objects.length >= MAX_PREVIEW_OBJECTS) {
      warning(
        warnings,
        `Only the first ${MAX_PREVIEW_OBJECTS} drawing objects are shown in preview.`,
      );
      break;
    }
    const placed: PreviewObject = {
      objectId: object.objectId,
      kind: KIND_NAMES[object.kind] ?? "object",
      name: object.name,
      x: start.x - origin.x,
      y: start.y - origin.y,
      width,
      height,
    };
    if (object.kind === DrawingObjectKind.Picture) {
      const picture = sheetHandle.picture(drawing).getProperties();
      const mime = IMAGE_MIME[picture.format];
      if (mime) {
        const base64 = Buffer.from(picture.bytes).toString("base64");
        if (embeddedBytes + base64.length <= MAX_EMBEDDED_IMAGE_BYTES) {
          embeddedBytes += base64.length;
          placed.image = { mime, base64 };
        }
      }
    }
    if (!placed.image) {
      placeholders += 1;
    }
    objects.push(placed);
    metadata.push({
      objectId: placed.objectId,
      kind: placed.kind,
      name: placed.name,
      rendered: placed.image !== undefined,
    });
  }
  if (placeholders > 0) {
    warning(
      warnings,
      "Charts, shapes, and unsupported or oversized pictures are drawn as dashed placeholders in preview.",
    );
  }
  return { objects, metadata };
}
