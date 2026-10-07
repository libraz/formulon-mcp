import {
  AnchorKind,
  assertStatus,
  DrawingObjectKind,
  GeometryMode,
  ImageFormat,
  requireListResult,
  type Workbook,
} from "../formulon.js";
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
  const listed = requireListResult(workbook.listDrawingObjects(sheet), "list drawing objects");
  const objects: PreviewObject[] = [];
  const metadata: PreviewObjectMetadata[] = [];
  if (listed.length === 0) {
    return { objects, metadata };
  }
  const origin = workbook.getCellRectPt(sheet, rect, GeometryMode.Display);
  assertStatus(origin.status, "read preview range geometry");
  const pointAt = (row: number, col: number, rowOff: number, colOff: number) => {
    const cell = workbook.getCellRectPt(
      sheet,
      { firstRow: row, firstCol: col, lastRow: row, lastCol: col },
      GeometryMode.Display,
    );
    assertStatus(cell.status, "read drawing anchor geometry");
    return { x: cell.x + colOff / EMU_PER_POINT, y: cell.y + rowOff / EMU_PER_POINT };
  };
  let embeddedBytes = 0;
  let placeholders = 0;
  for (const object of listed) {
    let start: { x: number; y: number };
    let width = object.cx / EMU_PER_POINT;
    let height = object.cy / EMU_PER_POINT;
    if (object.anchorKind === AnchorKind.Absolute) {
      start = { x: object.fromColOff / EMU_PER_POINT, y: object.fromRowOff / EMU_PER_POINT };
    } else {
      start = pointAt(object.fromRow, object.fromCol, object.fromRowOff, object.fromColOff);
      if (object.anchorKind === AnchorKind.TwoCell) {
        const end = pointAt(object.toRow, object.toCol, object.toRowOff, object.toColOff);
        width = end.x - start.x;
        height = end.y - start.y;
      }
    }
    const intersectsRect =
      width > 0 &&
      height > 0 &&
      start.x < origin.x + origin.width &&
      start.x + width > origin.x &&
      start.y < origin.y + origin.height &&
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
    const mime =
      object.kind === DrawingObjectKind.Picture ? IMAGE_MIME[object.imageFormat] : undefined;
    if (mime) {
      const image = workbook.getImage(sheet, object.objectId);
      assertStatus(image.status, `read picture ${object.objectId}`);
      const base64 = Buffer.from(image.bytes).toString("base64");
      if (embeddedBytes + base64.length <= MAX_EMBEDDED_IMAGE_BYTES) {
        embeddedBytes += base64.length;
        placed.image = { mime, base64 };
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
