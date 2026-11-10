import { writeFile } from "node:fs/promises";
import path from "node:path";
import {
  callWorkbookMethod,
  createSessionPivot,
  insertSessionImage,
  saveSession,
  setSessionRange,
} from "../../dist/sessions.js";
import { solidPng } from "./png.js";

export const SHEET = { kind: "Sheet", sheet: 0 };
export const WORKBOOK_OOXML = { kind: "WorkbookOoxml" };
export const SHEET_OOXML = { kind: "SheetOoxml", sheet: 0 };

/**
 * Fills the open session `id` with an object of every handle kind `formulon_workbook_call` can
 * address (a table, a pivot with its cache, a chart, a shape, a slicer and a picture), saves it to
 * `file` so the session is clean, and returns one target per kind naming those objects.
 */
export async function buildCallFixture(id, dir, file) {
  const picture = path.join(dir, "pixel.png");
  await writeFile(picture, solidPng(20, 10));
  const call = (target, method, args) => callWorkbookMethod(id, target, method, args).result;
  setSessionRange(
    id,
    "Sheet1!A1",
    [
      ["region", "item", "qty"],
      ["east", "pen", 2],
      ["west", "ink", 5],
      ["east", "ink", 7],
    ],
    0,
    false,
  );
  call(SHEET, "addTable", [{ range: "A1:C4", name: "Items" }]);
  const { cacheId } = createSessionPivot(id, {
    sourceRange: "Sheet1!A1:C4",
    targetAnchor: "F1",
    name: "Sales",
    rows: ["region"],
    columns: ["item"],
    values: [{ field: "qty", function: "sum" }],
  });
  const anchor = (fromRow) => ({ fromRow, fromCol: 0, widthEmu: 2_000_000, heightEmu: 1_000_000 });
  const chart = call(SHEET, "addChart", [{ type: 0, range: "A1:C4", anchor: anchor(10) }]);
  const shape = call(SHEET, "addShape", [
    { presetGeometry: "rect", text: "hi", anchor: anchor(20) },
  ]);
  const slicer = call(SHEET, "addSlicer", ["Items", "region", { anchor: anchor(30) }]);
  const image = await insertSessionImage(id, 0, { path: picture, anchor: "H2" });
  await saveSession(id, file);
  return {
    Workbook: undefined,
    Sheet: SHEET,
    Table: { kind: "Table", table: "Items" },
    PivotCache: { kind: "PivotCache", pivotCache: cacheId },
    WorkbookOoxml: WORKBOOK_OOXML,
    Drawing: { kind: "Drawing", sheet: 0, drawing: image.objectId },
    Picture: { kind: "Picture", sheet: 0, picture: image.objectId },
    Chart: { kind: "Chart", sheet: 0, chart: chart.properties.objectId },
    Shape: { kind: "Shape", sheet: 0, shape: shape.properties.objectId },
    Slicer: { kind: "Slicer", sheet: 0, slicer: slicer.properties.objectId },
    Pivot: { kind: "Pivot", sheet: 0, pivot: "Sales" },
    SheetOoxml: SHEET_OOXML,
    Formulon: { kind: "Formulon" },
  };
}
