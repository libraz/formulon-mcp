import { assertStatus, type Workbook } from "../formulon.js";
import { PRINT_PRESETS, type PrintSettingsInput, writePrintSettings } from "../print.js";
import {
  documentWidth,
  type Layout,
  NameTable,
  placeFields,
  placeSummary,
  placeTable,
  placeText,
  placeTitle,
  rectRef,
  ref,
} from "./layout.js";
import type { BuiltDocument, DocumentSpec, PlacedBlock } from "./types.js";

/**
 * Lays out and writes a document, returning where every block and named cell
 * landed so the caller can refine any of it with the primitive tools.
 */
export function buildDocument(
  wb: Workbook,
  sheet: number,
  anchor: { row: number; col: number },
  spec: DocumentSpec,
): BuiltDocument {
  if (spec.blocks.length === 0) {
    throw new Error("a document needs at least one block");
  }
  const width = documentWidth(spec);
  const layout: Layout = {
    wb,
    sheet,
    firstCol: anchor.col,
    lastCol: anchor.col + width - 1,
    theme: { border: "thin", outline: "medium", ...spec.theme },
    names: new NameTable(),
  };

  const placed: PlacedBlock[] = [];
  let top = anchor.row;
  let groupHeight = 0;
  // Columns already taken by blocks sharing the current row group. Two blocks
  // on one row that overlap would merge over each other, which the writer
  // cannot express and Excel would offer to repair.
  let groupSpans: { index: number; firstCol: number; lastCol: number }[] = [];
  spec.blocks.forEach((block, index) => {
    if (!block.sameRow) {
      top += groupHeight;
      groupHeight = 0;
      groupSpans = [];
    }
    if (block.type === "spacer") {
      const height = Math.max(0, block.rows ?? 1);
      placed.push({
        index,
        type: "spacer",
        range:
          height > 0
            ? rectRef({
                firstRow: top,
                firstCol: layout.firstCol,
                lastRow: top + height - 1,
                lastCol: layout.lastCol,
              })
            : "",
      });
      groupHeight = Math.max(groupHeight, height);
      return;
    }
    const result =
      block.type === "title"
        ? placeTitle(layout, block, top)
        : block.type === "text"
          ? placeText(layout, block, top)
          : block.type === "fields"
            ? placeFields(layout, block, top)
            : block.type === "table"
              ? placeTable(layout, block, top)
              : placeSummary(layout, block, top);
    const clash = groupSpans.find(
      (span) => result.firstCol <= span.lastCol && result.lastCol >= span.firstCol,
    );
    if (clash) {
      throw new Error(
        `block ${index} (${block.type}) shares a row with block ${clash.index} and overlaps its columns; give one of them a narrower span or drop sameRow`,
      );
    }
    groupSpans.push({ index, firstCol: result.firstCol, lastCol: result.lastCol });
    placed.push({ index, type: block.type, name: block.name, range: result.range });
    groupHeight = Math.max(groupHeight, result.height);
  });

  const lastRow = Math.max(anchor.row, top + groupHeight - 1);
  const range = rectRef({
    firstRow: anchor.row,
    firstCol: layout.firstCol,
    lastRow,
    lastCol: layout.lastCol,
  });

  let pageCount: number | undefined;
  if (spec.print !== undefined) {
    const preset = PRINT_PRESETS[spec.print];
    if (!preset) {
      throw new Error(
        `unknown print preset: ${spec.print} (expected one of ${Object.keys(PRINT_PRESETS).join(", ")})`,
      );
    }
    const settings: PrintSettingsInput = { ...preset, printArea: range };
    if (layout.headerRow !== undefined && spec.repeatTableHeader !== false) {
      const line = layout.headerRow + 1;
      settings.printTitles = { repeatRows: `${line}:${line}` };
    }
    writePrintSettings(wb, sheet, settings);
    const pagination = wb.paginate(sheet);
    assertStatus(pagination.status, "paginate document");
    pageCount = pagination.pageCount;
  }

  assertStatus(wb.recalc(), "recalc document");
  return {
    start: ref(anchor.row, anchor.col),
    range,
    width,
    blocks: placed,
    names: layout.names.toJson(),
    pageCount,
  };
}
