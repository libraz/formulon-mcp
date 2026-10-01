import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  applySessionMutations,
  applySheetOperation,
  editSessionStructure,
  type RangeWriteCell,
  setSessionDefinedName,
  setSessionRange,
  setSessionSheetView,
} from "../sessions.js";
import { fail, ok, sheetRefSchema } from "./shared.js";

const cellMutationSchema = z.discriminatedUnion("type", [
  z.object({
    type: z.literal("number"),
    sheet: sheetRefSchema.optional(),
    a1: z.string().optional(),
    row: z.number().int().nonnegative().optional(),
    col: z.number().int().nonnegative().optional(),
    value: z.number(),
  }),
  z.object({
    type: z.literal("bool"),
    sheet: sheetRefSchema.optional(),
    a1: z.string().optional(),
    row: z.number().int().nonnegative().optional(),
    col: z.number().int().nonnegative().optional(),
    value: z.boolean(),
  }),
  z.object({
    type: z.literal("text"),
    sheet: sheetRefSchema.optional(),
    a1: z.string().optional(),
    row: z.number().int().nonnegative().optional(),
    col: z.number().int().nonnegative().optional(),
    value: z.string(),
  }),
  z.object({
    type: z.literal("blank"),
    sheet: sheetRefSchema.optional(),
    a1: z.string().optional(),
    row: z.number().int().nonnegative().optional(),
    col: z.number().int().nonnegative().optional(),
  }),
  z.object({
    type: z.literal("formula"),
    sheet: sheetRefSchema.optional(),
    a1: z.string().optional(),
    row: z.number().int().nonnegative().optional(),
    col: z.number().int().nonnegative().optional(),
    formula: z.string(),
  }),
]);

const rangeWriteCellSchema = z.union([
  z.number(),
  z.boolean(),
  z.string(),
  z.null(),
  z.object({ f: z.string() }).strict(),
  z.object({ blank: z.literal(true) }).strict(),
]);

export function registerMutationTools(server: McpServer): void {
  server.registerTool(
    "formulon_set_cells",
    {
      title: "Set session cells",
      description:
        "Apply cell mutations to an open workbook session. Cells may use A1 refs or zero-based row/col.",
      inputSchema: {
        sessionId: z.string(),
        recalc: z.boolean().default(true),
        mutations: z.array(cellMutationSchema).min(1).max(10_000),
      },
    },
    ({ sessionId, mutations, recalc }) => {
      try {
        return ok(applySessionMutations(sessionId, mutations, recalc));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_set_range",
    {
      title: "Set range block",
      description:
        'Write a 2D block of values starting at an anchor cell. Each element\'s JSON type picks the cell type: number, boolean, or string. Use {"f":"=SUM(A1:A3)"} for a formula, {"blank":true} to clear a cell, and null to leave it untouched. Far more compact than set_cells for tables.',
      inputSchema: {
        sessionId: z.string(),
        start: z.string().describe("Anchor (top-left) A1 cell, for example Sheet1!B2 or B2."),
        sheet: sheetRefSchema
          .optional()
          .describe("Sheet used when start has no Sheet!-prefix; defaults to the first sheet."),
        values: z
          .array(z.array(rangeWriteCellSchema))
          .min(1)
          .describe("Row-major 2D array of cell values."),
        recalc: z.boolean().default(true),
      },
    },
    ({ sessionId, start, sheet, values, recalc }) => {
      try {
        return ok(setSessionRange(sessionId, start, values as RangeWriteCell[][], sheet, recalc));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_sheet_operation",
    {
      title: "Sheet operation",
      description: "Add, remove, rename, or move a sheet in an open workbook session.",
      inputSchema: {
        sessionId: z.string(),
        operation: z.enum(["add", "remove", "rename", "move"]),
        name: z.string().optional().describe("Sheet name for add."),
        index: z.number().int().nonnegative().optional().describe("Sheet index for remove/rename."),
        newName: z.string().optional().describe("New sheet name for rename."),
        fromIndex: z.number().int().nonnegative().optional().describe("Source index for move."),
        toIndex: z.number().int().nonnegative().optional().describe("Destination index for move."),
      },
    },
    ({ sessionId, operation, name, index, newName, fromIndex, toIndex }) => {
      try {
        return ok(
          applySheetOperation(sessionId, operation, { name, index, newName, fromIndex, toIndex }),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_set_defined_name",
    {
      title: "Set defined name",
      description:
        "Add, replace, or remove a defined name. Without `sheet` the name is workbook-scoped; with it the name is local to that sheet. Print settings (_xlnm.Print_Area, _xlnm.Print_Titles) must be sheet-scoped — Excel ignores a workbook-scoped one without reporting an error. Empty formula removes it.",
      inputSchema: {
        sessionId: z.string(),
        name: z.string(),
        formula: z.string().describe("Formula with or without '='; pass empty string to remove."),
        sheet: sheetRefSchema
          .optional()
          .describe("Scope the name to this sheet. Omit for workbook scope."),
      },
    },
    ({ sessionId, name, formula, sheet }) => {
      try {
        return ok(setSessionDefinedName(sessionId, name, formula, sheet));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_edit_structure",
    {
      title: "Edit rows or columns",
      description: "Insert or delete rows/columns in an open workbook session.",
      inputSchema: {
        sessionId: z.string(),
        operation: z.enum(["insertRows", "deleteRows", "insertCols", "deleteCols"]),
        sheet: sheetRefSchema.optional().default(0),
        start: z.number().int().nonnegative(),
        count: z.number().int().positive(),
      },
    },
    ({ sessionId, operation, sheet, start, count }) => {
      try {
        return ok(editSessionStructure(sessionId, operation, sheet, start, count));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_set_sheet_view",
    {
      title: "Set sheet view",
      description: "Set sheet zoom, frozen panes, or tab visibility.",
      inputSchema: {
        sessionId: z.string(),
        sheet: sheetRefSchema.optional().default(0),
        zoom: z.number().int().min(10).max(400).optional(),
        freezeRows: z.number().int().nonnegative().optional(),
        freezeCols: z.number().int().nonnegative().optional(),
        hidden: z
          .boolean()
          .optional()
          .describe("Two-state tab visibility; cannot state veryHidden."),
        visibility: z
          .enum(["visible", "hidden", "veryHidden"])
          .optional()
          .describe(
            "Three-state tab visibility. Excel leaves a veryHidden sheet out of its Unhide dialog.",
          ),
      },
    },
    ({ sessionId, sheet, zoom, freezeRows, freezeCols, hidden, visibility }) => {
      try {
        return ok(
          setSessionSheetView(sessionId, sheet, {
            zoom,
            freezeRows,
            freezeCols,
            hidden,
            visibility,
          }),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );
}
