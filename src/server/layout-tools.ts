import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { nonEmptyPrintSettingsSchema, styleSchema } from "../layout/schemas.js";
import type { LayoutOperation } from "../layout/types.js";
import { applySessionLayout } from "../session/layout.js";
import { fail, ok, sheetInputSchema } from "./shared.js";

const columnOperationSchema = z.object({
  type: z.literal("column"),
  first: z
    .number()
    .int()
    .nonnegative()
    .describe(
      "Zero-based first column index (maximum XFD); stored width is capped at 255 characters.",
    ),
  last: z
    .number()
    .int()
    .nonnegative()
    .optional()
    .describe("Zero-based last column index; defaults to first."),
  size: z.number().finite().nonnegative(),
  unit: z
    .enum(["chars", "pt", "px", "mm"])
    .optional()
    .describe("Column width unit; defaults to Excel character units."),
});

const rowOperationSchema = z.object({
  type: z.literal("row"),
  first: z.number().int().nonnegative().describe("Zero-based first row index."),
  last: z
    .number()
    .int()
    .nonnegative()
    .optional()
    .describe("Zero-based last row index; defaults to first."),
  size: z.number().finite().nonnegative(),
  unit: z.enum(["pt", "px", "mm"]).optional().describe("Row height unit; defaults to points."),
});

const mergeOperationSchema = z.object({
  type: z.literal("merge"),
  operation: z.enum(["add", "remove"]),
  range: z.string().describe("A1 range on the selected sheet, for example A1:D1."),
});

const styleOperationSchema = z.object({
  type: z.literal("style"),
  range: z
    .string()
    .describe(
      "A1 range, whole columns (C:C) or whole rows (1:1) on the selected sheet; whole lines are styled as columns or rows.",
    ),
  style: styleSchema,
  base: z
    .enum(["existing", "default"])
    .default("existing")
    .describe("Apply deltas to each cell's current style, or start from the workbook default."),
});

const printOperationSchema = z.object({
  type: z.literal("print"),
  settings: nonEmptyPrintSettingsSchema,
});

export const layoutOperationSchema = z.discriminatedUnion("type", [
  columnOperationSchema,
  rowOperationSchema,
  mergeOperationSchema,
  styleOperationSchema,
  printOperationSchema,
]);

/** Registers the bounded, preflighted presentation batch operation. */
export function registerLayoutTools(server: McpServer): void {
  server.registerTool(
    "formulon_apply_layout",
    {
      title: "Apply a layout batch",
      description:
        "Apply column widths, row heights, merges, styles, and print settings in one ordered batch. The complete operation list is validated before any workbook mutation; a failed preflight leaves the session unchanged. Column sizes default to Excel character units, row sizes to points, and px/mm convert through 96-DPI display geometry. The batch caps at 200 operations, 255 stored width characters, 409 points per row, 10,000 expanded row writes, and 100,000 styled cells.",
      inputSchema: {
        ...sheetInputSchema,
        operations: z.array(layoutOperationSchema).min(1).max(200),
      },
    },
    ({ sessionId, sheet, operations }) => {
      try {
        return ok(applySessionLayout(sessionId, sheet, operations as LayoutOperation[]));
      } catch (error) {
        return fail(error);
      }
    },
  );
}
