import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { tableOperation } from "../session/tables.js";
import { fail, ok, sheetRefSchema } from "./shared.js";

/** Registers native worksheet-table authoring rather than presentation blocks. */
export function registerTableTools(server: McpServer): void {
  server.registerTool(
    "formulon_table_operation",
    {
      title: "Worksheet table operation",
      description:
        "List, create, update, or remove a native Excel worksheet Table. Create derives unique non-empty column names from the first row when columns is omitted; this is structured Table metadata, not visual formatting. Renaming a table on update rewrites the structured references to it. Table indexes are global across the workbook.",
      inputSchema: {
        sessionId: z.string(),
        operation: z.enum(["list", "create", "update", "remove"]),
        sheet: sheetRefSchema
          .optional()
          .describe(
            "Fallback sheet for an unqualified range; a qualified range supplies its own sheet.",
          ),
        range: z.string().optional().describe("A1 table range, for example Data!A1:C20."),
        ref: z.string().optional().describe("Alias for range, useful for update."),
        index: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe("Global table index for update/remove."),
        name: z
          .string()
          .optional()
          .describe(
            "Native table name for create; on update, renames the table and the formulas that reference it.",
          ),
        columns: z
          .array(z.string())
          .optional()
          .describe(
            "One unique non-empty column name per range column; omitted means derive from the header row.",
          ),
        style: z
          .string()
          .optional()
          .describe("Excel TableStyle name; an empty string clears it on update."),
        styleName: z.string().optional().describe("Alias for style."),
        headerRow: z
          .boolean()
          .optional()
          .describe("Whether the first range row is the table header; defaults true on create."),
        totalsRow: z
          .boolean()
          .optional()
          .describe("Whether the table has a totals row; defaults false on create."),
        showFirstColumn: z.boolean().optional().describe("Emphasize the first column."),
        showLastColumn: z.boolean().optional().describe("Emphasize the last column."),
        showRowStripes: z
          .boolean()
          .optional()
          .describe("Band rows; a styled table starts with this on."),
        showColumnStripes: z.boolean().optional().describe("Band columns."),
        totals: z
          .array(
            z.object({
              column: z
                .union([z.string(), z.number().int().nonnegative()])
                .describe("Column name or zero-based column index within the table."),
              function: z
                .enum([
                  "none",
                  "sum",
                  "min",
                  "max",
                  "average",
                  "count",
                  "countNums",
                  "stdDev",
                  "var",
                  "custom",
                ])
                .optional()
                .describe("Totals function; Excel writes SUBTOTAL over the column."),
              label: z.string().optional().describe("Text shown instead of a function."),
              formula: z.string().optional().describe("Formula for the custom function."),
            }),
          )
          .optional()
          .describe(
            "Per-column totals-row settings, applied after range, headerRow and totalsRow; stored while the totals row is off and written into its cells while it is shown.",
          ),
      },
    },
    ({
      sessionId,
      operation,
      sheet,
      range,
      ref,
      index,
      name,
      columns,
      style,
      styleName,
      headerRow,
      totalsRow,
      ...rest
    }) => {
      try {
        return ok(
          tableOperation(sessionId, operation, {
            sheet,
            range,
            ref,
            index,
            name,
            columns,
            style,
            styleName,
            headerRow,
            totalsRow,
            ...rest,
          }),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );
}
