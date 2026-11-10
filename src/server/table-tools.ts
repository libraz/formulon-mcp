import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { TABLE_TOTALS_FUNCTION_NAMES, tableOperation } from "../session/tables.js";
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
          .array(
            z.object({
              name: z
                .string()
                .optional()
                .describe(
                  "Create only: the column's unique non-empty name. Name every column or none; none derives the names from the header row. Rename an existing column with Table.renameColumn through formulon_workbook_call.",
                ),
              totalsRowFunction: z
                .enum(TABLE_TOTALS_FUNCTION_NAMES)
                .nullable()
                .optional()
                .describe(
                  "The totals-row function (Excel writes SUBTOTAL over the column); custom needs totalsRowFormula; null is none.",
                ),
              totalsRowLabel: z
                .string()
                .nullable()
                .optional()
                .describe("Totals-row label text; null clears it."),
              totalsRowFormula: z
                .string()
                .nullable()
                .optional()
                .describe("The custom function's formula, leading = included; null clears it."),
              calculatedFormula: z
                .string()
                .nullable()
                .optional()
                .describe(
                  "The calculated-column formula as entered in the first data row (=[@Qty]*10). A formula that differs from the column's current one fills every data cell, constants included; the same formula changes nothing; null removes the mark and keeps the cells.",
                ),
            }),
          )
          .optional()
          .describe(
            "Per-column settings, element i for column i; it may be shorter than the table and {} changes nothing. An omitted key keeps the setting and null clears it. On update it applies after range, headerRow and totalsRow. Totals settings are stored while the totals row is off and written into its cells while it is shown.",
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
