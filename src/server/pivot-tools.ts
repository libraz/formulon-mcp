import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { createSessionPivot } from "../session/pivots.js";
import { fail, ok, sheetRefSchema } from "./shared.js";

const pivotValueSchema = z.object({
  field: z.string().min(1).describe("Source header to aggregate."),
  aggregation: z
    .enum([
      "sum",
      "count",
      "average",
      "max",
      "min",
      "product",
      "countNumbers",
      "stddev",
      "stddevp",
      "var",
      "varp",
    ])
    .default("sum")
    .describe("Aggregation function; defaults to sum."),
  name: z
    .string()
    .optional()
    .describe("Displayed data-field name; defaults to '<Aggregation> of <field>'."),
  numberFormat: z
    .string()
    .optional()
    .describe(
      "Excel format code, registered in the workbook before it is attached to the data field.",
    ),
});

/** Registers one bounded, source-backed PivotTable authoring operation. */
export function registerPivotTools(server: McpServer): void {
  server.registerTool(
    "formulon_create_pivot",
    {
      title: "Create PivotTable",
      description:
        "Create a native Excel PivotTable from a worksheet range. The first source row is a unique, non-empty text header row and at least one data row is required; source cells are capped at 10,000 and are read in one operation after recalculation. Every source header must be assigned exactly once to rows, columns, pages, or a distinct value field (a value field may have multiple aggregations); unknown, unused, or cross-axis fields are rejected before mutation. Values may use sum, count, average, max, min, product, countNumbers, stddev, stddevp, var, or varp. target is an A1 anchor and may include a sheet name; targetSheet defaults to the source sheet.",
      inputSchema: {
        sessionId: z.string(),
        sourceRange: z
          .string()
          .describe("Qualified or unqualified source A1 range, including its header row."),
        sourceSheet: sheetRefSchema
          .optional()
          .describe(
            "Fallback source sheet for an unqualified sourceRange; defaults to the first sheet.",
          ),
        target: z
          .string()
          .describe("Pivot top-left A1 anchor; may be qualified with a target sheet."),
        targetSheet: sheetRefSchema
          .optional()
          .describe(
            "Fallback target sheet when target is unqualified; defaults to the source sheet.",
          ),
        name: z.string().describe("Native PivotTable name."),
        rows: z.array(z.string()).default([]).describe("Source headers placed on the row axis."),
        columns: z
          .array(z.string())
          .default([])
          .describe("Source headers placed on the column axis."),
        pages: z
          .array(z.string())
          .default([])
          .describe("Source headers placed on the filter/page axis."),
        values: z
          .array(pivotValueSchema)
          .min(1)
          .describe("At least one value field; repeated entries for one source field are allowed."),
        layout: z.enum(["compact", "tabular", "outline"]).default("compact"),
        grandTotals: z
          .object({
            rows: z.boolean().default(true),
            columns: z.boolean().default(true),
          })
          .optional()
          .describe("Whether row and column grand-total bands are enabled."),
        sourceLimit: z.number().int().min(2).max(10_000).default(10_000),
      },
    },
    ({
      sessionId,
      sourceRange,
      sourceSheet,
      target,
      targetSheet,
      name,
      rows,
      columns,
      pages,
      values,
      layout,
      grandTotals,
      sourceLimit,
    }) => {
      try {
        return ok(
          createSessionPivot(sessionId, {
            sourceRange,
            sourceSheet,
            target,
            targetSheet,
            name,
            rows,
            columns,
            pages,
            values,
            layout,
            grandTotals,
            sourceLimit,
          }),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );
}
