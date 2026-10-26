import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  addSessionDataTable,
  addSessionSparklines,
  getSessionDataTable,
  insertSessionShape,
  listSessionShapes,
  listSessionSparklines,
  recalcSessionDataTables,
  removeSessionDataTable,
  removeSessionShape,
  removeSessionSparklines,
  updateSessionShape,
  updateSessionSparklines,
} from "../sessions.js";
import { fail, imageLengthSchema, ok, sheetInputSchema } from "./shared.js";

const paintSchema = z.string().describe("Opaque #RRGGBB colour, or none.");
const colorSchema = z.string().describe("#RRGGBB or #AARRGGBB colour.");
const axisBoundSchema = z
  .union([z.enum(["individual", "group"]), z.number().finite()])
  .describe("individual scales each sparkline alone, group shares one scale, a number fixes it.");

/** Registers the drawing shape, sparkline and What-If data table tools. */
export function registerSheetObjectTools(server: McpServer): void {
  server.registerTool(
    "formulon_shape_operation",
    {
      title: "Shape operation",
      description:
        "List, insert, update, or remove drawing shapes, text boxes and connectors (lines, arrows). Insert places a one-cell shape at anchor with width and height, or a two-cell shape spanning anchor to `to`. An update changes only what it states, keeps the anchor kind, and moves or resizes the shape when a placement field is given. Shapes have no stacking-order control; pictures and charts are handled by formulon_image_operation.",
      inputSchema: {
        ...sheetInputSchema,
        operation: z.enum(["list", "insert", "update", "remove"]),
        objectId: z
          .number()
          .int()
          .positive()
          .optional()
          .describe("Shape id from list; required for update and remove."),
        anchor: z
          .string()
          .optional()
          .describe("Top-left cell on the selected sheet, for example B2; required for insert."),
        to: z
          .string()
          .optional()
          .describe("Bottom-right cell of a two-cell shape; the box stretches with rows/columns."),
        offsetX: imageLengthSchema.optional().describe("Offset from the anchor cell's left edge."),
        offsetY: imageLengthSchema.optional().describe("Offset from the anchor cell's top edge."),
        width: imageLengthSchema.optional().describe("Width of a one-cell shape."),
        height: imageLengthSchema.optional().describe("Height of a one-cell shape."),
        geometry: z
          .string()
          .optional()
          .describe(
            "DrawingML preset geometry, for example rect, roundRect, ellipse, rightArrow, line or straightConnector1; defaults to rect. A line or connector geometry holds no text.",
          ),
        text: z.string().optional().describe("Shape text; newlines separate paragraphs."),
        textBox: z.boolean().optional().describe("Mark the shape as a text box."),
        name: z.string().optional(),
        fill: paintSchema.optional(),
        line: paintSchema.optional().describe("Outline colour as opaque #RRGGBB, or none."),
        lineWidth: imageLengthSchema.optional().describe("Outline width."),
        rotation: z.number().finite().optional().describe("Clockwise rotation in degrees."),
        hidden: z.boolean().optional(),
      },
    },
    ({ sessionId, sheet, operation, objectId, anchor, ...fields }) => {
      try {
        if (operation === "list") {
          return ok(listSessionShapes(sessionId, sheet));
        }
        if (operation === "insert") {
          if (anchor === undefined) {
            throw new Error("anchor is required for insert");
          }
          return ok(insertSessionShape(sessionId, sheet, { anchor, ...fields }));
        }
        if (objectId === undefined) {
          throw new Error(`objectId is required for ${operation}`);
        }
        if (operation === "remove") {
          return ok(removeSessionShape(sessionId, sheet, objectId));
        }
        return ok(updateSessionShape(sessionId, sheet, objectId, { anchor, ...fields }));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_sparkline_operation",
    {
      title: "Sparkline operation",
      description:
        "List, add, update, or remove sparkline groups: in-cell line, column or win/loss charts. A group shares its type, colours and axes across its sparklines; each sparkline plots one data range into one location cell. An update changes only the stated settings, and `sparklines` replaces the group's sparklines. Group indexes shift when one is removed.",
      inputSchema: {
        ...sheetInputSchema,
        operation: z.enum(["list", "add", "update", "remove"]),
        index: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe("Group index from list; required for update and remove."),
        sparklines: z
          .array(
            z.object({
              data: z
                .string()
                .describe("Data range, for example B2:M2; a bare range is on the selected sheet."),
              location: z.string().describe("Cell on the selected sheet that shows the sparkline."),
            }),
          )
          .optional()
          .describe("Required for add; on update, replaces the group's sparklines."),
        type: z.enum(["line", "column", "winLoss"]).optional().describe("Defaults to line on add."),
        markers: z.boolean().optional().describe("Mark every point of a line sparkline."),
        high: z.boolean().optional().describe("Mark the highest point."),
        low: z.boolean().optional().describe("Mark the lowest point."),
        first: z.boolean().optional().describe("Mark the first point."),
        last: z.boolean().optional().describe("Mark the last point."),
        negative: z.boolean().optional().describe("Mark negative points."),
        displayXAxis: z.boolean().optional().describe("Draw the horizontal axis."),
        displayHidden: z.boolean().optional().describe("Plot data in hidden rows and columns."),
        colors: z
          .object({
            series: colorSchema.optional(),
            negative: colorSchema.optional(),
            axis: colorSchema.optional(),
            markers: colorSchema.optional(),
            first: colorSchema.optional(),
            last: colorSchema.optional(),
            high: colorSchema.optional(),
            low: colorSchema.optional(),
          })
          .optional(),
        lineWeight: z.number().positive().optional().describe("Line width in points."),
        emptyCells: z
          .enum(["gap", "zero", "span"])
          .optional()
          .describe("How empty cells plot: a gap, zero, or a line spanning them."),
        minAxis: axisBoundSchema.optional(),
        maxAxis: axisBoundSchema.optional(),
        dateAxis: z
          .string()
          .optional()
          .describe("Date range that spaces the points by date; an empty string removes it."),
      },
    },
    ({ sessionId, sheet, operation, index, sparklines, ...options }) => {
      try {
        if (operation === "list") {
          return ok(listSessionSparklines(sessionId, sheet));
        }
        if (operation === "add") {
          if (sparklines === undefined) {
            throw new Error("sparklines are required for add");
          }
          return ok(addSessionSparklines(sessionId, sheet, sparklines, options));
        }
        if (index === undefined) {
          throw new Error(`index is required for ${operation}`);
        }
        if (operation === "remove") {
          return ok(removeSessionSparklines(sessionId, sheet, index));
        }
        return ok(updateSessionSparklines(sessionId, sheet, index, sparklines, options));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_data_table_operation",
    {
      title: "What-If data table operation",
      description:
        "Get, add, or remove a What-If data table ({=TABLE(row_input,col_input)}), or recompute every data table. The formula sits in the row above the result range (column input) or the column left of it (row input), with substitution values along the other edge; a two-variable table has the formula at the top-left corner outside the range. Result cells refuse writes until the table is removed. Add recalculates so the results are readable at once.",
      inputSchema: {
        ...sheetInputSchema,
        operation: z.enum(["get", "add", "remove", "recalc"]),
        range: z
          .string()
          .optional()
          .describe("Result range on the selected sheet, for example C3:E8; required for add."),
        rowInput: z
          .string()
          .optional()
          .describe("Input cell that receives the values in the row above the range."),
        colInput: z
          .string()
          .optional()
          .describe("Input cell that receives the values in the column left of the range."),
        cell: z
          .string()
          .optional()
          .describe("Any result cell of the table; required for get and remove."),
      },
    },
    ({ sessionId, sheet, operation, range, rowInput, colInput, cell }) => {
      try {
        if (operation === "recalc") {
          return ok(recalcSessionDataTables(sessionId));
        }
        if (operation === "add") {
          if (range === undefined) {
            throw new Error("range is required for add");
          }
          return ok(addSessionDataTable(sessionId, sheet, { range, rowInput, colInput }));
        }
        if (cell === undefined) {
          throw new Error(`cell is required for ${operation}`);
        }
        return ok(
          operation === "get"
            ? getSessionDataTable(sessionId, sheet, cell)
            : removeSessionDataTable(sessionId, sheet, cell),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );
}
