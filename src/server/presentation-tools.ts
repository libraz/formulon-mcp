import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  colorSchema,
  fontStyleSchema,
  printSettingsSchema,
  styleSchema,
} from "../layout/schemas.js";
import { PRINT_PRESET_NAMES } from "../print.js";
import {
  buildSessionDocument,
  sessionDefaultFont,
  sessionDimension,
  sessionPrintSettings,
  styleSessionRange,
} from "../sessions.js";
import { STYLE_VOCABULARY } from "../styles.js";
import { fail, ok, sheetInputSchema, sheetRefSchema } from "./shared.js";

const cellLiteralSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);

const formatFieldSchema = z
  .string()
  .describe(
    "Excel format code, or one of the aliases date, datetime, time, number, decimal, percent. An ISO date string under a date format is stored as a real date, not text.",
  );

const documentBlockSchema = z.discriminatedUnion("type", [
  z.object({
    type: z.literal("title"),
    text: z.string(),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    align: z.enum(["left", "center", "right"]).optional(),
    size: z.number().positive().optional(),
    bold: z.boolean().optional(),
  }),
  z.object({
    type: z.literal("text"),
    text: z.string(),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    align: z.enum(["left", "center", "right"]).optional(),
    size: z.number().positive().optional(),
    bold: z.boolean().optional(),
    span: z.number().int().positive().optional(),
    wrap: z.boolean().optional(),
    rows: z.number().int().positive().optional(),
  }),
  z.object({
    type: z.literal("fields"),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    align: z.enum(["left", "right"]).optional(),
    labelSpan: z.number().int().positive().optional(),
    valueSpan: z.number().int().positive().optional(),
    rule: z.boolean().optional().describe("Rule each value cell underneath, like a paper form."),
    items: z
      .array(
        z.object({
          label: z.string(),
          value: cellLiteralSchema.optional(),
          formula: z.string().optional(),
          format: formatFieldSchema.optional(),
          name: z.string().optional(),
        }),
      )
      .min(1),
  }),
  z.object({
    type: z.literal("table"),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    bandColor: colorSchema.optional().describe("Fill applied to every other body row."),
    rowCount: z
      .number()
      .int()
      .nonnegative()
      .optional()
      .describe("Blank ruled rows when `rows` is omitted — a form to fill in."),
    columns: z
      .array(
        z.object({
          header: z.string(),
          key: z.string().optional().describe("Name used in formulas; defaults to the header."),
          formula: z
            .string()
            .optional()
            .describe("Per-row formula; {key} binds to that column's cell in the same row."),
          width: z.number().positive().optional(),
          format: formatFieldSchema.optional(),
          align: z.enum(["left", "center", "right"]).optional(),
        }),
      )
      .min(1),
    rows: z
      .array(z.union([z.array(cellLiteralSchema), z.record(z.string(), cellLiteralSchema)]))
      .optional()
      .describe("Row objects keyed by column key/header, or positional arrays."),
  }),
  z.object({
    type: z.literal("summary"),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    align: z.enum(["left", "right"]).optional(),
    labelSpan: z.number().int().positive().optional(),
    valueSpan: z.number().int().positive().optional(),
    border: z.boolean().optional(),
    items: z
      .array(
        z.object({
          label: z.string(),
          value: cellLiteralSchema.optional(),
          formula: z.string().optional(),
          format: formatFieldSchema.optional(),
          emphasis: z.boolean().optional(),
          name: z.string().optional(),
        }),
      )
      .min(1),
  }),
  z.object({
    type: z.literal("spacer"),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    rows: z.number().int().nonnegative().optional(),
  }),
]);

export function registerPresentationTools(server: McpServer): void {
  server.registerTool(
    "formulon_dimension_operation",
    {
      title: "Column/row dimension operation",
      description:
        "List column-width/row-height overrides, or set width/height, hidden, or outline level. Columns act on an inclusive [first, last] span; rows act on an inclusive [row, row] span. Column sizes default to Excel character units; row sizes default to points. Use px or mm when a physical display size is easier to specify.",
      inputSchema: {
        ...sheetInputSchema,
        axis: z.enum(["column", "row"]),
        operation: z.enum(["list", "size", "hidden", "outline"]),
        first: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe("Zero-based first column index (column axis)."),
        last: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe("Zero-based last column index; defaults to first (column axis)."),
        row: z.number().int().nonnegative().optional().describe("Zero-based row index (row axis)."),
        size: z
          .number()
          .optional()
          .describe("Column width or row height for the 'size' operation."),
        unit: z
          .enum(["chars", "pt", "px", "mm"])
          .optional()
          .describe(
            "Size unit: column defaults to chars, row defaults to pt; px and mm convert through 96-DPI display points.",
          ),
        includeGeometry: z
          .boolean()
          .optional()
          .describe(
            "For list, include derived points/pixels/mm geometry for each stored override.",
          ),
        hidden: z.boolean().optional().describe("Hidden flag for the 'hidden' operation."),
        level: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe("Outline level for the 'outline' operation."),
      },
    },
    ({
      sessionId,
      sheet,
      axis,
      operation,
      first,
      last,
      row,
      size,
      unit,
      hidden,
      level,
      includeGeometry,
    }) => {
      try {
        return ok(
          sessionDimension(sessionId, axis, operation, {
            sheet,
            first,
            last,
            row,
            size,
            unit,
            hidden,
            level,
            includeGeometry,
          }),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_build_document",
    {
      title: "Build a document from blocks",
      description:
        'Lay out a document as a vertical stack of blocks — title, text, fields, table, summary, spacer — and write it in one call. Positions, ruling, number formats, column widths, merges and the print area are all resolved from the layout, so no row arithmetic is needed. Blocks reference each other by name: a table column registers as {table.<header>} over its body range, and a fields or summary item registers under its label, so "=SUM({table.Amount})" and "={Subtotal}+{Tax}" bind to the right cells once the layout is known. The response maps every block and name to A1 so the result can be refined with set_cells / style_range / print_settings. This tool carries no document semantics: labels, tax rules and totals are the caller\'s formulas.',
      inputSchema: {
        sessionId: z.string(),
        sheet: sheetRefSchema
          .optional()
          .describe("Sheet used when `start` has no Sheet!-prefix; defaults to the first sheet."),
        start: z
          .string()
          .default("B2")
          .describe("Top-left anchor, for example Sheet1!B2. Defaults to B2, leaving a margin."),
        width: z
          .number()
          .int()
          .positive()
          .optional()
          .describe("Document width in columns; defaults to the widest table's column count."),
        blocks: z.array(documentBlockSchema).min(1),
        theme: z
          .object({
            font: z.string().optional(),
            size: z.number().positive().optional(),
            accent: colorSchema
              .optional()
              .describe("Header-band fill as #RRGGBB. Omit for a plain ruled header."),
            headerText: z.string().optional(),
            border: z
              .enum(STYLE_VOCABULARY.borderStyle)
              .optional()
              .describe("Grid border style; defaults to thin."),
            outline: z
              .enum(STYLE_VOCABULARY.borderStyle)
              .optional()
              .describe("Outer border style; defaults to medium."),
          })
          .optional(),
        print: z
          .enum(PRINT_PRESET_NAMES)
          .optional()
          .describe("Named page setup; also sets the print area to the document's extent."),
        repeatTableHeader: z
          .boolean()
          .optional()
          .describe("Repeat the first table's header row on every printed page. Defaults to true."),
      },
    },
    ({ sessionId, sheet, ...spec }) => {
      try {
        return ok(
          buildSessionDocument(
            sessionId,
            spec as Parameters<typeof buildSessionDocument>[1],
            sheet,
          ),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_style_range",
    {
      title: "Style a range",
      description:
        "Apply fonts, fills, borders, number formats, and alignment across an A1 range, materializing blank cells so an empty ruled box renders. Each property is a delta: cells keep whatever the style does not state, so ruling a table does not undo the number format already on its amount column. Style in passes — the grid over the whole table, then the header row, then a total cell.",
      inputSchema: {
        sessionId: z.string(),
        range: z.string().describe("A1 range or single cell, for example Sheet1!B2:F20 or B2."),
        sheet: sheetRefSchema
          .optional()
          .describe("Sheet used when range has no Sheet!-prefix; defaults to the first sheet."),
        style: styleSchema,
        baseOn: z
          .enum(["existing", "default"])
          .default("existing")
          .describe(
            "Start from the style already on the range's top-left cell, or from the workbook default.",
          ),
      },
    },
    ({ sessionId, range, sheet, style, baseOn }) => {
      try {
        return ok(styleSessionRange(sessionId, range, style, sheet, baseOn));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_default_font",
    {
      title: "Workbook default font",
      description:
        "Read or redeclare the workbook default font — the one every cell that was never styled resolves to. A new workbook is seeded with Excel's Calibri 11, which renders Japanese text in a fallback face, so a Japanese document should declare its own default once instead of styling every cell. Omit `font` to read. Stating a `name` also cuts the theme link a loaded workbook's base font carries, since a theme-linked font is re-resolved from the theme rather than keeping the name.",
      inputSchema: {
        sessionId: z.string(),
        font: fontStyleSchema
          .optional()
          .describe(
            "Properties to change; anything omitted keeps its current value. Omit to read.",
          ),
      },
    },
    ({ sessionId, font }) => {
      try {
        return ok(sessionDefaultFont(sessionId, font));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_print_settings",
    {
      title: "Print settings",
      description:
        "Read or set a sheet's print settings: page setup, margins, print options, header/footer, print area, print titles, and manual page breaks. Omit every setting to read. Reads report the resulting pageCount, so a page layout can be checked without saving the file.",
      inputSchema: {
        ...sheetInputSchema,
        ...printSettingsSchema.shape,
      },
    },
    ({ sessionId, sheet, ...settings }) => {
      try {
        const stated = Object.fromEntries(
          Object.entries(settings).filter(([, value]) => value !== undefined),
        );
        return ok(
          sessionPrintSettings(
            sessionId,
            sheet,
            Object.keys(stated).length > 0 ? stated : undefined,
          ),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );
}
