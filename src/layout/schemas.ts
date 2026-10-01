import { z } from "zod";
import { STYLE_VOCABULARY } from "../styles.js";

const HEX_COLOR = /^#?(?:[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;

/** A colour accepted by style and print-facing tools. */
export const colorSchema = z
  .string()
  .regex(HEX_COLOR, "expected a hexadecimal colour in #RRGGBB or #AARRGGBB form");

export const borderSideSchema = z
  .union([
    z.enum(STYLE_VOCABULARY.borderStyle),
    z.object({
      style: z.enum(STYLE_VOCABULARY.borderStyle).optional(),
      color: colorSchema.optional(),
    }),
  ])
  .describe(
    `Border style name (${STYLE_VOCABULARY.borderStyle.join(", ")}), or {style, color} with a #RRGGBB color.`,
  );

export const fontStyleSchema = z.object({
  name: z.string().optional(),
  size: z.number().finite().positive().optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  strike: z.boolean().optional(),
  underline: z.enum(STYLE_VOCABULARY.underline).optional(),
  vertAlign: z.enum(STYLE_VOCABULARY.vertAlign).optional(),
  color: colorSchema.optional().describe("#RRGGBB or #AARRGGBB."),
});

export const styleSchema = z.object({
  font: fontStyleSchema.optional(),
  fill: z
    .object({
      color: colorSchema.optional().describe("Fill color as #RRGGBB; implies a solid pattern."),
      bgColor: colorSchema.optional(),
      pattern: z.enum(STYLE_VOCABULARY.fillPattern).optional(),
    })
    .optional(),
  border: z
    .object({
      all: borderSideSchema.optional().describe("Rules every cell in the range on all four sides."),
      outline: borderSideSchema.optional().describe("Draws a box around the range only."),
      left: borderSideSchema.optional(),
      right: borderSideSchema.optional(),
      top: borderSideSchema.optional(),
      bottom: borderSideSchema.optional(),
    })
    .optional(),
  numberFormat: z
    .string()
    .optional()
    .describe('Excel format code such as "#,##0" or "yyyy/mm/dd". Empty string means General.'),
  align: z
    .object({
      horizontal: z.enum(STYLE_VOCABULARY.horizontalAlign).optional(),
      vertical: z.enum(STYLE_VOCABULARY.verticalAlign).optional(),
      wrapText: z.boolean().optional(),
      indent: z.number().int().min(0).max(255).optional(),
      textRotation: z
        .union([z.number().int().min(0).max(180), z.literal(255)])
        .optional()
        .describe("Text rotation in degrees, or 255 for vertical text."),
      shrinkToFit: z.boolean().optional(),
    })
    .optional(),
});

const pageSetupSchema = z.object({
  orientation: z.enum(["default", "portrait", "landscape"]).optional(),
  paperSize: z
    .number()
    .int()
    .nonnegative()
    .optional()
    .describe("OOXML paper-size code: 9 = A4, 8 = A3, 11 = A5, 1 = Letter."),
  scale: z.number().int().min(10).max(400).optional(),
  fitToWidth: z.number().int().nonnegative().optional(),
  fitToHeight: z.number().int().nonnegative().optional(),
  fitToPage: z
    .boolean()
    .optional()
    .describe(
      "Selects fit-to-page mode; fitToWidth/fitToHeight state the target, so all three are needed to fit onto one page.",
    ),
});

const marginsSchema = z
  .object({
    left: z.number().finite().nonnegative().optional(),
    right: z.number().finite().nonnegative().optional(),
    top: z.number().finite().nonnegative().optional(),
    bottom: z.number().finite().nonnegative().optional(),
    header: z.number().finite().nonnegative().optional(),
    footer: z.number().finite().nonnegative().optional(),
  })
  .describe("Page margins in inches.");

const printOptionsSchema = z.object({
  gridLines: z.boolean().optional(),
  headings: z.boolean().optional(),
  horizontalCentered: z.boolean().optional(),
  verticalCentered: z.boolean().optional(),
});

const headerFooterSchema = z
  .object({
    oddHeader: z.string().optional(),
    oddFooter: z.string().optional(),
    evenHeader: z.string().optional(),
    evenFooter: z.string().optional(),
    firstHeader: z.string().optional(),
    firstFooter: z.string().optional(),
    differentOddEven: z.boolean().optional(),
    differentFirst: z.boolean().optional(),
    scaleWithDoc: z.boolean().optional(),
    alignWithMargins: z.boolean().optional(),
  })
  .describe(
    'Section text uses Excel\'s own codes — &L/&C/&R pick the section, &P the page number, &N the page count, &D the date, and && a literal ampersand (for example \'&C&"MS Gothic"Invoice &P/&N\'). Each section is tri-state: omit to leave it, "" to clear it.',
  );

/** The flat print-settings input shared by read/write and batch layout tools. */
export const printSettingsSchema = z.object({
  pageSetup: pageSetupSchema.optional(),
  margins: marginsSchema.optional(),
  printOptions: printOptionsSchema.optional(),
  headerFooter: headerFooterSchema.optional(),
  printArea: z
    .string()
    .optional()
    .describe(
      'Comma-separated A1 ranges, for example "A1:F40" or "A1:B10,D5:E20". Empty string removes the print area.',
    ),
  printTitles: z
    .object({
      repeatRows: z.string().optional().describe('Whole-row span such as "1:2".'),
      repeatCols: z.string().optional().describe('Whole-column span such as "A:A".'),
    })
    .optional()
    .describe("Rows and columns repeated on every printed page. Both empty removes them."),
  rowBreaks: z
    .array(z.number().int().nonnegative())
    .max(10_000)
    .optional()
    .describe("Zero-based rows each manual break precedes. Replaces the sheet's breaks."),
  colBreaks: z
    .array(z.number().int().nonnegative())
    .max(10_000)
    .optional()
    .describe("Zero-based columns each manual break precedes. Replaces the sheet's breaks."),
  pageSetupXml: z.string().optional().describe("Raw <pageSetup> fragment; empty removes it."),
  pageMarginsXml: z.string().optional().describe("Raw <pageMargins> fragment; empty removes it."),
  printOptionsXml: z.string().optional().describe("Raw <printOptions> fragment; empty removes it."),
  headerFooterXml: z.string().optional().describe("Raw <headerFooter> fragment; empty removes it."),
  sheetPrXml: z.string().optional().describe("Raw <sheetPr> fragment; empty removes it."),
});

/** The write-only variant; an empty object would otherwise become a read. */
export const nonEmptyPrintSettingsSchema = printSettingsSchema.refine(
  (settings) => Object.values(settings).some((value) => value !== undefined),
  "at least one print setting is required",
);
