import { z } from "zod";
import { MAX_COL, MAX_ROW } from "../session/shared.js";

const range = z
  .object({
    firstRow: z.number().int().min(0).max(MAX_ROW),
    firstCol: z.number().int().min(0).max(MAX_COL),
    lastRow: z.number().int().min(0).max(MAX_ROW),
    lastCol: z.number().int().min(0).max(MAX_COL),
  })
  .passthrough()
  .refine(
    (value) => value.firstRow <= value.lastRow && value.firstCol <= value.lastCol,
    "Range endpoints must be in top-left to bottom-right order.",
  );

export const validationSchema = z
  .object({
    ranges: z.array(range).min(1).describe("Inclusive zero-based ranges to validate."),
    type: z
      .number()
      .int()
      .min(0)
      .max(7)
      .describe("0 none, 1 whole, 2 decimal, 3 list, 4 date, 5 time, 6 textLength, 7 custom."),
    op: z
      .number()
      .int()
      .min(0)
      .max(7)
      .optional()
      .describe(
        "0 between, 1 notBetween, 2 equal, 3 notEqual, 4 greaterThan, 5 lessThan, 6 greaterThanOrEqual, 7 lessThanOrEqual.",
      ),
    errorStyle: z
      .number()
      .int()
      .min(0)
      .max(2)
      .optional()
      .describe("0 stop, 1 warning, 2 information."),
    allowBlank: z.boolean().optional(),
    showInputMessage: z.boolean().optional(),
    showErrorMessage: z.boolean().optional(),
    showDropDown: z
      .boolean()
      .optional()
      .describe("Whether the list dropdown arrow is shown; defaults to true."),
    formula1: z
      .string()
      .optional()
      .describe(
        "First bound, list formula or custom expression. A literal list uses Excel quotes, for example '\"Draft,Approved\"'.",
      ),
    formula2: z.string().optional().describe("Second bound for between/notBetween."),
    errorTitle: z.string().optional(),
    errorMessage: z.string().optional(),
    promptTitle: z.string().optional(),
    promptMessage: z.string().optional(),
  })
  .passthrough();

const channel = z.number().int().min(0).max(255);
const color = z
  .object({ r: channel, g: channel, b: channel, a: channel.optional() })
  .passthrough()
  .describe("RGBA channels from 0 to 255; omitted alpha is opaque.");
const threshold = z
  .object({
    type: z
      .number()
      .int()
      .min(0)
      .max(7)
      .describe(
        "0 number, 1 percent, 2 percentile, 3 min, 4 max, 5 formula, 6 autoMin, 7 autoMax.",
      ),
    value: z.string().optional().describe("Numeric text or formula, depending on type."),
    gte: z.boolean().optional().describe("Include equality at this threshold; defaults to true."),
  })
  .passthrough();

export const conditionalFormatSchema = z
  .object({
    sqref: z.array(range).min(1).describe("Inclusive zero-based ranges to format."),
    type: z
      .number()
      .int()
      .min(0)
      .max(17)
      .describe(
        "0 expression, 1 cellIs, 2 colorScale, 3 dataBar, 4 iconSet, 5 top10, 6 aboveAverage, 7 containsText, 8 notContainsText, 9 beginsWith, 10 endsWith, 11 containsBlanks, 12 notContainsBlanks, 13 containsErrors, 14 notContainsErrors, 15 timePeriod, 16 duplicateValues, 17 uniqueValues.",
      ),
    priority: z
      .number()
      .int()
      .optional()
      .describe("Lower priorities run first; omitted/nonpositive assigns the next priority."),
    stopIfTrue: z.boolean().optional(),
    id: z
      .string()
      .regex(
        /^(?:|\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\})$/,
      )
      .optional()
      .describe("Optional rule GUID in {8-4-4-4-12} hexadecimal form; empty assigns an id."),
    dxfId: z
      .number()
      .int()
      .nonnegative()
      .optional()
      .describe("Differential style id returned by addDxf."),
    formula1: z.string().optional(),
    formula2: z.string().optional(),
    op: z
      .number()
      .int()
      .min(0)
      .max(7)
      .optional()
      .describe(
        "cellIs: 0 lessThan, 1 lessThanOrEqual, 2 equal, 3 notEqual, 4 greaterThanOrEqual, 5 greaterThan, 6 between, 7 notBetween.",
      ),
    rank: z.number().int().positive().optional(),
    percent: z.boolean().optional(),
    bottom: z.boolean().optional(),
    aboveAverage: z.boolean().optional(),
    equalAverage: z.boolean().optional(),
    stdDev: z.number().int().nonnegative().optional(),
    text: z.string().optional(),
    timePeriod: z
      .number()
      .int()
      .min(0)
      .max(9)
      .optional()
      .describe(
        "0 today, 1 yesterday, 2 tomorrow, 3 last7Days, 4 thisWeek, 5 lastWeek, 6 nextWeek, 7 thisMonth, 8 lastMonth, 9 nextMonth.",
      ),
    colorScale: z
      .object({
        thresholds: z.array(threshold).min(2).max(3),
        colors: z.array(color).min(2).max(3),
      })
      .passthrough()
      .refine(
        (value) => value.thresholds.length === value.colors.length,
        "Color-scale threshold and color counts must match.",
      )
      .optional(),
    dataBar: z
      .object({
        min: threshold,
        max: threshold,
        fill: color,
        showValue: z.boolean().optional(),
        minLengthPct: z.number().int().min(0).max(100).optional(),
        maxLengthPct: z.number().int().min(0).max(100).optional(),
        gradient: z.boolean().optional(),
        axisPosition: z
          .number()
          .int()
          .min(0)
          .max(2)
          .optional()
          .describe("0 automatic, 1 middle, 2 none."),
        negativeFill: color.optional(),
        border: color.optional(),
        negativeBorder: color.optional(),
        axisColor: color.optional(),
        direction: z
          .number()
          .int()
          .min(0)
          .max(2)
          .optional()
          .describe("0 context, 1 left to right, 2 right to left."),
      })
      .passthrough()
      .refine(
        (value) => (value.minLengthPct ?? 10) <= (value.maxLengthPct ?? 90),
        "Data-bar minimum length must not exceed its maximum length.",
      )
      .optional(),
    iconSet: z
      .object({
        name: z
          .number()
          .int()
          .min(0)
          .max(16)
          .describe(
            "0 3Arrows, 1 3ArrowsGray, 2 3Flags, 3 3TrafficLights1, 4 3TrafficLights2, 5 3Signs, 6 3Symbols, 7 3Symbols2, 8 4Arrows, 9 4ArrowsGray, 10 4RedToBlack, 11 4Rating, 12 4TrafficLights, 13 5Arrows, 14 5ArrowsGray, 15 5Rating, 16 5Quarters.",
          ),
        thresholds: z
          .array(threshold)
          .min(2)
          .max(4)
          .describe("One threshold between each adjacent pair of icons (2, 3 or 4)."),
        reverse: z.boolean().optional(),
        showValue: z.boolean().optional(),
        percent: z.boolean().optional(),
        floor: threshold
          .optional()
          .describe(
            "Lower bound of the lowest icon bucket; values below it have no icon. Defaults to percent 0.",
          ),
      })
      .passthrough()
      .refine(
        (value) => value.thresholds.length === (value.name <= 7 ? 2 : value.name <= 12 ? 3 : 4),
        "Icon-set thresholds must match its icon count minus one.",
      )
      .optional(),
  })
  .passthrough()
  .superRefine((rule, context) => {
    const payload =
      rule.type === 2
        ? "colorScale"
        : rule.type === 3
          ? "dataBar"
          : rule.type === 4
            ? "iconSet"
            : undefined;
    if (payload && !rule[payload]) {
      context.addIssue({
        code: "custom",
        message: `${payload} is required for rule type ${rule.type}`,
        path: [payload],
      });
    }
  });
