import { enumName } from "@libraz/formulon";
import { z } from "zod";
import type { Workbook } from "../formulon.js";
import { errorText, FormulonError, jsonText, PROFILE_ID_PATTERN, StatusCode } from "../formulon.js";
import { callLabeled } from "../sessions.js";

export function ok(value: unknown) {
  return {
    content: [{ type: "text" as const, text: jsonText(value) }],
  };
}

/**
 * The MCP error result of a failed tool call. An engine `FormulonError` reads as its status name,
 * message and context, and carries those fields (with `parseError` for a formula syntax error) as
 * a second JSON content item.
 */
export function fail(error: unknown) {
  const content = [{ type: "text" as const, text: errorText(error) }];
  if (error instanceof FormulonError) {
    content.push({
      type: "text" as const,
      text: jsonText({
        code: enumName(StatusCode, error.code) ?? error.code,
        message: error.message,
        context: error.context,
        ...(error.parseError ? { parseError: error.parseError } : {}),
      }),
    });
  }
  return { isError: true, content };
}

/** The `ok` result of `callLabeled`: one named engine operation on the session's workbook. */
export function labeledOk(
  sessionId: string,
  method: string,
  access: "read" | "write",
  operation: (workbook: Workbook) => unknown,
) {
  return ok(callLabeled(sessionId, method, access, operation));
}

export const sheetRefSchema = z.union([z.number().int().nonnegative(), z.string()]);

export const profileSchema = z
  .string()
  .regex(PROFILE_ID_PATTERN)
  .describe(
    "Excel profile id: mac or win host with en_US, ja_JP, de_DE, fr_FR, es_ES, es_MX, pt_BR, it_IT, nl_NL, ru_RU, zh_CN, zh_TW, ko_KR or th_TH, for example win-365-ja_JP.",
  );

export const rangeSchema = z.object({
  firstRow: z.number().int().nonnegative(),
  firstCol: z.number().int().nonnegative(),
  lastRow: z.number().int().nonnegative(),
  lastCol: z.number().int().nonnegative(),
});

/** A drawing length (offset, size, line width) in points, 96-DPI pixels or millimetres. */
export const imageLengthSchema = z.object({
  size: z.number().finite().nonnegative(),
  unit: z.enum(["pt", "px", "mm"]).optional().describe("Defaults to points."),
});

export const sheetInputSchema = {
  sessionId: z.string(),
  sheet: sheetRefSchema
    .default(0)
    .describe("Zero-based sheet index or sheet name; defaults to the first sheet."),
};
