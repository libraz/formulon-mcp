import { z } from "zod";
import { jsonText, PROFILE_ID_PATTERN } from "../formulon.js";
import { callWorkbookMethod } from "../sessions.js";

export function ok(value: unknown) {
  return {
    content: [{ type: "text" as const, text: jsonText(value) }],
  };
}

export function fail(error: unknown) {
  const message = error instanceof Error ? error.message : String(error);
  return {
    isError: true,
    content: [{ type: "text" as const, text: message }],
  };
}

export function methodOk(sessionId: string, method: string, args: unknown[]) {
  return ok(callWorkbookMethod(sessionId, method, args));
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

export const sheetInputSchema = {
  sessionId: z.string(),
  sheet: sheetRefSchema
    .default(0)
    .describe("Zero-based sheet index or sheet name; defaults to the first sheet."),
};
