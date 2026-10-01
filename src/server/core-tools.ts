import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { assertStatus, formulonModule, normalizeFormula, valueToJson } from "../formulon.js";
import { resolveSessionSheet } from "../session/workbook.js";
import { fail, methodOk, ok, sheetRefSchema } from "./shared.js";

export function registerCoreTools(server: McpServer, version: string): void {
  server.registerTool(
    "formulon_version",
    {
      title: "Formulon version",
      description: "Return the loaded Formulon engine version.",
      inputSchema: {},
    },
    async () => {
      try {
        const module = await formulonModule();
        return ok({ version: module.versionString(), serverVersion: version });
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_eval_formula",
    {
      title: "Evaluate formula",
      description:
        "Evaluate one Excel formula with Formulon. With sessionId, evaluates read-only against the open workbook (resolving refs, defined names, and ROW()/COLUMN() anchored at the given cell) without mutating it.",
      inputSchema: {
        formula: z.string().describe("Excel formula, with or without a leading '='."),
        sessionId: z
          .string()
          .optional()
          .describe("Evaluate against this open workbook session instead of a fresh workbook."),
        sheet: sheetRefSchema.default(0),
        row: z
          .number()
          .int()
          .nonnegative()
          .default(0)
          .describe("Anchor row for ROW()/COLUMN() and relative refs (session mode)."),
        col: z
          .number()
          .int()
          .nonnegative()
          .default(0)
          .describe("Anchor column for ROW()/COLUMN() and relative refs (session mode)."),
      },
    },
    async ({ formula, sessionId, sheet, row, col }) => {
      try {
        if (sessionId !== undefined) {
          const sheetIndex = resolveSessionSheet(sessionId, sheet);
          return methodOk(sessionId, "evaluateFormulaText", [
            sheetIndex,
            row,
            col,
            normalizeFormula(formula),
          ]);
        }
        if (typeof sheet === "string") {
          throw new Error(
            "sheet names require sessionId; global evaluation has only the default sheet",
          );
        }
        const module = await formulonModule();
        const result = module.evalFormula(normalizeFormula(formula));
        assertStatus(result.status, "evaluate formula");
        return ok({
          formula: normalizeFormula(formula),
          status: result.status,
          value: valueToJson(result.value),
        });
      } catch (error) {
        return fail(error);
      }
    },
  );
}
