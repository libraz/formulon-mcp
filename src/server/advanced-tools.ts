import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { callWorkbookMethod, sessionProfileId } from "../sessions.js";
import { fail, methodOk, ok, profileSchema } from "./shared.js";

const jsonArgsSchema = z.array(z.unknown()).default([]);

export function registerAdvancedTools(server: McpServer): void {
  server.registerTool(
    "formulon_function_lookup",
    {
      title: "Function and locale lookup",
      description:
        "List functions, read name/arity/availability metadata, translate a function name or a whole formula between the stored English form and a profile's spelling, or read a profile's locale facts (separators, boolean names, date order, currency, error names).",
      inputSchema: {
        sessionId: z.string(),
        operation: z.enum([
          "names",
          "metadata",
          "localize",
          "canonicalize",
          "localizeFormula",
          "canonicalizeFormula",
          "localeFacts",
        ]),
        name: z
          .string()
          .optional()
          .describe("Function name for metadata, localize and canonicalize."),
        formula: z
          .string()
          .optional()
          .describe("Formula text for localizeFormula and canonicalizeFormula."),
        profile: profileSchema
          .optional()
          .describe("Profile to translate into or describe; defaults to the session's profile."),
      },
    },
    ({ sessionId, operation, name, formula, profile }) => {
      try {
        if (operation === "names") {
          return methodOk(sessionId, "functionNames", []);
        }
        const profileId = profile ?? sessionProfileId(sessionId);
        if (operation === "localeFacts") {
          return methodOk(sessionId, "localeFacts", [profileId]);
        }
        if (operation === "localizeFormula" || operation === "canonicalizeFormula") {
          if (formula === undefined) {
            throw new Error(`formula is required for ${operation}`);
          }
          return methodOk(sessionId, operation, [formula, profileId]);
        }
        if (!name) {
          throw new Error("name is required");
        }
        if (operation === "metadata") {
          return methodOk(sessionId, "functionMetadata", [name]);
        }
        return methodOk(
          sessionId,
          operation === "localize" ? "localizeFunctionName" : "canonicalizeFunctionName",
          [name, profileId],
        );
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_workbook_call",
    {
      title: "Call Formulon Workbook method",
      description:
        "Low-level allowlisted access to the Formulon Workbook API for advanced features. Use formulon_workbook_api first to discover method signatures, input records and enum values.",
      inputSchema: {
        sessionId: z.string(),
        method: z
          .string()
          .describe(
            "Allowlisted Workbook method name, for example getMerges, addMerge, pivotLayout.",
          ),
        args: jsonArgsSchema.describe("Positional JSON arguments passed to the Workbook method."),
      },
    },
    ({ sessionId, method, args }) => {
      try {
        return ok(callWorkbookMethod(sessionId, method, args));
      } catch (error) {
        return fail(error);
      }
    },
  );
}
