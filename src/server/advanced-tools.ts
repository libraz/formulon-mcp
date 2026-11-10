import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { formulonModule, type ProfileId } from "../formulon.js";
import { callWorkbookMethod, sessionProfileId } from "../sessions.js";
import { fail, labeledOk, ok, profileSchema } from "./shared.js";

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
    async ({ sessionId, operation, name, formula, profile }) => {
      try {
        const engine = await formulonModule();
        if (operation === "names") {
          return labeledOk(sessionId, "Formulon.listFunctionNames", "read", () =>
            engine.listFunctionNames(),
          );
        }
        const profileId = (profile as ProfileId | undefined) ?? sessionProfileId(sessionId);
        if (operation === "localeFacts") {
          return labeledOk(sessionId, "Formulon.getLocaleFacts", "read", () =>
            engine.getLocaleFacts(profileId),
          );
        }
        if (operation === "localizeFormula" || operation === "canonicalizeFormula") {
          if (formula === undefined) {
            throw new Error(`formula is required for ${operation}`);
          }
          return labeledOk(sessionId, `Formulon.${operation}`, "read", () =>
            engine[operation](formula, profileId),
          );
        }
        if (!name) {
          throw new Error("name is required");
        }
        if (operation === "metadata") {
          return labeledOk(sessionId, "Formulon.getFunctionMetadata", "read", () =>
            engine.getFunctionMetadata(name),
          );
        }
        const method =
          operation === "localize" ? "localizeFunctionName" : "canonicalizeFunctionName";
        return labeledOk(sessionId, `Formulon.${method}`, "read", () =>
          engine[method](name, profileId),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_workbook_call",
    {
      title: "Call a Formulon engine method",
      description:
        'Low-level access to the installed Formulon engine for features no task tool covers. `target` names the object the method belongs to: a handle kind (Workbook, Sheet, Table, Pivot, PivotCache, Chart, ...) and the key of each accessor on the way, or Formulon for the module functions of the engine (collation, locale facts, formula translation), for example { kind: "Pivot", sheet: 0, pivot: "Sales" }. Use formulon_workbook_api first: each method it lists carries its target kind and keys, its signature, input records and enum values. Raw bytes, callbacks and saving are not available here.',
      inputSchema: {
        sessionId: z.string(),
        target: z
          .object({ kind: z.string() })
          .catchall(z.union([z.number(), z.string()]))
          .optional()
          .describe(
            "The handle the method belongs to: { kind } plus that kind's keys as formulon_workbook_api lists them. Omitted, the workbook.",
          ),
        method: z
          .string()
          .describe("Method name on the target, for example listMerges on a Sheet."),
        args: jsonArgsSchema.describe("Positional JSON arguments passed to the method."),
      },
    },
    ({ sessionId, target, method, args }) => {
      try {
        return ok(callWorkbookMethod(sessionId, target, method, args));
      } catch (error) {
        return fail(error);
      }
    },
  );
}
