import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { callWorkbookMethod } from "../sessions.js";
import { fail, methodOk, ok } from "./shared.js";

const jsonArgsSchema = z.array(z.unknown()).default([]);

export function registerAdvancedTools(server: McpServer): void {
  server.registerTool(
    "formulon_function_lookup",
    {
      title: "Function lookup",
      description:
        "List functions, read name/arity/availability metadata, localize names, or canonicalize localized names.",
      inputSchema: {
        sessionId: z.string(),
        operation: z.enum(["names", "metadata", "localize", "canonicalize"]),
        name: z.string().optional(),
        locale: z.number().int().nonnegative().default(0),
      },
    },
    ({ sessionId, operation, name, locale }) => {
      try {
        if (operation === "names") {
          return methodOk(sessionId, "functionNames", []);
        }
        if (!name) {
          throw new Error("name is required");
        }
        if (operation === "metadata") {
          return methodOk(sessionId, "functionMetadata", [name, locale]);
        }
        return methodOk(
          sessionId,
          operation === "localize" ? "localizeFunctionName" : "canonicalizeFunctionName",
          [name, locale],
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
