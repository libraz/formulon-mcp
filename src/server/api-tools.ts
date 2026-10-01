import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { describeWorkbookApi, searchWorkbookApi, WORKBOOK_API_MAX_LIMIT } from "../workbook-api.js";
import { fail, ok } from "./shared.js";

export function registerApiTools(server: McpServer): void {
  server.registerTool(
    "formulon_workbook_api",
    {
      title: "Discover Formulon Workbook API",
      description:
        "Search or describe the installed Formulon Workbook methods. This tool is session-free and returns signatures, source documentation, read/write access, and referenced TypeScript declarations for AI callers.",
      inputSchema: {
        operation: z.enum(["search", "describe"]).default("search"),
        query: z
          .string()
          .optional()
          .describe(
            "Case-insensitive substring matched against method names, signatures, and docs.",
          ),
        method: z
          .string()
          .optional()
          .describe("Exact Workbook method name for the describe operation."),
        limit: z
          .number()
          .int()
          .min(1)
          .default(20)
          .describe(`Maximum methods to return (capped at ${WORKBOOK_API_MAX_LIMIT}).`),
        offset: z.number().int().nonnegative().default(0).describe("Zero-based result offset."),
      },
    },
    ({ operation, query, method, limit, offset }) => {
      try {
        if (operation === "describe") {
          if (!method || method.trim().length === 0) {
            throw new Error("method is required for the describe operation");
          }
          return ok(describeWorkbookApi(method.trim()));
        }
        return ok(searchWorkbookApi(query, limit, offset));
      } catch (error) {
        return fail(error);
      }
    },
  );
}
