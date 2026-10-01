import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { auditSessionFormulas } from "../session/audit.js";
import { fail, ok, sheetRefSchema } from "./shared.js";

/** Registers the bounded formula-pattern audit. */
export function registerAuditTools(server: McpServer): void {
  server.registerTool(
    "formulon_audit_formulas",
    {
      title: "Audit repeated formula patterns",
      description:
        "Scan a workbook range for repeated formula patterns and report constants, blanks, outliers, and formula errors for review. The default range is the complete stored used rectangle, with no silent prefix truncation. A fully blank row or column is a hard boundary, so include the populated input columns when auditing a table body. The audit never repairs formulas and does not establish formula correctness. Recalculation is opt-in and marks the session dirty before the engine call.",
      inputSchema: {
        sessionId: z.string(),
        sheet: sheetRefSchema
          .optional()
          .describe("Zero-based sheet index or sheet name; defaults to the first sheet."),
        range: z
          .string()
          .optional()
          .describe("Optional A1 range; a qualified range chooses its sheet."),
        direction: z.enum(["vertical", "horizontal", "both"]).default("vertical"),
        window: z.number().int().min(2).max(50).default(5),
        minPeers: z.number().int().min(3).max(20).default(3),
        maxCells: z.number().int().min(1).max(50_000).default(50_000),
        maxFindings: z.number().int().min(1).max(500).default(100),
        recalc: z.boolean().default(false),
      },
    },
    ({ sessionId, sheet, range, direction, window, minPeers, maxCells, maxFindings, recalc }) => {
      try {
        return ok(
          auditSessionFormulas(sessionId, {
            sheet,
            range,
            direction,
            window,
            minPeers,
            maxCells,
            maxFindings,
            recalc,
          }),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );
}
