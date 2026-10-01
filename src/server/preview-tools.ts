import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { jsonText } from "../formulon.js";
import { previewSessionRange } from "../session/preview.js";
import { fail, sheetRefSchema } from "./shared.js";

export function registerPreviewTools(server: McpServer): void {
  server.registerTool(
    "formulon_preview_range",
    {
      title: "Preview workbook range",
      description:
        "Render a bounded workbook range as PNG for visual inspection. The response contains the PNG image first and compact geometry, formatting, pagination, and warning metadata second. Omit range to use a single print area or the used cells plus merges.",
      inputSchema: {
        sessionId: z.string(),
        sheet: sheetRefSchema
          .optional()
          .describe("Zero-based sheet index or name; defaults to the first sheet."),
        range: z.string().optional().describe("Explicit A1 range, such as Sheet1!A1:D20."),
        scale: z.number().min(0.25).max(2).default(1),
        showGridLines: z.boolean().default(false),
        showPageBreaks: z.boolean().default(true),
        recalc: z
          .boolean()
          .default(false)
          .describe("Recalculate before rendering; marks the session dirty."),
        outputPath: z.string().optional().describe("Optional .png or .svg artifact path."),
      },
    },
    ({ sessionId, sheet, range, scale, showGridLines, showPageBreaks, recalc, outputPath }) => {
      try {
        const result = previewSessionRange(sessionId, {
          sheet,
          range,
          scale,
          showGridLines,
          showPageBreaks,
          recalc,
          outputPath,
        });
        return {
          content: [
            { type: "image" as const, data: result.png.toString("base64"), mimeType: "image/png" },
            { type: "text" as const, text: jsonText(result.metadata) },
          ],
        };
      } catch (error) {
        return fail(error);
      }
    },
  );
}
