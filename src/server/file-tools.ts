import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  applyMutation,
  assertStatus,
  type CellMutation,
  createOrLoadWorkbook,
  loadWorkbook,
  saveWorkbook,
  type Workbook,
  workbookSummary,
} from "../formulon.js";
import { fail, ok } from "./shared.js";

const concreteMutationSchema = z.discriminatedUnion("type", [
  z.object({
    type: z.literal("number"),
    sheet: z.number().int().nonnegative(),
    row: z.number().int().nonnegative(),
    col: z.number().int().nonnegative(),
    value: z.number(),
  }),
  z.object({
    type: z.literal("bool"),
    sheet: z.number().int().nonnegative(),
    row: z.number().int().nonnegative(),
    col: z.number().int().nonnegative(),
    value: z.boolean(),
  }),
  z.object({
    type: z.literal("text"),
    sheet: z.number().int().nonnegative(),
    row: z.number().int().nonnegative(),
    col: z.number().int().nonnegative(),
    value: z.string(),
  }),
  z.object({
    type: z.literal("blank"),
    sheet: z.number().int().nonnegative(),
    row: z.number().int().nonnegative(),
    col: z.number().int().nonnegative(),
  }),
  z.object({
    type: z.literal("formula"),
    sheet: z.number().int().nonnegative(),
    row: z.number().int().nonnegative(),
    col: z.number().int().nonnegative(),
    formula: z.string(),
  }),
]);

export function registerFileTools(server: McpServer): void {
  server.registerTool(
    "formulon_inspect_workbook",
    {
      title: "Inspect workbook path",
      description: "Load an xlsx or xlsb workbook path and return a one-shot summary.",
      inputSchema: {
        path: z.string().describe("Path to an .xlsx or .xlsb workbook."),
        recalc: z.boolean().default(false),
        includeCells: z.boolean().default(false),
        maxCellsPerSheet: z.number().int().nonnegative().max(10_000).default(200),
      },
    },
    async ({ path, recalc, includeCells, maxCellsPerSheet }) => {
      let workbook: Workbook | undefined;
      try {
        workbook = await loadWorkbook(path);
        if (recalc) {
          assertStatus(workbook.recalc(), "recalc workbook");
        }
        return ok(workbookSummary(workbook, includeCells, maxCellsPerSheet));
      } catch (error) {
        return fail(error);
      } finally {
        workbook?.delete();
      }
    },
  );

  server.registerTool(
    "formulon_update_workbook",
    {
      title: "One-shot workbook update",
      description:
        "Create or load a workbook, apply zero-based cell mutations, recalculate, and save as .xlsx or .xlsb. Inspect returned losses; use XLSX to preserve authored tables, pivots, print settings, and comments.",
      inputSchema: {
        inputPath: z.string().optional(),
        outputPath: z.string(),
        recalc: z.boolean().default(true),
        mutations: z.array(concreteMutationSchema).min(1).max(10_000),
      },
    },
    async ({ inputPath, outputPath, recalc, mutations }) => {
      let workbook: Workbook | undefined;
      try {
        workbook = await createOrLoadWorkbook(inputPath);
        for (const [idx, mutation] of mutations.entries()) {
          assertStatus(applyMutation(workbook, mutation as CellMutation), `apply mutation ${idx}`);
        }
        if (recalc) {
          assertStatus(workbook.recalc(), "recalc workbook");
        }
        const saved = await saveWorkbook(workbook, outputPath);
        return ok({
          outputPath,
          bytes: saved.bytes,
          format: saved.format,
          losses: saved.losses,
          summary: workbookSummary(workbook, false, 0),
        });
      } catch (error) {
        return fail(error);
      } finally {
        workbook?.delete();
      }
    },
  );
}
