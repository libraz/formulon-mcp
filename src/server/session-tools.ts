import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { ExcelProfileId } from "../formulon.js";
import {
  analyzeSessionWorkbook,
  closeSession,
  detectSessionRegions,
  findSessionCells,
  inspectSession,
  inspectSessionLayout,
  listSessions,
  openSession,
  recalcSession,
  replaceSessionCells,
} from "../sessions.js";
import { fail, ok, profileSchema, sheetRefSchema } from "./shared.js";

const searchInputSchema = {
  sessionId: z.string(),
  query: z.string(),
  sheet: sheetRefSchema.optional(),
  target: z.enum(["texts", "formulas", "both"]).default("both"),
  matchCase: z.boolean().default(false),
  wholeCell: z.boolean().default(false),
  regex: z.boolean().default(false),
};

export function registerSessionTools(server: McpServer): void {
  server.registerTool(
    "formulon_open_workbook",
    {
      title: "Open workbook session",
      description:
        "Create an in-memory workbook session from an existing .xlsx/.xlsb file or a new default workbook. Anything the reader could not decode is reported in the session's `loadLosses`.",
      inputSchema: {
        path: z
          .string()
          .optional()
          .describe("Optional .xlsx or .xlsb path. Omit to create a new default workbook."),
        sessionId: z
          .string()
          .optional()
          .describe("Optional stable session id; defaults to a UUID."),
        profile: profileSchema
          .optional()
          .describe(
            "Excel profile the session evaluates with (number/date text, TEXT, collation, code page); defaults to FORMULON_MCP_PROFILE, else win-365-en_US. A new workbook is also created as Excel creates it in that locale: body font, Normal style and sheet names, margins.",
          ),
      },
    },
    async ({ path, sessionId, profile }) => {
      try {
        return ok({
          session: await openSession(path, sessionId, profile as ExcelProfileId | undefined),
        });
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_list_sessions",
    {
      title: "List workbook sessions",
      description: "List currently open in-memory workbook sessions.",
      inputSchema: {},
    },
    () => ok({ sessions: listSessions() }),
  );

  server.registerTool(
    "formulon_close_workbook",
    {
      title: "Close workbook session",
      description: "Close an in-memory workbook session and release its native workbook handle.",
      inputSchema: {
        sessionId: z.string(),
      },
    },
    ({ sessionId }) => {
      try {
        return ok({ session: closeSession(sessionId) });
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_inspect_session",
    {
      title: "Inspect session workbook",
      description: "Inspect an open workbook session, optionally including sparse non-empty cells.",
      inputSchema: {
        sessionId: z.string(),
        includeCells: z.boolean().default(false),
        maxCellsPerSheet: z.number().int().nonnegative().max(10_000).default(200),
      },
    },
    ({ sessionId, includeCells, maxCellsPerSheet }) => {
      try {
        return ok(inspectSession(sessionId, includeCells, maxCellsPerSheet));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_recalc_session",
    {
      title: "Recalculate session",
      description: "Recalculate an open workbook session.",
      inputSchema: {
        sessionId: z.string(),
      },
    },
    ({ sessionId }) => {
      try {
        return ok(recalcSession(sessionId));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_find_cells",
    {
      title: "Find cells",
      description: "Search text cell values and/or formula text in an open workbook session.",
      inputSchema: {
        ...searchInputSchema,
        maxResults: z.number().int().positive().max(10_000).default(1_000),
      },
    },
    ({ sessionId, query, sheet, target, matchCase, wholeCell, regex, maxResults }) => {
      try {
        return ok(
          findSessionCells(sessionId, query, {
            sheet,
            target,
            matchCase,
            wholeCell,
            regex,
            maxResults,
          }),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_replace_cells",
    {
      title: "Replace cells",
      description:
        "Replace matching text cell values and/or formula text in an open workbook session.",
      inputSchema: {
        ...searchInputSchema,
        replacement: z.string(),
        maxResults: z.number().int().positive().max(10_000).default(1_000),
        maxReplacements: z.number().int().positive().max(10_000).optional(),
        recalc: z.boolean().default(true),
      },
    },
    ({
      sessionId,
      query,
      sheet,
      target,
      matchCase,
      wholeCell,
      regex,
      replacement,
      maxResults,
      maxReplacements,
      recalc,
    }) => {
      try {
        return ok(
          replaceSessionCells(sessionId, query, {
            sheet,
            target,
            matchCase,
            wholeCell,
            regex,
            maxResults,
            maxReplacements: maxReplacements ?? maxResults,
            replacement,
            recalc,
          }),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_inspect_layout",
    {
      title: "Inspect workbook layout",
      description:
        "Return stable cell, layout, merge, row/column, protection, and optional style data for one or all sheets.",
      inputSchema: {
        sessionId: z.string(),
        sheet: sheetRefSchema.optional(),
        includeCells: z.boolean().default(true),
        includeStyles: z.boolean().default(false),
        maxCells: z.number().int().nonnegative().max(50_000).default(10_000),
      },
    },
    ({ sessionId, sheet, includeCells, includeStyles, maxCells }) => {
      try {
        return ok(
          inspectSessionLayout(sessionId, { sheet, includeCells, includeStyles, maxCells }),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_detect_regions",
    {
      title: "Detect workbook regions",
      description:
        "Detect table-like regions, label-value pairs, and total-like regions with rule-based evidence.",
      inputSchema: {
        sessionId: z.string(),
        sheet: sheetRefSchema.optional(),
        maxCells: z.number().int().nonnegative().max(50_000).default(10_000),
      },
    },
    ({ sessionId, sheet, maxCells }) => {
      try {
        return ok(detectSessionRegions(sessionId, sheet, maxCells));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_analyze_workbook",
    {
      title: "Analyze workbook",
      description:
        "Classify workbook shape such as invoice, list, report, schedule, or form using deterministic features and evidence.",
      inputSchema: {
        sessionId: z.string(),
        includeEvidence: z.boolean().default(true),
        maxCellsPerSheet: z.number().int().nonnegative().max(50_000).default(10_000),
      },
    },
    ({ sessionId, includeEvidence, maxCellsPerSheet }) => {
      try {
        return ok(analyzeSessionWorkbook(sessionId, includeEvidence, maxCellsPerSheet));
      } catch (error) {
        return fail(error);
      }
    },
  );
}
