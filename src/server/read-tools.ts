import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { parseCellRef } from "../a1.js";
import { loadWorkbook, StatusCode, statusError, type Workbook } from "../formulon.js";
import { readWorkbookCell } from "../session/cells.js";
import {
  getSessionCell,
  getSessionCellByA1,
  getSessionMetadata,
  getSessionRange,
  saveSession,
} from "../sessions.js";
import { fail, ok, sheetRefSchema } from "./shared.js";

export function registerReadTools(server: McpServer): void {
  server.registerTool(
    "formulon_get_cell",
    {
      title: "Get cell",
      description:
        "Get one cell from either an open session or a workbook path. Pass exactly one of `sessionId` or `path`; A1 refs and sheet names work for both. Date/currency/percent cells carry a decoded `formatted` string alongside the raw value; formatted numbers also carry `text`, exactly what Excel displays (`textStatus` flags overflow or an invalid format). A formula cell whose value was not recalculated since its inputs changed (manual or automatic-except-tables calculation with `recalc` false) carries `stale: true`; saving the session computes such formulas first, which clears the mark, unless the workbook's `calcOnSave` is off.",
      inputSchema: {
        sessionId: z.string().optional(),
        path: z.string().optional(),
        a1: z.string().optional(),
        sheet: sheetRefSchema.optional().default(0),
        row: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe("Zero-based row (when a1 is omitted)."),
        col: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe("Zero-based column (when a1 is omitted)."),
        recalc: z
          .boolean()
          .default(true)
          .describe(
            "Recalculate before reading. With `path` and false, formulas read the values the file cached, as Excel shows them when it opens the file in Manual calculation.",
          ),
      },
    },
    async ({ sessionId, path, a1, sheet, row, col, recalc }) => {
      let workbook: Workbook | undefined;
      try {
        const hasSession = sessionId !== undefined;
        const hasPath = path !== undefined;
        if (hasSession === hasPath) {
          throw statusError(
            StatusCode.InvalidArgument,
            hasSession
              ? "provide exactly one of sessionId or path, not both"
              : "exactly one of sessionId or path is required",
          );
        }
        if (sessionId !== undefined) {
          if (a1 !== undefined) {
            return ok(getSessionCellByA1(sessionId, a1, sheet, recalc));
          }
          if (row === undefined || col === undefined) {
            throw statusError(
              StatusCode.InvalidArgument,
              "row and col are required when a1 is omitted",
            );
          }
          return ok(getSessionCell(sessionId, sheet, row, col, recalc));
        }
        if (path === undefined) {
          throw statusError(
            StatusCode.InvalidArgument,
            "path is required when sessionId is omitted",
          );
        }
        if (a1 !== undefined) {
          const parsed = parseCellRef(a1);
          workbook = await loadWorkbook(path, { cachedValues: !recalc });
          return ok(
            readWorkbookCell(
              { workbook },
              parsed.sheetName ?? sheet,
              parsed.row,
              parsed.col,
              recalc,
            ),
          );
        }
        if (row === undefined || col === undefined) {
          throw statusError(
            StatusCode.InvalidArgument,
            "row and col are required when a1 is omitted",
          );
        }
        workbook = await loadWorkbook(path, { cachedValues: !recalc });
        return ok(readWorkbookCell({ workbook }, sheet, row, col, recalc));
      } catch (error) {
        return fail(error);
      } finally {
        workbook?.dispose();
      }
    },
  );

  server.registerTool(
    "formulon_get_range",
    {
      title: "Get range",
      description:
        "Get a rectangular A1 range from an open workbook session. Output is sparse (blank cells omitted) and capped at maxCells. Cells with a date/currency/percent number format carry a decoded `formatted` string (for example an ISO date) alongside the raw value; formatted numbers also carry `text`, exactly what Excel displays. A formula cell whose value was not recalculated since its inputs changed (manual or automatic-except-tables calculation with `recalc` false) carries `stale: true`; saving the session computes such formulas first, which clears the mark, unless the workbook's `calcOnSave` is off.",
      inputSchema: {
        sessionId: z.string(),
        range: z.string().describe("A1 range, for example Sheet1!A1:C10 or A1:C10."),
        sheet: sheetRefSchema
          .optional()
          .describe("Sheet used when range has no Sheet!-prefix; defaults to the first sheet."),
        maxCells: z.number().int().positive().max(50_000).default(10_000),
        includeFormulas: z
          .boolean()
          .default(false)
          .describe("Include each cell's formula text (empty for constants)."),
        recalc: z.boolean().default(false).describe("Recalculate before reading."),
      },
    },
    ({ sessionId, range, sheet, maxCells, includeFormulas, recalc }) => {
      try {
        return ok(getSessionRange(sessionId, range, { sheet, maxCells, includeFormulas, recalc }));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_save_session",
    {
      title: "Save session workbook",
      description:
        "Save an open workbook session as .xlsx or .xlsb. XLSB output writes PivotTables, including one built in the session, and the custom views of a workbook opened from .xlsx; anything the binary format cannot hold is left out and counted in `losses.deferredFeatures`, so inspect `losses` before relying on the file. An .xlsb save of a workbook holding something the format cannot store at all, such as a protected range whose security descriptor the format cannot hold, fails with `IoFeatureNotWritable` and writes nothing. Under manual calculation the save first computes formulas left pending, unless the workbook's `calcOnSave` is off, in which case they are written as pending. Otherwise the session stays as it was: where the save writes formatted blanks in Excel's saved form (a row of default-format blanks as a row format, a blank equal to the format it inherits left out), the session still reports those blanks and no row format until the file is reopened; values, display text and effective styles are unaffected. The reopened file can also number its cell and differential formats differently, hold fewer conditional-format rules where Excel merges them on open, and list them in another order; what they resolve to is unchanged.",
      inputSchema: {
        sessionId: z.string(),
        outputPath: z.string().optional(),
      },
    },
    async ({ sessionId, outputPath }) => {
      try {
        return ok(await saveSession(sessionId, outputPath));
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_session_metadata",
    {
      title: "Session metadata",
      description: "Read broad workbook metadata such as registered functions or external links.",
      inputSchema: {
        sessionId: z.string(),
        kind: z.enum(["functions", "externalLinks"]),
      },
    },
    async ({ sessionId, kind }) => {
      try {
        return ok(await getSessionMetadata(sessionId, kind));
      } catch (error) {
        return fail(error);
      }
    },
  );
}
