#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  applyMutation,
  assertStatus,
  type CellMutation,
  createOrLoadWorkbook,
  formulonModule,
  jsonText,
  loadWorkbook,
  normalizeFormula,
  saveWorkbook,
  valueToJson,
  type Workbook,
  workbookSummary,
} from "./formulon.js";
import { runInit, runUninstall } from "./init.js";
import { PRINT_PRESET_NAMES } from "./print.js";
import {
  analyzeSessionWorkbook,
  applySessionMutations,
  applySheetOperation,
  buildSessionDocument,
  callWorkbookMethod,
  closeSession,
  detectSessionRegions,
  editSessionStructure,
  findSessionCells,
  getSessionCell,
  getSessionCellByA1,
  getSessionMetadata,
  getSessionRange,
  inspectSession,
  inspectSessionLayout,
  listSessions,
  openSession,
  type RangeWriteCell,
  recalcSession,
  replaceSessionCells,
  resolveSessionSheet,
  saveSession,
  sessionDimension,
  sessionPrintSettings,
  setSessionDefinedName,
  setSessionRange,
  setSessionSheetView,
  styleSessionRange,
  traceSession,
} from "./sessions.js";
import { STYLE_VOCABULARY } from "./styles.js";

/** Single source of truth for the server version: package.json at runtime. */
const PACKAGE_VERSION = ((): string => {
  try {
    const pkgUrl = new URL("../package.json", import.meta.url);
    const pkg = JSON.parse(readFileSync(pkgUrl, "utf8")) as { version?: string };
    return pkg.version ?? "0.0.0";
  } catch {
    return "0.0.0";
  }
})();

const server = new McpServer({
  name: "formulon-mcp",
  version: PACKAGE_VERSION,
});

function ok(value: unknown) {
  return {
    content: [{ type: "text" as const, text: jsonText(value) }],
  };
}

function fail(error: unknown) {
  const message = error instanceof Error ? error.message : String(error);
  return {
    isError: true,
    content: [{ type: "text" as const, text: message }],
  };
}

const sheetRefSchema = z.union([z.number().int().nonnegative(), z.string()]);

const cellMutationSchema = z.discriminatedUnion("type", [
  z.object({
    type: z.literal("number"),
    sheet: sheetRefSchema.optional(),
    a1: z.string().optional(),
    row: z.number().int().nonnegative().optional(),
    col: z.number().int().nonnegative().optional(),
    value: z.number(),
  }),
  z.object({
    type: z.literal("bool"),
    sheet: sheetRefSchema.optional(),
    a1: z.string().optional(),
    row: z.number().int().nonnegative().optional(),
    col: z.number().int().nonnegative().optional(),
    value: z.boolean(),
  }),
  z.object({
    type: z.literal("text"),
    sheet: sheetRefSchema.optional(),
    a1: z.string().optional(),
    row: z.number().int().nonnegative().optional(),
    col: z.number().int().nonnegative().optional(),
    value: z.string(),
  }),
  z.object({
    type: z.literal("blank"),
    sheet: sheetRefSchema.optional(),
    a1: z.string().optional(),
    row: z.number().int().nonnegative().optional(),
    col: z.number().int().nonnegative().optional(),
  }),
  z.object({
    type: z.literal("formula"),
    sheet: sheetRefSchema.optional(),
    a1: z.string().optional(),
    row: z.number().int().nonnegative().optional(),
    col: z.number().int().nonnegative().optional(),
    formula: z.string(),
  }),
]);

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

const jsonArgsSchema = z.array(z.unknown()).default([]);

const rangeSchema = z.object({
  firstRow: z.number().int().nonnegative(),
  firstCol: z.number().int().nonnegative(),
  lastRow: z.number().int().nonnegative(),
  lastCol: z.number().int().nonnegative(),
});

const searchInputSchema = {
  sessionId: z.string(),
  query: z.string(),
  sheet: sheetRefSchema.optional(),
  target: z.enum(["texts", "formulas", "both"]).default("both"),
  matchCase: z.boolean().default(false),
  wholeCell: z.boolean().default(false),
  regex: z.boolean().default(false),
};

const sheetInputSchema = {
  sessionId: z.string(),
  sheet: sheetRefSchema
    .default(0)
    .describe("Zero-based sheet index or sheet name; defaults to the first sheet."),
};

function methodOk(sessionId: string, method: string, args: unknown[]) {
  return ok(callWorkbookMethod(sessionId, method, args));
}

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
      return ok({ version: module.versionString(), serverVersion: PACKAGE_VERSION });
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
      sheet: z.number().int().nonnegative().default(0),
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
        return methodOk(sessionId, "evaluateFormulaText", [
          sheet,
          row,
          col,
          normalizeFormula(formula),
        ]);
      }
      const module = await formulonModule();
      const result = module.evalFormula(normalizeFormula(formula));
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
      sessionId: z.string().optional().describe("Optional stable session id; defaults to a UUID."),
    },
  },
  async ({ path, sessionId }) => {
    try {
      return ok({ session: await openSession(path, sessionId) });
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
      return ok(inspectSessionLayout(sessionId, { sheet, includeCells, includeStyles, maxCells }));
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

server.registerTool(
  "formulon_set_cells",
  {
    title: "Set session cells",
    description:
      "Apply cell mutations to an open workbook session. Cells may use A1 refs or zero-based row/col.",
    inputSchema: {
      sessionId: z.string(),
      recalc: z.boolean().default(true),
      mutations: z.array(cellMutationSchema).min(1).max(10_000),
    },
  },
  ({ sessionId, mutations, recalc }) => {
    try {
      return ok(applySessionMutations(sessionId, mutations, recalc));
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_set_range",
  {
    title: "Set range block",
    description:
      'Write a 2D block of values starting at an anchor cell. Each element\'s JSON type picks the cell type: number, boolean, or string. Use {"f":"=SUM(A1:A3)"} for a formula and null to skip a cell. Far more compact than set_cells for tables.',
    inputSchema: {
      sessionId: z.string(),
      start: z.string().describe("Anchor (top-left) A1 cell, for example Sheet1!B2 or B2."),
      sheet: sheetRefSchema
        .optional()
        .describe("Sheet used when start has no Sheet!-prefix; defaults to the first sheet."),
      values: z
        .array(
          z.array(
            z.union([z.number(), z.boolean(), z.string(), z.null(), z.object({ f: z.string() })]),
          ),
        )
        .min(1)
        .describe("Row-major 2D array of cell values."),
      recalc: z.boolean().default(true),
    },
  },
  ({ sessionId, start, sheet, values, recalc }) => {
    try {
      return ok(setSessionRange(sessionId, start, values as RangeWriteCell[][], sheet, recalc));
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_sheet_operation",
  {
    title: "Sheet operation",
    description: "Add, remove, rename, or move a sheet in an open workbook session.",
    inputSchema: {
      sessionId: z.string(),
      operation: z.enum(["add", "remove", "rename", "move"]),
      name: z.string().optional().describe("Sheet name for add."),
      index: z.number().int().nonnegative().optional().describe("Sheet index for remove/rename."),
      newName: z.string().optional().describe("New sheet name for rename."),
      fromIndex: z.number().int().nonnegative().optional().describe("Source index for move."),
      toIndex: z.number().int().nonnegative().optional().describe("Destination index for move."),
    },
  },
  ({ sessionId, operation, name, index, newName, fromIndex, toIndex }) => {
    try {
      return ok(
        applySheetOperation(sessionId, operation, { name, index, newName, fromIndex, toIndex }),
      );
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_set_defined_name",
  {
    title: "Set defined name",
    description:
      "Add, replace, or remove a defined name. Without `sheet` the name is workbook-scoped; with it the name is local to that sheet. Print settings (_xlnm.Print_Area, _xlnm.Print_Titles) must be sheet-scoped — Excel ignores a workbook-scoped one without reporting an error. Empty formula removes it.",
    inputSchema: {
      sessionId: z.string(),
      name: z.string(),
      formula: z.string().describe("Formula with or without '='; pass empty string to remove."),
      sheet: sheetRefSchema
        .optional()
        .describe("Scope the name to this sheet. Omit for workbook scope."),
    },
  },
  ({ sessionId, name, formula, sheet }) => {
    try {
      return ok(setSessionDefinedName(sessionId, name, formula, sheet));
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_edit_structure",
  {
    title: "Edit rows or columns",
    description: "Insert or delete rows/columns in an open workbook session.",
    inputSchema: {
      sessionId: z.string(),
      operation: z.enum(["insertRows", "deleteRows", "insertCols", "deleteCols"]),
      sheet: sheetRefSchema.optional().default(0),
      start: z.number().int().nonnegative(),
      count: z.number().int().positive(),
    },
  },
  ({ sessionId, operation, sheet, start, count }) => {
    try {
      return ok(editSessionStructure(sessionId, operation, sheet, start, count));
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_set_sheet_view",
  {
    title: "Set sheet view",
    description: "Set sheet zoom, frozen panes, or tab visibility.",
    inputSchema: {
      sessionId: z.string(),
      sheet: sheetRefSchema.optional().default(0),
      zoom: z.number().int().min(10).max(400).optional(),
      freezeRows: z.number().int().nonnegative().optional(),
      freezeCols: z.number().int().nonnegative().optional(),
      hidden: z.boolean().optional().describe("Two-state tab visibility; cannot state veryHidden."),
      visibility: z
        .enum(["visible", "hidden", "veryHidden"])
        .optional()
        .describe(
          "Three-state tab visibility. Excel leaves a veryHidden sheet out of its Unhide dialog.",
        ),
    },
  },
  ({ sessionId, sheet, zoom, freezeRows, freezeCols, hidden, visibility }) => {
    try {
      return ok(
        setSessionSheetView(sessionId, sheet, {
          zoom,
          freezeRows,
          freezeCols,
          hidden,
          visibility,
        }),
      );
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_get_cell",
  {
    title: "Get cell",
    description:
      "Get one cell from either an open session or a workbook path. A1 refs are supported for sessions. Date/currency/percent cells carry a decoded `formatted` string alongside the raw value.",
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
      recalc: z.boolean().default(true),
    },
  },
  async ({ sessionId, path, a1, sheet, row, col, recalc }) => {
    let workbook: Workbook | undefined;
    try {
      if (sessionId) {
        if (a1) {
          return ok(getSessionCellByA1(sessionId, a1, sheet, recalc));
        }
        if (row === undefined || col === undefined) {
          throw new Error("row and col are required when a1 is omitted");
        }
        return ok(getSessionCell(sessionId, sheet, row, col, recalc));
      }
      if (!path) {
        throw new Error("path is required when sessionId is omitted");
      }
      if (a1) {
        throw new Error("a1 path reads are session-only; open the workbook first");
      }
      if (row === undefined || col === undefined || typeof sheet !== "number") {
        throw new Error("path reads require numeric sheet, row, and col");
      }
      workbook = await loadWorkbook(path);
      if (recalc) {
        assertStatus(workbook.recalc(), "recalc workbook");
      }
      const result = workbook.getValue(sheet, row, col);
      return ok({
        status: result.status,
        value: valueToJson(result.value),
      });
    } catch (error) {
      return fail(error);
    } finally {
      workbook?.delete();
    }
  },
);

server.registerTool(
  "formulon_get_range",
  {
    title: "Get range",
    description:
      "Get a rectangular A1 range from an open workbook session. Output is sparse (blank cells omitted), clipped to the sheet's used range, and capped at maxCells. Cells with a date/currency/percent number format carry a decoded `formatted` string (for example an ISO date) alongside the raw value.",
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
      "Save an open workbook session. The container follows the output extension (.xlsb writes XLSB, anything else XLSX), and anything the writer had to drop or downgrade is reported in `losses`.",
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
  ({ sessionId, kind }) => {
    try {
      return ok(getSessionMetadata(sessionId, kind));
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_merge_operation",
  {
    title: "Merge operation",
    description: "List, add, remove, remove by index, or clear merged ranges on a sheet.",
    inputSchema: {
      ...sheetInputSchema,
      operation: z.enum(["list", "add", "remove", "removeAt", "clear"]),
      range: rangeSchema.optional(),
      index: z.number().int().nonnegative().optional(),
    },
  },
  ({ sessionId, sheet, operation, range, index }) => {
    try {
      const sheetIndex = resolveSessionSheet(sessionId, sheet);
      if (operation === "list") {
        return methodOk(sessionId, "getMerges", [sheetIndex]);
      }
      if (operation === "clear") {
        return methodOk(sessionId, "clearMerges", [sheetIndex]);
      }
      if (operation === "removeAt") {
        if (index === undefined) {
          throw new Error("index is required for removeAt");
        }
        return methodOk(sessionId, "removeMergeAt", [sheetIndex, index]);
      }
      if (!range) {
        throw new Error("range is required for add/remove");
      }
      return methodOk(sessionId, operation === "add" ? "addMerge" : "removeMerge", [
        sheetIndex,
        range,
      ]);
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_comment_operation",
  {
    title: "Comment operation",
    description: "List, get, set, or remove a cell comment.",
    inputSchema: {
      ...sheetInputSchema,
      operation: z.enum(["list", "get", "set", "remove"]),
      row: z.number().int().nonnegative().optional(),
      col: z.number().int().nonnegative().optional(),
      author: z.string().optional(),
      text: z.string().optional(),
    },
  },
  ({ sessionId, sheet, operation, row, col, author, text }) => {
    try {
      const sheetIndex = resolveSessionSheet(sessionId, sheet);
      if (operation === "list") {
        return methodOk(sessionId, "getComments", [sheetIndex]);
      }
      if (row === undefined || col === undefined) {
        throw new Error("row and col are required");
      }
      if (operation === "get") {
        return methodOk(sessionId, "getComment", [sheetIndex, row, col]);
      }
      return methodOk(sessionId, "setComment", [
        sheetIndex,
        row,
        col,
        operation === "remove" ? "" : (author ?? ""),
        operation === "remove" ? "" : (text ?? ""),
      ]);
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_hyperlink_operation",
  {
    title: "Hyperlink operation",
    description:
      "List, add, remove, remove by index, or clear hyperlinks on a sheet. An added link covers one cell, or the rectangle through lastRow/lastCol.",
    inputSchema: {
      ...sheetInputSchema,
      operation: z.enum(["list", "add", "remove", "removeAt", "clear"]),
      row: z.number().int().nonnegative().optional(),
      col: z.number().int().nonnegative().optional(),
      lastRow: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe("Inclusive last row of a multi-cell hyperlink; defaults to row."),
      lastCol: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe("Inclusive last column of a multi-cell hyperlink; defaults to col."),
      index: z.number().int().nonnegative().optional(),
      target: z.string().optional().describe("External URL; leave empty for an in-workbook link."),
      display: z.string().optional(),
      tooltip: z.string().optional(),
      location: z
        .string()
        .optional()
        .describe("In-workbook destination such as Sheet2!A1, used when target is empty."),
    },
  },
  ({
    sessionId,
    sheet,
    operation,
    row,
    col,
    lastRow,
    lastCol,
    index,
    target,
    display,
    tooltip,
    location,
  }) => {
    try {
      const sheetIndex = resolveSessionSheet(sessionId, sheet);
      if (operation === "list") {
        return methodOk(sessionId, "getHyperlinks", [sheetIndex]);
      }
      if (operation === "clear") {
        return methodOk(sessionId, "clearHyperlinks", [sheetIndex]);
      }
      if (operation === "removeAt") {
        if (index === undefined) {
          throw new Error("index is required for removeAt");
        }
        return methodOk(sessionId, "removeHyperlinkAt", [sheetIndex, index]);
      }
      if (row === undefined || col === undefined) {
        throw new Error("row and col are required");
      }
      if (operation === "remove") {
        return methodOk(sessionId, "removeHyperlink", [sheetIndex, row, col]);
      }
      if (lastRow !== undefined || lastCol !== undefined) {
        return methodOk(sessionId, "addHyperlinkRange", [
          sheetIndex,
          row,
          col,
          lastRow ?? row,
          lastCol ?? col,
          target ?? "",
          display ?? "",
          tooltip ?? "",
          location ?? "",
        ]);
      }
      return methodOk(sessionId, "addHyperlink", [
        sheetIndex,
        row,
        col,
        target ?? "",
        display ?? "",
        tooltip ?? "",
        location ?? "",
      ]);
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_validation_operation",
  {
    title: "Validation operation",
    description:
      "List, add, remove by index, or clear data validations on a sheet. An omitted boolean field defaults to false, so a rule that should accept empty cells has to spell allowBlank: true; showDropDown is the only field that defaults to true.",
    inputSchema: {
      ...sheetInputSchema,
      operation: z.enum(["list", "add", "removeAt", "clear"]),
      index: z.number().int().nonnegative().optional(),
      validation: z.record(z.string(), z.unknown()).optional(),
    },
  },
  ({ sessionId, sheet, operation, index, validation }) => {
    try {
      const sheetIndex = resolveSessionSheet(sessionId, sheet);
      if (operation === "list") {
        return methodOk(sessionId, "getValidations", [sheetIndex]);
      }
      if (operation === "clear") {
        return methodOk(sessionId, "clearValidations", [sheetIndex]);
      }
      if (operation === "removeAt") {
        if (index === undefined) {
          throw new Error("index is required for removeAt");
        }
        return methodOk(sessionId, "removeValidationAt", [sheetIndex, index]);
      }
      if (!validation) {
        throw new Error("validation is required for add");
      }
      return methodOk(sessionId, "addValidation", [sheetIndex, validation]);
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_conditional_format_operation",
  {
    title: "Conditional format operation",
    description: "List, add, remove by index, clear, or evaluate conditional format rules.",
    inputSchema: {
      ...sheetInputSchema,
      operation: z.enum(["list", "add", "removeAt", "clear", "evaluate"]),
      index: z.number().int().nonnegative().optional(),
      rule: z.record(z.string(), z.unknown()).optional(),
      firstRow: z.number().int().nonnegative().optional(),
      firstCol: z.number().int().nonnegative().optional(),
      lastRow: z.number().int().nonnegative().optional(),
      lastCol: z.number().int().nonnegative().optional(),
      todaySerial: z.number().optional(),
    },
  },
  ({
    sessionId,
    sheet,
    operation,
    index,
    rule,
    firstRow,
    firstCol,
    lastRow,
    lastCol,
    todaySerial,
  }) => {
    try {
      const sheetIndex = resolveSessionSheet(sessionId, sheet);
      if (operation === "list") {
        return methodOk(sessionId, "getConditionalFormats", [sheetIndex]);
      }
      if (operation === "clear") {
        return methodOk(sessionId, "clearConditionalFormats", [sheetIndex]);
      }
      if (operation === "removeAt") {
        if (index === undefined) {
          throw new Error("index is required for removeAt");
        }
        return methodOk(sessionId, "removeConditionalFormatAt", [sheetIndex, index]);
      }
      if (operation === "evaluate") {
        if (
          firstRow === undefined ||
          firstCol === undefined ||
          lastRow === undefined ||
          lastCol === undefined
        ) {
          throw new Error("firstRow, firstCol, lastRow, and lastCol are required for evaluate");
        }
        return methodOk(sessionId, "evaluateCfRange", [
          sheetIndex,
          firstRow,
          firstCol,
          lastRow,
          lastCol,
          todaySerial ?? Number.NaN,
        ]);
      }
      if (!rule) {
        throw new Error("rule is required for add");
      }
      return methodOk(sessionId, "addConditionalFormat", [sheetIndex, rule]);
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_trace",
  {
    title: "Trace dependencies",
    description:
      "Read precedents, dependents, or spill info for a cell. Results carry A1 references (with sheet names), not raw indices.",
    inputSchema: {
      ...sheetInputSchema,
      operation: z.enum(["precedents", "dependents", "spillInfo"]),
      row: z.number().int().nonnegative().describe("Zero-based row of the cell to trace."),
      col: z.number().int().nonnegative().describe("Zero-based column of the cell to trace."),
      depth: z.number().int().positive().max(32).default(1),
    },
  },
  ({ sessionId, sheet, operation, row, col, depth }) => {
    try {
      return ok(traceSession(sessionId, operation, sheet, row, col, depth));
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_dimension_operation",
  {
    title: "Column/row dimension operation",
    description:
      "List column-width/row-height overrides, or set width/height, hidden, or outline level. Columns act on an inclusive [first, last] span; rows act on a single row index.",
    inputSchema: {
      ...sheetInputSchema,
      axis: z.enum(["column", "row"]),
      operation: z.enum(["list", "size", "hidden", "outline"]),
      first: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe("Zero-based first column index (column axis)."),
      last: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe("Zero-based last column index; defaults to first (column axis)."),
      row: z.number().int().nonnegative().optional().describe("Zero-based row index (row axis)."),
      size: z.number().optional().describe("Column width or row height for the 'size' operation."),
      hidden: z.boolean().optional().describe("Hidden flag for the 'hidden' operation."),
      level: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe("Outline level for the 'outline' operation."),
    },
  },
  ({ sessionId, sheet, axis, operation, first, last, row, size, hidden, level }) => {
    try {
      return ok(
        sessionDimension(sessionId, axis, operation, {
          sheet,
          first,
          last,
          row,
          size,
          hidden,
          level,
        }),
      );
    } catch (error) {
      return fail(error);
    }
  },
);

const borderSideSchema = z
  .union([
    z.string(),
    z.object({
      style: z.string().optional(),
      color: z.string().optional(),
    }),
  ])
  .describe(
    `Border style name (${STYLE_VOCABULARY.borderStyle.join(", ")}), or {style, color} with a #RRGGBB color.`,
  );

const styleSchema = z.object({
  font: z
    .object({
      name: z.string().optional(),
      size: z.number().positive().optional(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      strike: z.boolean().optional(),
      underline: z.enum(STYLE_VOCABULARY.underline).optional(),
      vertAlign: z.enum(STYLE_VOCABULARY.vertAlign).optional(),
      color: z.string().optional().describe("#RRGGBB or #AARRGGBB."),
    })
    .optional(),
  fill: z
    .object({
      color: z.string().optional().describe("Fill color as #RRGGBB; implies a solid pattern."),
      bgColor: z.string().optional(),
      pattern: z.enum(STYLE_VOCABULARY.fillPattern).optional(),
    })
    .optional(),
  border: z
    .object({
      all: borderSideSchema.optional().describe("Rules every cell in the range on all four sides."),
      outline: borderSideSchema.optional().describe("Draws a box around the range only."),
      left: borderSideSchema.optional(),
      right: borderSideSchema.optional(),
      top: borderSideSchema.optional(),
      bottom: borderSideSchema.optional(),
    })
    .optional(),
  numberFormat: z
    .string()
    .optional()
    .describe('Excel format code such as "#,##0" or "yyyy/mm/dd". Empty string means General.'),
  align: z
    .object({
      horizontal: z.enum(STYLE_VOCABULARY.horizontalAlign).optional(),
      vertical: z.enum(STYLE_VOCABULARY.verticalAlign).optional(),
      wrapText: z.boolean().optional(),
      indent: z.number().int().min(0).max(255).optional(),
      textRotation: z.number().int().min(0).max(255).optional(),
      shrinkToFit: z.boolean().optional(),
    })
    .optional(),
});

const cellLiteralSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);

const formatFieldSchema = z
  .string()
  .describe(
    "Excel format code, or one of the aliases date, datetime, time, number, decimal, percent. An ISO date string under a date format is stored as a real date, not text.",
  );

const documentBlockSchema = z.discriminatedUnion("type", [
  z.object({
    type: z.literal("title"),
    text: z.string(),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    align: z.enum(["left", "center", "right"]).optional(),
    size: z.number().positive().optional(),
    bold: z.boolean().optional(),
  }),
  z.object({
    type: z.literal("text"),
    text: z.string(),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    align: z.enum(["left", "center", "right"]).optional(),
    size: z.number().positive().optional(),
    bold: z.boolean().optional(),
    span: z.number().int().positive().optional(),
    wrap: z.boolean().optional(),
    rows: z.number().int().positive().optional(),
  }),
  z.object({
    type: z.literal("fields"),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    align: z.enum(["left", "right"]).optional(),
    labelSpan: z.number().int().positive().optional(),
    valueSpan: z.number().int().positive().optional(),
    rule: z.boolean().optional().describe("Rule each value cell underneath, like a paper form."),
    items: z
      .array(
        z.object({
          label: z.string(),
          value: cellLiteralSchema.optional(),
          formula: z.string().optional(),
          format: formatFieldSchema.optional(),
          name: z.string().optional(),
        }),
      )
      .min(1),
  }),
  z.object({
    type: z.literal("table"),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    bandColor: z.string().optional().describe("Fill applied to every other body row."),
    rowCount: z
      .number()
      .int()
      .nonnegative()
      .optional()
      .describe("Blank ruled rows when `rows` is omitted — a form to fill in."),
    columns: z
      .array(
        z.object({
          header: z.string(),
          key: z.string().optional().describe("Name used in formulas; defaults to the header."),
          formula: z
            .string()
            .optional()
            .describe("Per-row formula; {key} binds to that column's cell in the same row."),
          width: z.number().positive().optional(),
          format: formatFieldSchema.optional(),
          align: z.enum(["left", "center", "right"]).optional(),
        }),
      )
      .min(1),
    rows: z
      .array(z.union([z.array(cellLiteralSchema), z.record(z.string(), cellLiteralSchema)]))
      .optional()
      .describe("Row objects keyed by column key/header, or positional arrays."),
  }),
  z.object({
    type: z.literal("summary"),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    align: z.enum(["left", "right"]).optional(),
    labelSpan: z.number().int().positive().optional(),
    valueSpan: z.number().int().positive().optional(),
    border: z.boolean().optional(),
    items: z
      .array(
        z.object({
          label: z.string(),
          value: cellLiteralSchema.optional(),
          formula: z.string().optional(),
          format: formatFieldSchema.optional(),
          emphasis: z.boolean().optional(),
          name: z.string().optional(),
        }),
      )
      .min(1),
  }),
  z.object({
    type: z.literal("spacer"),
    name: z.string().optional(),
    sameRow: z.boolean().optional(),
    rows: z.number().int().nonnegative().optional(),
  }),
]);

server.registerTool(
  "formulon_build_document",
  {
    title: "Build a document from blocks",
    description:
      'Lay out a document as a vertical stack of blocks — title, text, fields, table, summary, spacer — and write it in one call. Positions, ruling, number formats, column widths, merges and the print area are all resolved from the layout, so no row arithmetic is needed. Blocks reference each other by name: a table column registers as {table.<header>} over its body range, and a fields or summary item registers under its label, so "=SUM({table.Amount})" and "={Subtotal}+{Tax}" bind to the right cells once the layout is known. The response maps every block and name to A1 so the result can be refined with set_cells / style_range / print_settings. This tool carries no document semantics: labels, tax rules and totals are the caller\'s formulas.',
    inputSchema: {
      sessionId: z.string(),
      sheet: sheetRefSchema
        .optional()
        .describe("Sheet used when `start` has no Sheet!-prefix; defaults to the first sheet."),
      start: z
        .string()
        .default("B2")
        .describe("Top-left anchor, for example Sheet1!B2. Defaults to B2, leaving a margin."),
      width: z
        .number()
        .int()
        .positive()
        .optional()
        .describe("Document width in columns; defaults to the widest table's column count."),
      blocks: z.array(documentBlockSchema).min(1),
      theme: z
        .object({
          font: z.string().optional(),
          size: z.number().positive().optional(),
          accent: z
            .string()
            .optional()
            .describe("Header-band fill as #RRGGBB. Omit for a plain ruled header."),
          headerText: z.string().optional(),
          border: z.string().optional().describe("Grid border style; defaults to thin."),
          outline: z.string().optional().describe("Outer border style; defaults to medium."),
        })
        .optional(),
      print: z
        .enum(PRINT_PRESET_NAMES)
        .optional()
        .describe("Named page setup; also sets the print area to the document's extent."),
      repeatTableHeader: z
        .boolean()
        .optional()
        .describe("Repeat the first table's header row on every printed page. Defaults to true."),
    },
  },
  ({ sessionId, sheet, ...spec }) => {
    try {
      return ok(
        buildSessionDocument(sessionId, spec as Parameters<typeof buildSessionDocument>[1], sheet),
      );
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_style_range",
  {
    title: "Style a range",
    description:
      "Apply fonts, fills, borders, number formats, and alignment across an A1 range, materializing blank cells so an empty ruled box renders. Each property is a delta: cells keep whatever the style does not state, so ruling a table does not undo the number format already on its amount column. Style in passes — the grid over the whole table, then the header row, then a total cell.",
    inputSchema: {
      sessionId: z.string(),
      range: z.string().describe("A1 range or single cell, for example Sheet1!B2:F20 or B2."),
      sheet: sheetRefSchema
        .optional()
        .describe("Sheet used when range has no Sheet!-prefix; defaults to the first sheet."),
      style: styleSchema,
      baseOn: z
        .enum(["existing", "default"])
        .default("existing")
        .describe(
          "Start from the style already on the range's top-left cell, or from the workbook default.",
        ),
    },
  },
  ({ sessionId, range, sheet, style, baseOn }) => {
    try {
      return ok(styleSessionRange(sessionId, range, style, sheet, baseOn));
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_print_settings",
  {
    title: "Print settings",
    description:
      "Read or set a sheet's print settings: page setup, margins, print options, header/footer, print area, print titles, and manual page breaks. Omit every setting to read. Reads report the resulting pageCount, so a page layout can be checked without saving the file.",
    inputSchema: {
      ...sheetInputSchema,
      pageSetup: z
        .object({
          orientation: z.enum(["default", "portrait", "landscape"]).optional(),
          paperSize: z
            .number()
            .int()
            .nonnegative()
            .optional()
            .describe("OOXML paper-size code: 9 = A4, 8 = A3, 11 = A5, 1 = Letter."),
          scale: z.number().int().min(10).max(400).optional(),
          fitToWidth: z.number().int().nonnegative().optional(),
          fitToHeight: z.number().int().nonnegative().optional(),
          fitToPage: z
            .boolean()
            .optional()
            .describe(
              "Selects fit-to-page mode; fitToWidth/fitToHeight state the target, so all three are needed to fit onto one page.",
            ),
        })
        .optional(),
      margins: z
        .object({
          left: z.number().nonnegative().optional(),
          right: z.number().nonnegative().optional(),
          top: z.number().nonnegative().optional(),
          bottom: z.number().nonnegative().optional(),
          header: z.number().nonnegative().optional(),
          footer: z.number().nonnegative().optional(),
        })
        .optional()
        .describe("Page margins in inches."),
      printOptions: z
        .object({
          gridLines: z.boolean().optional(),
          headings: z.boolean().optional(),
          horizontalCentered: z.boolean().optional(),
          verticalCentered: z.boolean().optional(),
        })
        .optional(),
      headerFooter: z
        .object({
          oddHeader: z.string().optional(),
          oddFooter: z.string().optional(),
          evenHeader: z.string().optional(),
          evenFooter: z.string().optional(),
          firstHeader: z.string().optional(),
          firstFooter: z.string().optional(),
          differentOddEven: z.boolean().optional(),
          differentFirst: z.boolean().optional(),
          scaleWithDoc: z.boolean().optional(),
          alignWithMargins: z.boolean().optional(),
        })
        .optional()
        .describe(
          'Section text uses Excel\'s own codes — &L/&C/&R pick the section, &P the page number, &N the page count, &D the date, and && a literal ampersand (for example \'&C&"MS Gothic"Invoice &P/&N\'). Each section is tri-state: omit to leave it, "" to clear it.',
        ),
      printArea: z
        .string()
        .optional()
        .describe(
          'Comma-separated A1 ranges, for example "A1:F40" or "A1:B10,D5:E20". Empty string removes the print area.',
        ),
      printTitles: z
        .object({
          repeatRows: z.string().optional().describe('Whole-row span such as "1:2".'),
          repeatCols: z.string().optional().describe('Whole-column span such as "A:A".'),
        })
        .optional()
        .describe("Rows and columns repeated on every printed page. Both empty removes them."),
      rowBreaks: z
        .array(z.number().int().nonnegative())
        .optional()
        .describe("Zero-based rows each manual break precedes. Replaces the sheet's breaks."),
      colBreaks: z
        .array(z.number().int().nonnegative())
        .optional()
        .describe("Zero-based columns each manual break precedes. Replaces the sheet's breaks."),
      pageSetupXml: z.string().optional().describe("Raw <pageSetup> fragment; empty removes it."),
      pageMarginsXml: z
        .string()
        .optional()
        .describe("Raw <pageMargins> fragment; empty removes it."),
      printOptionsXml: z
        .string()
        .optional()
        .describe("Raw <printOptions> fragment; empty removes it."),
      headerFooterXml: z
        .string()
        .optional()
        .describe("Raw <headerFooter> fragment; empty removes it."),
      sheetPrXml: z.string().optional().describe("Raw <sheetPr> fragment; empty removes it."),
    },
  },
  ({ sessionId, sheet, ...settings }) => {
    try {
      const stated = Object.fromEntries(
        Object.entries(settings).filter(([, value]) => value !== undefined),
      );
      return ok(
        sessionPrintSettings(sessionId, sheet, Object.keys(stated).length > 0 ? stated : undefined),
      );
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_function_lookup",
  {
    title: "Function lookup",
    description: "List functions, get metadata, localize names, or canonicalize localized names.",
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
      "Low-level allowlisted access to the Formulon Workbook API for advanced features: pivot tables, style tables, merges, comments, hyperlinks, validations, conditional formats, raw print-settings XML, dependency graph, spill info, and more.",
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

server.registerTool(
  "formulon_inspect_workbook",
  {
    title: "Inspect workbook path",
    description: "Load an xlsx workbook path and return a one-shot summary.",
    inputSchema: {
      path: z.string().describe("Path to an .xlsx workbook."),
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
      "Create or load a workbook, apply zero-based cell mutations, recalculate, and save.",
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

const HELP = `formulon-mcp ${PACKAGE_VERSION}
MCP server for Formulon Excel-compatible formula and workbook evaluation.

Usage:
  formulon-mcp            Start the stdio MCP server.
  formulon-mcp init       Interactive setup: write the formulon MCP entry into
                          Claude Code, Codex CLI, and/or Claude Desktop configs.
  formulon-mcp uninstall  Interactive removal: drop the formulon MCP entry from
                          those config files.
  formulon-mcp --help     Show this help.
  formulon-mcp --version  Show version.

Docs: https://github.com/libraz/formulon-mcp
`;

const main = async (): Promise<void> => {
  const argv = process.argv.slice(2);
  if (argv.includes("--help") || argv.includes("-h")) {
    process.stdout.write(HELP);
    return;
  }
  if (argv.includes("--version") || argv.includes("-v")) {
    process.stdout.write(`${PACKAGE_VERSION}\n`);
    return;
  }
  if (argv[0] === "init") {
    await runInit();
    return;
  }
  if (argv[0] === "uninstall") {
    await runUninstall();
    return;
  }
  const transport = new StdioServerTransport();
  await server.connect(transport);
};

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  process.stderr.write(`[formulon-mcp] ${message}\n`);
  process.exit(1);
});
