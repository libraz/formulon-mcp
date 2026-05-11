#!/usr/bin/env node
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
import {
  analyzeSessionWorkbook,
  applySessionMutations,
  applySheetOperation,
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
  recalcSession,
  replaceSessionCells,
  saveSession,
  setSessionDefinedName,
  setSessionSheetView,
} from "./sessions.js";

const server = new McpServer({
  name: "formuron-mcp",
  version: "0.1.0",
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
  sheet: z.number().int().nonnegative().default(0),
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
      return ok({ version: module.versionString() });
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_eval_formula",
  {
    title: "Evaluate formula",
    description: "Evaluate one Excel formula with Formulon.",
    inputSchema: {
      formula: z.string().describe("Excel formula, with or without a leading '='."),
    },
  },
  async ({ formula }) => {
    try {
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
      "Create an in-memory workbook session from an existing xlsx file or a new default workbook.",
    inputSchema: {
      path: z
        .string()
        .optional()
        .describe("Optional .xlsx path. Omit to create a new default workbook."),
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
      "Add, replace, or remove a workbook-scoped defined name. Empty formula removes it.",
    inputSchema: {
      sessionId: z.string(),
      name: z.string(),
      formula: z.string().describe("Formula with or without '='; pass empty string to remove."),
    },
  },
  ({ sessionId, name, formula }) => {
    try {
      return ok(setSessionDefinedName(sessionId, name, formula));
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
    description: "Set sheet zoom, frozen panes, or tab hidden flag.",
    inputSchema: {
      sessionId: z.string(),
      sheet: sheetRefSchema.optional().default(0),
      zoom: z.number().int().min(10).max(400).optional(),
      freezeRows: z.number().int().nonnegative().optional(),
      freezeCols: z.number().int().nonnegative().optional(),
      hidden: z.boolean().optional(),
    },
  },
  ({ sessionId, sheet, zoom, freezeRows, freezeCols, hidden }) => {
    try {
      return ok(setSessionSheetView(sessionId, sheet, { zoom, freezeRows, freezeCols, hidden }));
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
      "Get one cell from either an open session or a workbook path. A1 refs are supported for sessions.",
    inputSchema: {
      sessionId: z.string().optional(),
      path: z.string().optional(),
      a1: z.string().optional(),
      sheet: sheetRefSchema.optional().default(0),
      row: z.number().int().nonnegative().optional(),
      col: z.number().int().nonnegative().optional(),
      recalc: z.boolean().default(true),
    },
  },
  async ({ sessionId, path, a1, sheet, row, col, recalc }) => {
    let workbook: Workbook | undefined;
    try {
      if (sessionId) {
        if (a1) {
          return ok(getSessionCellByA1(sessionId, a1));
        }
        if (row === undefined || col === undefined) {
          throw new Error("row and col are required when a1 is omitted");
        }
        return ok(getSessionCell(sessionId, sheet, row, col));
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
    description: "Get a rectangular A1 range from an open workbook session.",
    inputSchema: {
      sessionId: z.string(),
      range: z.string().describe("A1 range, for example Sheet1!A1:C10 or A1:C10."),
    },
  },
  ({ sessionId, range }) => {
    try {
      return ok(getSessionRange(sessionId, range));
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_save_session",
  {
    title: "Save session workbook",
    description: "Save an open workbook session to xlsx.",
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
      if (operation === "list") {
        return methodOk(sessionId, "getMerges", [sheet]);
      }
      if (operation === "clear") {
        return methodOk(sessionId, "clearMerges", [sheet]);
      }
      if (operation === "removeAt") {
        if (index === undefined) {
          throw new Error("index is required for removeAt");
        }
        return methodOk(sessionId, "removeMergeAt", [sheet, index]);
      }
      if (!range) {
        throw new Error("range is required for add/remove");
      }
      return methodOk(sessionId, operation === "add" ? "addMerge" : "removeMerge", [sheet, range]);
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_comment_operation",
  {
    title: "Comment operation",
    description: "Get, set, or remove a cell comment.",
    inputSchema: {
      ...sheetInputSchema,
      operation: z.enum(["get", "set", "remove"]),
      row: z.number().int().nonnegative(),
      col: z.number().int().nonnegative(),
      author: z.string().optional(),
      text: z.string().optional(),
    },
  },
  ({ sessionId, sheet, operation, row, col, author, text }) => {
    try {
      if (operation === "get") {
        return methodOk(sessionId, "getComment", [sheet, row, col]);
      }
      return methodOk(sessionId, "setComment", [
        sheet,
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
    description: "List, add, remove, remove by index, or clear hyperlinks on a sheet.",
    inputSchema: {
      ...sheetInputSchema,
      operation: z.enum(["list", "add", "remove", "removeAt", "clear"]),
      row: z.number().int().nonnegative().optional(),
      col: z.number().int().nonnegative().optional(),
      index: z.number().int().nonnegative().optional(),
      target: z.string().optional(),
      display: z.string().optional(),
      tooltip: z.string().optional(),
    },
  },
  ({ sessionId, sheet, operation, row, col, index, target, display, tooltip }) => {
    try {
      if (operation === "list") {
        return methodOk(sessionId, "getHyperlinks", [sheet]);
      }
      if (operation === "clear") {
        return methodOk(sessionId, "clearHyperlinks", [sheet]);
      }
      if (operation === "removeAt") {
        if (index === undefined) {
          throw new Error("index is required for removeAt");
        }
        return methodOk(sessionId, "removeHyperlinkAt", [sheet, index]);
      }
      if (row === undefined || col === undefined) {
        throw new Error("row and col are required");
      }
      if (operation === "remove") {
        return methodOk(sessionId, "removeHyperlink", [sheet, row, col]);
      }
      return methodOk(sessionId, "addHyperlink", [
        sheet,
        row,
        col,
        target ?? "",
        display ?? "",
        tooltip ?? "",
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
    description: "List, add, remove by index, or clear data validations on a sheet.",
    inputSchema: {
      ...sheetInputSchema,
      operation: z.enum(["list", "add", "removeAt", "clear"]),
      index: z.number().int().nonnegative().optional(),
      validation: z.record(z.string(), z.unknown()).optional(),
    },
  },
  ({ sessionId, sheet, operation, index, validation }) => {
    try {
      if (operation === "list") {
        return methodOk(sessionId, "getValidations", [sheet]);
      }
      if (operation === "clear") {
        return methodOk(sessionId, "clearValidations", [sheet]);
      }
      if (operation === "removeAt") {
        if (index === undefined) {
          throw new Error("index is required for removeAt");
        }
        return methodOk(sessionId, "removeValidationAt", [sheet, index]);
      }
      if (!validation) {
        throw new Error("validation is required for add");
      }
      return methodOk(sessionId, "addValidation", [sheet, validation]);
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
      if (operation === "list") {
        return methodOk(sessionId, "getConditionalFormats", [sheet]);
      }
      if (operation === "clear") {
        return methodOk(sessionId, "clearConditionalFormats", [sheet]);
      }
      if (operation === "removeAt") {
        if (index === undefined) {
          throw new Error("index is required for removeAt");
        }
        return methodOk(sessionId, "removeConditionalFormatAt", [sheet, index]);
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
          sheet,
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
      return methodOk(sessionId, "addConditionalFormat", [sheet, rule]);
    } catch (error) {
      return fail(error);
    }
  },
);

server.registerTool(
  "formulon_trace",
  {
    title: "Trace dependencies",
    description: "Read precedents, dependents, or spill info for a cell.",
    inputSchema: {
      ...sheetInputSchema,
      operation: z.enum(["precedents", "dependents", "spillInfo"]),
      row: z.number().int().nonnegative(),
      col: z.number().int().nonnegative(),
      depth: z.number().int().positive().max(32).default(1),
    },
  },
  ({ sessionId, sheet, operation, row, col, depth }) => {
    try {
      if (operation === "spillInfo") {
        return methodOk(sessionId, "spillInfo", [sheet, row, col]);
      }
      return methodOk(sessionId, operation, [sheet, row, col, depth]);
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
      "Low-level allowlisted access to the Formulon Workbook API for advanced features: pivot tables, styles, merges, comments, hyperlinks, validations, conditional formats, dependency graph, spill info, and more.",
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
      const bytes = await saveWorkbook(workbook, outputPath);
      return ok({
        outputPath,
        bytes,
        summary: workbookSummary(workbook, false, 0),
      });
    } catch (error) {
      return fail(error);
    } finally {
      workbook?.delete();
    }
  },
);

const transport = new StdioServerTransport();
await server.connect(transport);
