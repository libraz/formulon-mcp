import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { resolveSessionSheet, traceSession } from "../sessions.js";
import { conditionalFormatSchema, validationSchema } from "./object-schemas.js";
import { fail, methodOk, ok, rangeSchema, sheetInputSchema } from "./shared.js";

export function registerObjectTools(server: McpServer): void {
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
          return methodOk(sessionId, "getCommentResult", [sheetIndex, row, col]);
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
        target: z
          .string()
          .optional()
          .describe("External URL; leave empty for an in-workbook link."),
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
        validation: validationSchema.optional(),
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
        rule: conditionalFormatSchema.optional(),
        firstRow: z.number().int().nonnegative().optional(),
        firstCol: z.number().int().nonnegative().optional(),
        lastRow: z.number().int().nonnegative().optional(),
        lastCol: z.number().int().nonnegative().optional(),
        todaySerial: z
          .number()
          .optional()
          .describe(
            "Excel date serial for time-period evaluation; omitted disables time-period rules.",
          ),
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
}
