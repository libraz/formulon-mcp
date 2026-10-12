import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { jsonText } from "../formulon.js";
import {
  getSessionImage,
  insertSessionImage,
  listSessionImages,
  moveSessionImage,
  removeSessionImage,
  reorderSessionImage,
  resolveSessionSheet,
  traceSession,
} from "../sessions.js";
import { conditionalFormatSchema, validationSchema } from "./object-schemas.js";
import { fail, methodOk, ok, rangeSchema, sheetInputSchema } from "./shared.js";

const imageLengthSchema = z.object({
  size: z.number().finite().nonnegative(),
  unit: z.enum(["pt", "px", "mm"]).optional().describe("Defaults to points."),
});

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
    "formulon_image_operation",
    {
      title: "Image operation",
      description:
        "List a sheet's drawing objects (pictures, shapes, charts) with A1 anchors and sizes, get a picture's bytes (returned as an image, optionally written to outputPath), insert a PNG/JPEG/GIF/BMP file anchored at a cell, move/resize a picture in place, change the stacking order, or remove a picture. Insertion keeps existing drawings; an omitted width or height keeps the aspect ratio, and omitting both uses the pixel size at 96 DPI. A move changes only what it states and keeps the picture's id, crop and rotation.",
      inputSchema: {
        ...sheetInputSchema,
        operation: z.enum(["list", "get", "insert", "move", "reorder", "remove"]),
        objectId: z
          .number()
          .int()
          .positive()
          .optional()
          .describe("Drawing object id from list; required for get, move, reorder and remove."),
        path: z.string().optional().describe("Image file to insert; required for insert."),
        anchor: z
          .string()
          .optional()
          .describe(
            "Top-left cell on the selected sheet, for example B2; required for insert, optional for move.",
          ),
        offsetX: imageLengthSchema.optional().describe("Offset from the anchor cell's left edge."),
        offsetY: imageLengthSchema.optional().describe("Offset from the anchor cell's top edge."),
        width: imageLengthSchema.optional(),
        height: imageLengthSchema.optional(),
        anchorKind: z
          .enum(["oneCell", "twoCell"])
          .optional()
          .describe(
            "oneCell keeps the size when rows/columns resize; twoCell stretches with them. Insert defaults to oneCell; move keeps the current kind.",
          ),
        index: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe(
            "For reorder: target position in list order, which is back to front (0 is the back).",
          ),
        name: z.string().optional(),
        description: z.string().optional().describe("Alternative text."),
        outputPath: z.string().optional().describe("For get: write the image bytes to this path."),
      },
    },
    ({ sessionId, sheet, operation, objectId, path, anchor, outputPath, index, ...placement }) => {
      try {
        if (operation === "list") {
          return ok(listSessionImages(sessionId, sheet));
        }
        if (operation === "insert") {
          if (path === undefined || anchor === undefined) {
            throw new Error("path and anchor are required for insert");
          }
          return ok(insertSessionImage(sessionId, sheet, { path, anchor, ...placement }));
        }
        if (objectId === undefined) {
          throw new Error(`objectId is required for ${operation}`);
        }
        if (operation === "remove") {
          return ok(removeSessionImage(sessionId, sheet, objectId));
        }
        if (operation === "move") {
          const { name, description, ...move } = placement;
          if (name !== undefined || description !== undefined) {
            throw new Error("name and description apply to insert only; move keeps them");
          }
          return ok(moveSessionImage(sessionId, sheet, objectId, { anchor, ...move }));
        }
        if (operation === "reorder") {
          if (index === undefined) {
            throw new Error("index is required for reorder");
          }
          return ok(reorderSessionImage(sessionId, sheet, objectId, index));
        }
        const image = getSessionImage(sessionId, sheet, objectId, outputPath);
        return {
          content: [
            ...(image.mimeType
              ? [
                  {
                    type: "image" as const,
                    data: image.bytes.toString("base64"),
                    mimeType: image.mimeType,
                  },
                ]
              : []),
            { type: "text" as const, text: jsonText(image.metadata) },
          ],
        };
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
