import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { formulonModule, jsonText } from "../formulon.js";
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
import {
  conditionalFormatInput,
  conditionalFormatSchema,
  validationInput,
  validationSchema,
} from "./object-schemas.js";
import { fail, imageLengthSchema, labeledOk, ok, rangeSchema, sheetInputSchema } from "./shared.js";

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
          return labeledOk(sessionId, "Sheet.listMerges", "read", (wb) =>
            wb.sheet(sheetIndex).listMerges(),
          );
        }
        if (operation === "clear") {
          return labeledOk(sessionId, "Sheet.clearMerges", "write", (wb) =>
            wb.sheet(sheetIndex).clearMerges(),
          );
        }
        if (operation === "removeAt") {
          if (index === undefined) {
            throw new Error("index is required for removeAt");
          }
          return labeledOk(sessionId, "Sheet.removeMerge", "write", (wb) => {
            const handle = wb.sheet(sheetIndex);
            const merges = handle.listMerges();
            if (index >= merges.length) {
              throw new Error(`merge index out of range: ${index} (${merges.length} merges)`);
            }
            handle.removeMerge(merges[index]);
          });
        }
        if (!range) {
          throw new Error("range is required for add/remove");
        }
        if (operation === "add") {
          return labeledOk(sessionId, "Sheet.addMerge", "write", (wb) =>
            wb.sheet(sheetIndex).addMerge(range),
          );
        }
        return labeledOk(sessionId, "Sheet.removeMerge", "write", (wb) =>
          wb.sheet(sheetIndex).removeMerge(range),
        );
      } catch (error) {
        return fail(error);
      }
    },
  );

  server.registerTool(
    "formulon_comment_operation",
    {
      title: "Comment operation",
      description:
        "List, get, set, or remove a cell comment, or show/hide it. A shown comment stays open on the sheet; a hidden one appears on hover.",
      inputSchema: {
        ...sheetInputSchema,
        operation: z.enum(["list", "get", "set", "remove", "visibility"]),
        row: z.number().int().nonnegative().optional(),
        col: z.number().int().nonnegative().optional(),
        author: z.string().optional(),
        text: z.string().optional(),
        visible: z
          .boolean()
          .optional()
          .describe("For visibility: true shows the comment, false hides it."),
      },
    },
    ({ sessionId, sheet, operation, row, col, author, text, visible }) => {
      try {
        const sheetIndex = resolveSessionSheet(sessionId, sheet);
        if (operation === "list") {
          return labeledOk(sessionId, "Sheet.listComments", "read", (wb) =>
            wb.sheet(sheetIndex).listComments(),
          );
        }
        if (row === undefined || col === undefined) {
          throw new Error("row and col are required");
        }
        if (operation === "get") {
          return labeledOk(sessionId, "Sheet.getComment", "read", (wb) =>
            wb.sheet(sheetIndex).getComment(row, col),
          );
        }
        if (operation === "visibility") {
          if (visible === undefined) {
            throw new Error("visible is required for visibility");
          }
          return labeledOk(sessionId, "Sheet.setCommentVisible", "write", (wb) =>
            wb.sheet(sheetIndex).setCommentVisible(row, col, visible),
          );
        }
        return labeledOk(sessionId, "Sheet.setComment", "write", (wb) =>
          wb
            .sheet(sheetIndex)
            .setComment(
              row,
              col,
              operation === "remove" ? null : { text: text ?? "", author: author ?? null },
            ),
        );
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
        "List, add, remove, remove by index, or clear hyperlinks on a sheet. add and remove take one cell at row/col, or the rectangle through lastRow/lastCol. An added link replaces a link covering exactly the same cells and is otherwise added beside the others, overlapping ones included; every cell it covers takes Excel's Hyperlink cell style (the font becomes the Normal font underlined in theme colour 10, other formatting is kept), and an empty top-left cell receives display, else the target, else the location, as text. remove deletes every link that intersects the cells, each one whole, and resets every cell of each removed link to the default format; clear does the same for every link. removeAt deletes only the link at that list index and touches no cell formats.",
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
          .describe("Inclusive last row of the cells to add or remove; defaults to row."),
        lastCol: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe("Inclusive last column of the cells to add or remove; defaults to col."),
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
          return labeledOk(sessionId, "Sheet.listHyperlinks", "read", (wb) =>
            wb.sheet(sheetIndex).listHyperlinks(),
          );
        }
        if (operation === "clear") {
          return labeledOk(sessionId, "Sheet.clearHyperlinks", "write", (wb) =>
            wb.sheet(sheetIndex).clearHyperlinks(),
          );
        }
        if (operation === "removeAt") {
          if (index === undefined) {
            throw new Error("index is required for removeAt");
          }
          return labeledOk(sessionId, "Sheet.removeHyperlinkAt", "write", (wb) =>
            wb.sheet(sheetIndex).removeHyperlinkAt(index),
          );
        }
        if (row === undefined || col === undefined) {
          throw new Error("row and col are required");
        }
        const range = {
          firstRow: row,
          firstCol: col,
          lastRow: lastRow ?? row,
          lastCol: lastCol ?? col,
        };
        if (operation === "remove") {
          return labeledOk(sessionId, "Sheet.removeHyperlink", "write", (wb) =>
            wb.sheet(sheetIndex).removeHyperlink(range),
          );
        }
        return labeledOk(sessionId, "Sheet.setHyperlink", "write", (wb) =>
          wb.sheet(sheetIndex).setHyperlink({
            target: target || null,
            display: display || null,
            tooltip: tooltip || null,
            location: location || null,
            range,
          }),
        );
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
          return labeledOk(sessionId, "Sheet.listValidations", "read", (wb) =>
            wb.sheet(sheetIndex).listValidations(),
          );
        }
        if (operation === "clear") {
          return labeledOk(sessionId, "Sheet.clearValidations", "write", (wb) =>
            wb.sheet(sheetIndex).clearValidations(),
          );
        }
        if (operation === "removeAt") {
          if (index === undefined) {
            throw new Error("index is required for removeAt");
          }
          return labeledOk(sessionId, "Sheet.removeValidation", "write", (wb) =>
            wb.sheet(sheetIndex).removeValidation(index),
          );
        }
        if (!validation) {
          throw new Error("validation is required for add");
        }
        const input = validationInput(validation);
        return labeledOk(sessionId, "Sheet.addValidation", "write", (wb) =>
          wb.sheet(sheetIndex).addValidation(input),
        );
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
            "Excel date serial in the workbook's date system that time-period rules compare with; omitted, the pinned clock, else today.",
          ),
      },
    },
    async ({
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
          return labeledOk(sessionId, "Sheet.listCfRules", "read", (wb) =>
            wb.sheet(sheetIndex).listCfRules(),
          );
        }
        if (operation === "clear") {
          return labeledOk(sessionId, "Sheet.clearCfRules", "write", (wb) =>
            wb.sheet(sheetIndex).clearCfRules(),
          );
        }
        if (operation === "removeAt") {
          if (index === undefined) {
            throw new Error("index is required for removeAt");
          }
          return labeledOk(sessionId, "Sheet.removeCfRule", "write", (wb) =>
            wb.sheet(sheetIndex).removeCfRule(index),
          );
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
          const engine = await formulonModule();
          return labeledOk(sessionId, "Sheet.evaluateCfRange", "read", (wb) =>
            wb.sheet(sheetIndex).evaluateCfRange(
              { firstRow, firstCol, lastRow, lastCol },
              todaySerial === undefined
                ? undefined
                : {
                    today: engine.convertSerialToCivil(todaySerial, {
                      date1904: wb.getDate1904(),
                    }),
                  },
            ),
          );
        }
        if (!rule) {
          throw new Error("rule is required for add");
        }
        const input = conditionalFormatInput(rule);
        return labeledOk(sessionId, "Sheet.addCfRule", "write", (wb) =>
          wb.sheet(sheetIndex).addCfRule(input),
        );
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
        "List a sheet's drawing objects (pictures, shapes, charts) with A1 anchors and sizes, get a picture's bytes (returned as an image, optionally written to outputPath), insert a PNG/JPEG/GIF/BMP file anchored at a cell, move/resize a picture or chart in place, change the stacking order, or remove a picture or chart; shapes are edited with formulon_shape_operation. Insertion keeps existing drawings; an omitted width or height keeps the aspect ratio, and omitting both uses the pixel size at 96 DPI. A move changes only what it states and keeps the object's id; a picture also keeps its crop and rotation.",
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
    async ({
      sessionId,
      sheet,
      operation,
      objectId,
      path,
      anchor,
      outputPath,
      index,
      ...placement
    }) => {
      try {
        if (operation === "list") {
          return ok(listSessionImages(sessionId, sheet));
        }
        if (operation === "insert") {
          if (path === undefined || anchor === undefined) {
            throw new Error("path and anchor are required for insert");
          }
          return ok(await insertSessionImage(sessionId, sheet, { path, anchor, ...placement }));
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
        "Read precedents, dependents, or spill info for a cell. Results carry A1 references (with sheet names), not raw indices. spillInfo returns spill null when the cell is in no dynamic-array spill range and is no PivotTable report's anchor, else { blocked, cause, anchor, rows, cols, range }. A report is read only at its anchor, and its range includes its page-field rows above the anchor. At the anchor of a dynamic array refused with #SPILL! or of a blocked report, blocked is true, range is the rectangle it needs clipped to the sheet, and cause is occupied, merged or beyondEdge; cause is null when blocked is false.",
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
