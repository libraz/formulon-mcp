import assert from "node:assert/strict";
import { test } from "vitest";
import { callTool, errorPayload, withClient } from "./helpers/mcp.js";

const SESSION = "error-form";

// One row per error path the MCP layer raises: [label, tool, arguments, expected status name].
const ROWS = [
  // An unknown session or sheet.
  ["unknown session", "formulon_get_cell", { sessionId: "nope", row: 0, col: 0 }, "NotFound"],
  [
    "unknown session in a mutation",
    "formulon_merge_operation",
    { sessionId: "nope", operation: "list" },
    "NotFound",
  ],
  [
    "unknown sheet name",
    "formulon_get_cell",
    { sessionId: SESSION, sheet: "Missing", row: 0, col: 0 },
    "NotFound",
  ],
  [
    "sheet index past the end",
    "formulon_get_cell",
    { sessionId: SESSION, sheet: 9, row: 0, col: 0 },
    "SheetIndexOutOfRange",
  ],
  [
    "unknown target kind",
    "formulon_workbook_call",
    { sessionId: SESSION, target: { kind: "Nope" }, method: "x", args: [] },
    "NotFound",
  ],
  [
    "unknown method",
    "formulon_workbook_call",
    { sessionId: SESSION, method: "noSuchMethod", args: [] },
    "NotFound",
  ],
  [
    "unknown API method",
    "formulon_workbook_api",
    { operation: "describe", method: "noSuchMethod" },
    "NotFound",
  ],

  // An index past a collection's bounds.
  [
    "merge index",
    "formulon_merge_operation",
    { sessionId: SESSION, operation: "removeAt", index: 9 },
    "IndexOutOfRange",
  ],
  [
    "table index (update)",
    "formulon_table_operation",
    { sessionId: SESSION, operation: "update", index: 9, name: "T" },
    "IndexOutOfRange",
  ],
  [
    "table index (remove)",
    "formulon_table_operation",
    { sessionId: SESSION, operation: "remove", index: 9 },
    "IndexOutOfRange",
  ],
  [
    "sparkline group index (remove)",
    "formulon_sparkline_operation",
    { sessionId: SESSION, operation: "remove", index: 9 },
    "IndexOutOfRange",
  ],
  [
    "sparkline group index (update)",
    "formulon_sparkline_operation",
    { sessionId: SESSION, operation: "update", index: 9, high: true },
    "IndexOutOfRange",
  ],
  [
    "data table without a table",
    "formulon_data_table_operation",
    { sessionId: SESSION, operation: "remove", cell: "A1" },
    "NotFound",
  ],
  [
    "shape id",
    "formulon_shape_operation",
    { sessionId: SESSION, operation: "remove", objectId: 99 },
    "NotFound",
  ],
  [
    "picture in cell",
    "formulon_image_operation",
    { sessionId: SESSION, operation: "getInCell", anchor: "A1" },
    "NotFound",
  ],

  // A missing operation-specific field.
  [
    "merge removeAt index",
    "formulon_merge_operation",
    { sessionId: SESSION, operation: "removeAt" },
    "InvalidArgument",
  ],
  [
    "merge add range",
    "formulon_merge_operation",
    { sessionId: SESSION, operation: "add" },
    "InvalidArgument",
  ],
  [
    "comment row/col",
    "formulon_comment_operation",
    { sessionId: SESSION, operation: "get" },
    "InvalidArgument",
  ],
  [
    "comment visible",
    "formulon_comment_operation",
    { sessionId: SESSION, operation: "visibility", row: 0, col: 0 },
    "InvalidArgument",
  ],
  [
    "hyperlink index",
    "formulon_hyperlink_operation",
    { sessionId: SESSION, operation: "removeAt" },
    "InvalidArgument",
  ],
  [
    "hyperlink row/col",
    "formulon_hyperlink_operation",
    { sessionId: SESSION, operation: "add" },
    "InvalidArgument",
  ],
  [
    "validation index",
    "formulon_validation_operation",
    { sessionId: SESSION, operation: "removeAt" },
    "InvalidArgument",
  ],
  [
    "validation rule",
    "formulon_validation_operation",
    { sessionId: SESSION, operation: "add" },
    "InvalidArgument",
  ],
  [
    "conditional format index",
    "formulon_conditional_format_operation",
    { sessionId: SESSION, operation: "removeAt" },
    "InvalidArgument",
  ],
  [
    "conditional format rule",
    "formulon_conditional_format_operation",
    { sessionId: SESSION, operation: "add" },
    "InvalidArgument",
  ],
  [
    "conditional format range",
    "formulon_conditional_format_operation",
    { sessionId: SESSION, operation: "evaluate" },
    "InvalidArgument",
  ],
  [
    "image anchor",
    "formulon_image_operation",
    { sessionId: SESSION, operation: "insertInCell" },
    "InvalidArgument",
  ],
  [
    "image path",
    "formulon_image_operation",
    { sessionId: SESSION, operation: "insert" },
    "InvalidArgument",
  ],
  [
    "image object id",
    "formulon_image_operation",
    { sessionId: SESSION, operation: "remove" },
    "InvalidArgument",
  ],
  [
    "shape anchor",
    "formulon_shape_operation",
    { sessionId: SESSION, operation: "insert" },
    "InvalidArgument",
  ],
  [
    "shape object id",
    "formulon_shape_operation",
    { sessionId: SESSION, operation: "remove" },
    "InvalidArgument",
  ],
  [
    "sparkline list",
    "formulon_sparkline_operation",
    { sessionId: SESSION, operation: "add" },
    "InvalidArgument",
  ],
  [
    "sparkline index",
    "formulon_sparkline_operation",
    { sessionId: SESSION, operation: "remove" },
    "InvalidArgument",
  ],
  [
    "data table range",
    "formulon_data_table_operation",
    { sessionId: SESSION, operation: "add" },
    "InvalidArgument",
  ],
  [
    "data table cell",
    "formulon_data_table_operation",
    { sessionId: SESSION, operation: "remove" },
    "InvalidArgument",
  ],
  [
    "table index",
    "formulon_table_operation",
    { sessionId: SESSION, operation: "remove" },
    "InvalidArgument",
  ],
  [
    "function formula",
    "formulon_function_lookup",
    { sessionId: SESSION, operation: "localizeFormula" },
    "InvalidArgument",
  ],
  [
    "function name",
    "formulon_function_lookup",
    { sessionId: SESSION, operation: "metadata" },
    "InvalidArgument",
  ],
  ["API method", "formulon_workbook_api", { operation: "describe" }, "InvalidArgument"],
  ["cell row/col", "formulon_get_cell", { sessionId: SESSION }, "InvalidArgument"],
  ["cell path", "formulon_get_cell", { a1: "A1" }, "InvalidArgument"],
  [
    "cell session and path",
    "formulon_get_cell",
    { sessionId: SESSION, path: "x.xlsx", a1: "A1" },
    "InvalidArgument",
  ],
  [
    "evaluate sheet name without a session",
    "formulon_eval_formula",
    { formula: "=1", sheet: "Sheet1" },
    "InvalidArgument",
  ],

  // A bad combination or malformed value.
  [
    "set cell",
    "formulon_set_cells",
    { sessionId: SESSION, mutations: [{ type: "text", value: "x" }] },
    "InvalidArgument",
  ],
  ["search query", "formulon_find_cells", { sessionId: SESSION, query: "" }, "InvalidArgument"],
  [
    "audit window and minPeers",
    "formulon_audit_formulas",
    { sessionId: SESSION, window: 2, minPeers: 10 },
    "InvalidArgument",
  ],
];

test("errors the server raises use the engine's status form", async () => {
  await withClient(async (client) => {
    await callTool(client, "formulon_open_workbook", { sessionId: SESSION });
    for (const [label, tool, args, status] of ROWS) {
      const result = await callTool(client, tool, args);
      assert.equal(result.isError, true, `${label}: ${JSON.stringify(result.content)}`);
      const text = errorPayload(result);
      assert.match(text, /^[A-Z][A-Za-z]+: /, `${label}: ${text}`);
      assert.ok(text.startsWith(`${status}: `), `${label}: ${text}`);
      assert.equal(result.content.length, 2, `${label}: ${text}`);
      assert.equal(result.content[1].type, "text", label);
      const detail = JSON.parse(result.content[1].text);
      assert.equal(detail.code, status, label);
      assert.equal(typeof detail.message, "string", label);
      assert.equal(typeof detail.context, "string", label);
    }
  });
});
