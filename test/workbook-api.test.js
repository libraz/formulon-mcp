import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "vitest";
import { READ_ONLY_METHODS, WORKBOOK_METHODS } from "../dist/session/workbook-call.js";
import {
  describeWorkbookApi,
  getWorkbookApiCatalog,
  searchWorkbookApi,
  WORKBOOK_API_MAX_LIMIT,
} from "../dist/workbook-api.js";
import { textPayload, withClient } from "./helpers/mcp.js";

const require = createRequire(import.meta.url);
const FORMULON_VERSION = require("@libraz/formulon/package.json").version;

test("catalogs every allowlisted installed Workbook method exactly once", () => {
  const catalog = getWorkbookApiCatalog();
  const names = catalog.methods.map((method) => method.name);
  const expected = [...WORKBOOK_METHODS].filter(
    (name) =>
      !["delete", "save", "saveAs", "saveWithDiagnostics", "setIterativeProgress"].includes(name),
  );
  assert.equal(catalog.version, FORMULON_VERSION);
  assert.deepEqual(new Set(names), new Set(expected));
  assert.equal(names.length, expected.length);
  assert.equal(new Set(names).size, names.length);
  assert.equal(names.includes("delete"), false);
  assert.equal(names.includes("setIterativeProgress"), false);
  assert.equal(names.includes("save"), false);
  assert.equal(names.includes("saveAs"), false);
  assert.equal(names.includes("saveWithDiagnostics"), false);
  for (const method of catalog.methods) {
    assert.equal(method.access, READ_ONLY_METHODS.has(method.name) ? "read" : "write");
    assert.equal(typeof method.signature, "string");
    assert.equal(typeof method.description, "string");
  }
});

test("keeps multiline date signatures and recursively describes table, CF, and pivot types", () => {
  const date = describeWorkbookApi("pivotFieldSetDateGroup");
  assert.match(date.methods[0].signature, /PivotDateGrouping/);
  assert.match(date.methods[0].signature, /PivotCalendar/);
  assert.equal(
    date.declarations.some((entry) => entry.name === "PivotDateGrouping"),
    true,
  );
  assert.equal(
    date.declarations.some((entry) => entry.name === "PivotCalendar"),
    true,
  );

  const table = describeWorkbookApi("createTable");
  assert.equal(
    table.declarations.some((entry) => entry.name === "TableInput"),
    true,
  );

  const conditional = describeWorkbookApi("addConditionalFormat");
  assert.equal(
    conditional.declarations.some((entry) => entry.name === "ConditionalFormatInput"),
    true,
  );
  assert.equal(
    conditional.declarations.some((entry) => entry.name === "CfColor"),
    true,
  );
});

test("search supports substring matching with capped limit and offset paging", () => {
  const all = searchWorkbookApi(undefined, WORKBOOK_API_MAX_LIMIT, 0);
  assert.equal(all.version, FORMULON_VERSION);
  assert.equal(all.count, Math.min(all.total, WORKBOOK_API_MAX_LIMIT));
  assert.equal(all.methods.length, all.count);

  const first = searchWorkbookApi("pivot", 2, 0);
  const second = searchWorkbookApi("pivot", 2, 2);
  assert.equal(first.methods.length <= 2, true);
  assert.equal(
    first.methods.every((method) =>
      /pivot/i.test(`${method.name} ${method.signature} ${method.description}`),
    ),
    true,
  );
  const expectedSecond = searchWorkbookApi("pivot", WORKBOOK_API_MAX_LIMIT, 0).methods.slice(2, 4);
  assert.deepEqual(
    second.methods.map((method) => method.name),
    expectedSecond.map((method) => method.name),
  );
  const capped = searchWorkbookApi(undefined, WORKBOOK_API_MAX_LIMIT + 1, 0);
  assert.equal(capped.count <= WORKBOOK_API_MAX_LIMIT, true);
  assert.equal(capped.truncated, capped.total > WORKBOOK_API_MAX_LIMIT);
  assert.equal(searchWorkbookApi("pivot", 2, 10_000).count, 0);
});

test("registers the API catalog as a session-free MCP tool", async () => {
  await withClient(async (client) => {
    const tools = await client.listTools();
    assert.equal(
      tools.tools.some((tool) => tool.name === "formulon_workbook_api"),
      true,
    );

    const search = textPayload(
      await client.callTool({
        name: "formulon_workbook_api",
        arguments: { operation: "search", query: "pivotFieldSetDateGroup", limit: 5 },
      }),
    );
    assert.equal(search.version, FORMULON_VERSION);
    assert.equal(search.count, 1);
    assert.equal(search.methods[0].name, "pivotFieldSetDateGroup");

    const described = textPayload(
      await client.callTool({
        name: "formulon_workbook_api",
        arguments: { operation: "describe", method: "createTable" },
      }),
    );
    assert.equal(described.methods[0].name, "createTable");
    assert.equal(
      described.declarations.some((entry) => entry.name === "TableInput"),
      true,
    );
  });
});
