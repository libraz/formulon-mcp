import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "vitest";
import {
  describeWorkbookApi,
  getWorkbookApiCatalog,
  searchWorkbookApi,
  WORKBOOK_API_MAX_LIMIT,
} from "../dist/workbook-api.js";
import { textPayload, withClient } from "./helpers/mcp.js";

const require = createRequire(import.meta.url);
const FORMULON_VERSION = require("@libraz/formulon/package.json").version;

const qualified = (method) => `${method.target.kind}.${method.name}`;

test("catalogs each callable method of every handle kind exactly once", () => {
  const catalog = getWorkbookApiCatalog();
  assert.equal(catalog.version, FORMULON_VERSION);
  const labels = catalog.methods.map(qualified);
  assert.equal(new Set(labels).size, labels.length);

  const kinds = new Map(catalog.kinds.map((kind) => [kind.kind, kind]));
  assert.deepEqual([...kinds.keys()].sort(), [
    "Chart",
    "Drawing",
    "Picture",
    "Pivot",
    "PivotCache",
    "Shape",
    "Sheet",
    "SheetOoxml",
    "Slicer",
    "Table",
    "Workbook",
    "WorkbookOoxml",
  ]);
  for (const method of catalog.methods) {
    const kind = kinds.get(method.target.kind);
    assert.ok(kind, `${qualified(method)} has an unknown kind`);
    assert.deepEqual(method.target.keys, kind.keys);
    assert.ok(method.access === "read" || method.access === "write");
    assert.equal(typeof method.signature, "string");
    assert.equal(typeof method.description, "string");
    // Members that need raw bytes, a callback or an iterator are left out.
    assert.doesNotMatch(method.signature, /Uint8Array|BytesInput|=>|Iterator</);
  }

  // Session-owned members: the session closes and saves its own workbook.
  for (const method of catalog.methods) {
    assert.doesNotMatch(method.name, /^(dispose|save)/, qualified(method));
  }
  assert.equal(labels.includes("Workbook.dispose"), false);
  assert.equal(labels.includes("Workbook.saveWithDiagnostics"), false);
  assert.equal(labels.includes("SheetOoxml.snapshotRangeXfs"), false);
  assert.equal(labels.includes("SheetOoxml.restoreRangeXfs"), false);
  assert.equal(labels.includes("Picture.snapshot"), false);
});

test("lists a chart with the members it inherits from Drawing", () => {
  const catalog = getWorkbookApiCatalog();
  const names = (kind) =>
    new Set(catalog.methods.filter((m) => m.target.kind === kind).map((m) => m.name));
  const drawing = names("Drawing");
  const chart = names("Chart");
  assert.ok(drawing.size > 0);
  for (const name of drawing) {
    assert.equal(chart.has(name), true, `Chart lacks ${name}`);
  }
  assert.equal(chart.has("setType"), true);
  assert.equal(drawing.has("setType"), false);
});

test("reads a method that only hands out a handle", () => {
  const access = new Map(getWorkbookApiCatalog().methods.map((m) => [qualified(m), m.access]));
  for (const label of ["Workbook.sheet", "Sheet.pivot", "Sheet.chart", "Workbook.pivotCache"]) {
    assert.equal(access.get(label), "read", label);
  }
  assert.equal(access.get("Sheet.addPivot"), "write");
  assert.equal(access.get("Sheet.listMerges"), "read");
  assert.equal(access.get("Sheet.addMerge"), "write");
});

test("keeps overloads together and describes table, CF, and pivot input types", () => {
  const date = describeWorkbookApi("Pivot.setFieldDateGroup");
  assert.equal(date.methods[0].target.kind, "Pivot");
  assert.match(date.methods[0].signature, /PivotDateGroupInput/);
  for (const name of ["PivotDateGroupInput", "PivotCalendar"]) {
    assert.equal(
      date.declarations.some((entry) => entry.name === name),
      true,
      name,
    );
  }

  const table = describeWorkbookApi("Sheet.addTable");
  assert.equal(
    table.declarations.some((entry) => entry.name === "TableInput"),
    true,
  );

  const conditional = describeWorkbookApi("Sheet.addCfRule");
  // Overloads are merged: one signature line each.
  assert.equal(conditional.methods[0].signature.split("\n").length >= 2, true);
  assert.match(conditional.methods[0].signature, /ConditionalFormatInput/);
  for (const name of ["ConditionalFormatInput", "CfIconSetRuleInput", "IconSetName"]) {
    assert.equal(
      conditional.declarations.some((entry) => entry.name === name),
      true,
      name,
    );
  }
});

test("describes a method by Kind.method or by a bare name only one kind has", () => {
  const qualifiedForm = describeWorkbookApi("Sheet.listMerges");
  assert.equal(qualifiedForm.count, 1);
  assert.equal(qualifiedForm.methods[0].name, "listMerges");
  assert.deepEqual(qualifiedForm.methods[0].target, { kind: "Sheet", keys: ["sheet"] });

  const bare = describeWorkbookApi("listMerges");
  assert.deepEqual(bare.methods, qualifiedForm.methods);

  // getProperties belongs to several kinds, so the bare name is ambiguous.
  assert.throws(
    () => describeWorkbookApi("getProperties"),
    /exists on several targets; name one of .*Pivot\.getProperties/,
  );
  assert.equal(describeWorkbookApi("Pivot.getProperties").methods[0].target.kind, "Pivot");
  assert.throws(() => describeWorkbookApi("Sheet.noSuchMethod"), /unknown or unavailable/);
  assert.throws(() => describeWorkbookApi("noSuchMethod"), /unknown or unavailable/);
  assert.throws(() => describeWorkbookApi("Workbook.save"), /unknown or unavailable/);
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
      /pivot/i.test(
        `${method.name} ${method.target.kind} ${method.signature} ${method.description}`,
      ),
    ),
    true,
  );
  const expectedSecond = searchWorkbookApi("pivot", WORKBOOK_API_MAX_LIMIT, 0).methods.slice(2, 4);
  assert.deepEqual(second.methods.map(qualified), expectedSecond.map(qualified));
  const capped = searchWorkbookApi(undefined, WORKBOOK_API_MAX_LIMIT + 1, 0);
  assert.equal(capped.count <= WORKBOOK_API_MAX_LIMIT, true);
  assert.equal(capped.truncated, capped.total > WORKBOOK_API_MAX_LIMIT);
  assert.equal(searchWorkbookApi("pivot", 2, 10_000).count, 0);

  // A kind name matches the methods that kind carries.
  const chart = searchWorkbookApi("chart", WORKBOOK_API_MAX_LIMIT, 0);
  assert.equal(
    chart.methods.some((method) => method.target.kind === "Chart"),
    true,
  );
  assert.throws(() => searchWorkbookApi("pivot", 0, 0), /limit must be an integer/);
  assert.throws(() => searchWorkbookApi("pivot", 2, -1), /offset must be a non-negative/);
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
        arguments: { operation: "search", query: "setFieldDateGroup", limit: 5 },
      }),
    );
    assert.equal(search.version, FORMULON_VERSION);
    assert.equal(search.count, 1);
    assert.equal(search.methods[0].name, "setFieldDateGroup");
    assert.deepEqual(search.methods[0].target, { kind: "Pivot", keys: ["sheet", "pivot"] });

    const described = textPayload(
      await client.callTool({
        name: "formulon_workbook_api",
        arguments: { operation: "describe", method: "Sheet.addTable" },
      }),
    );
    assert.equal(described.methods[0].name, "addTable");
    assert.deepEqual(described.methods[0].target, { kind: "Sheet", keys: ["sheet"] });
    assert.equal(
      described.declarations.some((entry) => entry.name === "TableInput"),
      true,
    );

    const bare = textPayload(
      await client.callTool({
        name: "formulon_workbook_api",
        arguments: { operation: "describe", method: "addTable" },
      }),
    );
    assert.equal(bare.methods[0].name, "addTable");
  });
});
