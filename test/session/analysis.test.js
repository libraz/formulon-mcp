import assert from "node:assert/strict";
import { test } from "vitest";
import {
  analyzeSessionWorkbook,
  applySessionMutations,
  applySheetOperation,
  callWorkbookMethod,
  closeSession,
  detectSessionRegions,
  inspectSessionLayout,
  openSession,
  traceSession,
} from "../../dist/sessions.js";

test("inspects layout, detects regions, and analyzes workbook shape across sheets", async () => {
  await openSession(undefined, "high-level");
  try {
    applySheetOperation("high-level", "add", { name: "Data" });
    applySessionMutations(
      "high-level",
      [
        { type: "text", a1: "Sheet1!A1", value: "Statement" },
        { type: "text", a1: "Sheet1!F2", value: "Issued" },
        { type: "text", a1: "Sheet1!G2", value: "2026/05/11" },
        { type: "text", a1: "Sheet1!A4", value: "Party" },
        { type: "text", a1: "Sheet1!B4", value: "Example Corp" },
        { type: "text", a1: "Sheet1!A8", value: "Line" },
        { type: "text", a1: "Sheet1!B8", value: "Units" },
        { type: "text", a1: "Sheet1!C8", value: "Rate" },
        { type: "text", a1: "Sheet1!D8", value: "Amount" },
        { type: "text", a1: "Sheet1!A9", value: "Service fee" },
        { type: "number", a1: "Sheet1!B9", value: 2 },
        { type: "number", a1: "Sheet1!C9", value: 10000 },
        { type: "formula", a1: "Sheet1!D9", formula: "=B9*C9" },
        { type: "text", a1: "Sheet1!C11", value: "Summary" },
        { type: "formula", a1: "Sheet1!D11", formula: "=SUM(D9:D10)" },
        { type: "text", a1: "Data!A1", value: "Customer List" },
      ],
      true,
    );

    const layout = inspectSessionLayout("high-level", {
      includeCells: true,
      includeStyles: true,
      maxCells: 100,
    });
    assert.deepEqual(
      layout.sheets.map((sheet) => sheet.name),
      ["Sheet1", "Data"],
    );
    const formulaCell = layout.sheets[0].cells.find((cell) => cell.a1 === "D9");
    assert.equal(formulaCell.formula, "=B9*C9");
    assert.deepEqual(formulaCell.value, { kind: "number", value: 20000 });

    const regions = detectSessionRegions("high-level", undefined, 100);
    assert.equal(
      regions.regions.some((region) => region.type === "table"),
      true,
    );
    assert.equal(
      regions.regions.some((region) => region.type === "total"),
      true,
    );

    const analysis = analyzeSessionWorkbook("high-level", true, 100);
    assert.equal(analysis.classification.primaryType, "invoice");
    assert.equal(analysis.summary.tables.length > 0, true);
    assert.equal(
      analysis.evidence.some((entry) =>
        entry.includes("upper metadata, line-item table, and below-table summary"),
      ),
      true,
    );
  } finally {
    closeSession("high-level");
  }
});

test("supports dependency, spill-info, and function metadata APIs", async () => {
  await openSession(undefined, "analysis");
  try {
    applySessionMutations(
      "analysis",
      [
        { type: "number", a1: "Sheet1!A1", value: 2 },
        { type: "formula", a1: "Sheet1!B1", formula: "=A1+1" },
      ],
      true,
    );
    assert.deepEqual(callWorkbookMethod("analysis", "precedents", [0, 0, 1, 1]).result.items, [
      { sheet: 0, row: 0, col: 0 },
    ]);
    assert.deepEqual(callWorkbookMethod("analysis", "dependents", [0, 0, 0, 1]).result.items, [
      { sheet: 0, row: 0, col: 1 },
    ]);
    const spill = callWorkbookMethod("analysis", "spillInfo", [0, 0, 2]);
    assert.deepEqual(spill.result, {
      status: { ok: true, status: 0, message: "", context: "" },
      engaged: false,
      anchorRow: 0,
      anchorCol: 0,
      rows: 0,
      cols: 0,
    });

    const metadata = callWorkbookMethod("analysis", "functionMetadata", ["SUM"]);
    assert.equal(metadata.result.ok, true);
    assert.equal(metadata.result.name, "SUM");
    assert.equal(
      callWorkbookMethod("analysis", "localizeFunctionName", ["SUM", "win-365-de_DE"]).result.value,
      "SUMME",
    );
    assert.equal(
      callWorkbookMethod("analysis", "canonicalizeFunctionName", ["SUMME", "win-365-de_DE"]).result
        .value,
      "SUM",
    );
  } finally {
    closeSession("analysis");
  }
});

test("trace returns A1 references with sheet names", async () => {
  await openSession(undefined, "trace-a1");
  try {
    applySessionMutations(
      "trace-a1",
      [
        { type: "number", row: 0, col: 0, value: 10 },
        { type: "formula", row: 0, col: 1, formula: "=A1+1" },
      ],
      true,
    );
    const precedents = traceSession("trace-a1", "precedents", 0, 0, 1, 1);
    assert.equal(precedents.count, 1);
    assert.deepEqual(precedents.cells, [
      { sheet: 0, sheetName: "Sheet1", row: 0, col: 0, ref: "Sheet1!A1" },
    ]);
    assert.equal(precedents.cell.ref, "Sheet1!B1");

    const spill = traceSession("trace-a1", "spillInfo", 0, 0, 0, 1);
    assert.equal(typeof spill.spill.engaged, "boolean");
  } finally {
    closeSession("trace-a1");
  }
});

test("rejects invalid direct trace coordinates and depth", async () => {
  await openSession(undefined, "trace-invalid");
  try {
    assert.throws(
      () => traceSession("trace-invalid", "spillInfo", 0, -1, 0, 1),
      /coordinates must be finite nonnegative integers/,
    );
    assert.throws(
      () => traceSession("trace-invalid", "precedents", 0, 0.5, 0, 1),
      /coordinates must be finite nonnegative integers/,
    );
    assert.throws(
      () => traceSession("trace-invalid", "precedents", 0, 0, 0, Number.NaN),
      /trace depth must be an integer in 1\.\.32/,
    );
  } finally {
    closeSession("trace-invalid");
  }
});
