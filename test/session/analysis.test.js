import assert from "node:assert/strict";
import { test } from "vitest";
import { formulonModule } from "../../dist/formulon.js";
import {
  analyzeSessionWorkbook,
  applySessionMutations,
  applySheetOperation,
  callWorkbookMethod,
  closeSession,
  detectSessionRegions,
  getSession,
  inspectSessionLayout,
  openSession,
  styleSessionRange,
  traceSession,
} from "../../dist/sessions.js";

test("lists formatted cells with no content alongside the stored cells' styles", async () => {
  await openSession(undefined, "formatted-cells");
  try {
    applySessionMutations(
      "formatted-cells",
      [
        { type: "number", a1: "Sheet1!A1", value: 1 },
        { type: "number", a1: "Sheet1!B1", value: 2 },
        { type: "number", a1: "Sheet1!B2", value: 3 },
      ],
      true,
    );
    // B1 is never styled; B2 carries an explicit XF 0. Both read XF 0, the default style.
    getSession("formatted-cells").workbook.sheet(0).ooxml.setCellXfIndex(1, 1, 0);
    styleSessionRange("formatted-cells", "A1", { numberFormat: "0.00" }, undefined, "existing");
    styleSessionRange("formatted-cells", "C3", { numberFormat: "0.0%" }, undefined, "existing");

    const [sheet] = inspectSessionLayout("formatted-cells", {
      includeCells: true,
      includeStyles: true,
      maxCells: 100,
    }).sheets;
    const byA1 = Object.fromEntries(sheet.cells.map((cell) => [cell.a1, cell]));
    assert.deepEqual(Object.keys(byA1).sort(), ["A1", "B1", "B2"]);
    assert.notEqual(byA1.A1.style.xfIndex, 0);
    for (const a1 of ["B1", "B2"]) {
      assert.equal(byA1[a1].style.xfIndex, 0);
      assert.notEqual(byA1[a1].style.xf, null);
    }
    assert.equal(sheet.formattedCellCount, 1);
    assert.equal(sheet.formattedTruncated, false);
    assert.deepEqual(
      sheet.formattedCells.map((cell) => cell.a1),
      ["C3"],
    );
    assert.notEqual(sheet.formattedCells[0].style.xfIndex, 0);
    assert.notEqual(sheet.formattedCells[0].style.xfIndex, byA1.A1.style.xfIndex);

    const plain = inspectSessionLayout("formatted-cells", {
      includeCells: true,
      includeStyles: false,
      maxCells: 100,
    }).sheets[0];
    assert.equal(plain.formattedCells, undefined);
  } finally {
    closeSession("formatted-cells");
  }
});

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
    const sheet = { kind: "Sheet", sheet: 0 };
    assert.deepEqual(callWorkbookMethod("analysis", sheet, "getPrecedents", ["B1", 1]).result, [
      { sheet: 0, row: 0, col: 0 },
    ]);
    assert.deepEqual(callWorkbookMethod("analysis", sheet, "getDependents", ["A1", 1]).result, [
      { sheet: 0, row: 0, col: 1 },
    ]);
    // A cell outside any spill region reads null.
    assert.equal(callWorkbookMethod("analysis", sheet, "getSpill", ["A1"]).result, null);

    const engine = await formulonModule();
    const metadata = engine.getFunctionMetadata("SUM");
    assert.equal(metadata.name, "SUM");
    assert.equal(engine.localizeFunctionName("SUM", "win-365-de_DE"), "SUMME");
    assert.equal(engine.canonicalizeFunctionName("SUMME", "win-365-de_DE"), "SUM");
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
    assert.equal(spill.spill, null);
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
