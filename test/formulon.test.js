import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { formulonModule, loadWorkbook } from "../dist/formulon.js";
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
  resolveSessionSheet,
  saveSession,
  sessionDimension,
  setSessionDefinedName,
  setSessionRange,
  setSessionSheetView,
  traceSession,
  WORKBOOK_METHODS,
} from "../dist/sessions.js";

const FORMULON_VERSION = createRequire(import.meta.url)("@libraz/formulon/package.json").version;

test("evaluates a formula through Formulon", async () => {
  const module = await formulonModule();
  const result = module.evalFormula("=SUM(1,2,3)");
  assert.equal(module.versionString(), FORMULON_VERSION);
  assert.equal(result.status.ok, true);
  assert.equal(result.value.kind, 1);
  assert.equal(result.value.number, 6);
});

test("manages session lifecycle", async () => {
  const session = await openSession(undefined, "lifecycle");
  assert.equal(session.id, "lifecycle");
  assert.equal(
    listSessions().some((item) => item.id === "lifecycle"),
    true,
  );
  const closed = closeSession("lifecycle");
  assert.equal(closed.id, "lifecycle");
  assert.equal(
    listSessions().some((item) => item.id === "lifecycle"),
    false,
  );
});

test("mutates cells, recalculates, and reads cells and ranges", async () => {
  await openSession(undefined, "calc");
  try {
    applySessionMutations(
      "calc",
      [
        { type: "number", a1: "Sheet1!A1", value: 41 },
        { type: "formula", a1: "Sheet1!B1", formula: "=A1+1" },
      ],
      true,
    );

    const b1 = getSessionCellByA1("calc", "Sheet1!B1");
    assert.deepEqual(b1.value, { kind: "number", value: 42 });

    const range = getSessionRange("calc", "Sheet1!A1:B1", {
      maxCells: 100,
      recalc: false,
      includeFormulas: true,
    });
    assert.deepEqual(
      range.cells.map((entry) => entry.value),
      [
        { kind: "number", value: 41 },
        { kind: "number", value: 42 },
      ],
    );
    assert.equal(range.cells[1].formula, "=A1+1");
    assert.equal(range.cells[0].formula, "");
    assert.equal(range.truncated, false);
  } finally {
    closeSession("calc");
  }
});

test("handles a moderately sized cell mutation batch", async () => {
  await openSession(undefined, "batch");
  try {
    const mutations = Array.from({ length: 100 }, (_, index) => ({
      type: "number",
      sheet: 0,
      row: index,
      col: 0,
      value: index + 1,
    }));
    applySessionMutations("batch", mutations, true);
    applySessionMutations(
      "batch",
      [{ type: "formula", a1: "Sheet1!B1", formula: "=SUM(A1:A100)" }],
      true,
    );
    assert.deepEqual(getSessionCellByA1("batch", "Sheet1!B1").value, {
      kind: "number",
      value: 5050,
    });
  } finally {
    closeSession("batch");
  }
});

test("finds and replaces text cells and formula text", async () => {
  await openSession(undefined, "search");
  try {
    applySessionMutations(
      "search",
      [
        { type: "text", a1: "Sheet1!A1", value: "Alpha budget" },
        { type: "text", a1: "Sheet1!A2", value: "beta budget" },
        { type: "number", a1: "Sheet1!B1", value: 10 },
        { type: "formula", a1: "Sheet1!B2", formula: '=CONCAT("budget ",B1)' },
      ],
      true,
    );

    const found = findSessionCells("search", "budget", {
      target: "both",
      matchCase: false,
      wholeCell: false,
      regex: false,
      maxResults: 10,
    });
    assert.deepEqual(
      found.results.map((result) => [result.a1, result.target]),
      [
        ["A1", "text"],
        ["A2", "text"],
        ["B2", "formula"],
      ],
    );

    const replaced = replaceSessionCells("search", "budget", {
      target: "texts",
      matchCase: false,
      wholeCell: false,
      regex: false,
      maxResults: 10,
      maxReplacements: 10,
      replacement: "forecast",
      recalc: true,
    });
    assert.equal(replaced.count, 2);
    assert.deepEqual(getSessionCellByA1("search", "Sheet1!A1").value, {
      kind: "text",
      value: "Alpha forecast",
    });
    assert.deepEqual(getSessionCellByA1("search", "Sheet1!A2").value, {
      kind: "text",
      value: "beta forecast",
    });

    const formulaOnly = replaceSessionCells("search", "budget", {
      target: "formulas",
      matchCase: false,
      wholeCell: false,
      regex: false,
      maxResults: 10,
      maxReplacements: 10,
      replacement: "forecast",
      recalc: true,
    });
    assert.equal(formulaOnly.count, 1);
    assert.deepEqual(getSessionCellByA1("search", "Sheet1!B2").value, {
      kind: "text",
      value: "forecast 10",
    });
  } finally {
    closeSession("search");
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

test("saves and reloads a workbook", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-test-"));
  const file = path.join(dir, "roundtrip.xlsx");
  await openSession(undefined, "roundtrip");
  try {
    applySessionMutations(
      "roundtrip",
      [{ type: "formula", a1: "Sheet1!A1", formula: "=SUM(10,20)" }],
      true,
    );
    const saved = await saveSession("roundtrip", file);
    assert.equal(saved.bytes > 0, true);
  } finally {
    closeSession("roundtrip");
  }

  const wb = await loadWorkbook(file);
  try {
    const status = wb.recalc();
    assert.equal(status.ok, true);
    const a1 = wb.getValue(0, 0, 0);
    assert.equal(a1.status.ok, true);
    assert.equal(a1.value.kind, 1);
    assert.equal(a1.value.number, 30);
  } finally {
    wb.delete();
    await rm(dir, { recursive: true, force: true });
  }
});

test("names error codes outside the classic seven", async () => {
  await openSession(undefined, "error-names");
  try {
    assert.equal(
      callWorkbookMethod("error-names", "addMerge", [
        0,
        { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 },
      ]).result.ok,
      true,
    );
    const written = applySessionMutations(
      "error-names",
      [{ type: "formula", a1: "Sheet1!A1", formula: "=SEQUENCE(2)" }],
      true,
    );
    assert.deepEqual(
      written.errorCells.map((cell) => cell.errorName),
      ["#SPILL!"],
    );
  } finally {
    closeSession("error-names");
  }
});

test("keeps the workbook method allowlist callable on the engine", async () => {
  const module = await formulonModule();
  const wb = module.Workbook.createDefault();
  try {
    const missing = Array.from(WORKBOOK_METHODS).filter((name) => typeof wb[name] !== "function");
    assert.deepEqual(missing, []);
  } finally {
    wb.delete();
  }
});

test("reports the written container and its loss counters", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-save-"));
  await openSession(undefined, "save-diagnostics");
  try {
    applySessionMutations(
      "save-diagnostics",
      [{ type: "formula", a1: "Sheet1!A1", formula: "=SUM(1,2)" }],
      true,
    );
    const xlsx = await saveSession("save-diagnostics", path.join(dir, "book.xlsx"));
    assert.equal(xlsx.format, "xlsx");
    assert.equal(xlsx.bytes > 0, true);
    assert.equal(xlsx.losses, undefined);

    const xlsb = await saveSession("save-diagnostics", path.join(dir, "book.xlsb"));
    assert.equal(xlsb.format, "xlsb");
    assert.equal(xlsb.bytes > 0, true);
  } finally {
    closeSession("save-diagnostics");
  }

  const reopened = await openSession(path.join(dir, "book.xlsb"), "save-diagnostics-reload");
  try {
    assert.equal(reopened.loadLosses, undefined);
    const a1 = getSessionCell("save-diagnostics-reload", 0, 0, 0);
    assert.equal(a1.value.value, 3);
  } finally {
    closeSession("save-diagnostics-reload");
    await rm(dir, { recursive: true, force: true });
  }
});

test("pins the workbook clock so NOW and TODAY are reproducible", async () => {
  await openSession(undefined, "clock");
  try {
    assert.equal(
      callWorkbookMethod("clock", "setPinnedNow", [2026, 8, 18, 9, 30, 0]).result.ok,
      true,
    );
    assert.deepEqual(callWorkbookMethod("clock", "pinnedNow", []).result, {
      year: 2026,
      month: 8,
      day: 18,
      hour: 9,
      minute: 30,
      second: 0,
    });
    applySessionMutations(
      "clock",
      [{ type: "formula", a1: "Sheet1!A1", formula: "=TODAY()" }],
      true,
    );
    // Excel serial 25569 is 1970-01-01, the anchor the engine and Date share.
    const expected = (Date.UTC(2026, 7, 18) - Date.UTC(1970, 0, 1)) / 86_400_000 + 25569;
    assert.equal(getSessionCell("clock", 0, 0, 0).value.value, expected);

    assert.equal(callWorkbookMethod("clock", "clearPinnedNow", []).result.ok, true);
    assert.equal(callWorkbookMethod("clock", "pinnedNow", []).result, null);
  } finally {
    closeSession("clock");
  }
});

test("creates, updates, and removes a worksheet table", async () => {
  await openSession(undefined, "tables");
  try {
    setSessionRange(
      "tables",
      "Sheet1!A1",
      [
        ["item", "qty"],
        ["pen", 2],
      ],
      0,
      false,
    );
    const created = callWorkbookMethod("tables", "createTable", [
      { sheetIndex: 0, ref: "A1:B2", name: "Items", columns: ["item", "qty"] },
    ]);
    assert.equal(created.result.status.ok, true);
    assert.equal(callWorkbookMethod("tables", "tableCount", []).result, 1);
    assert.equal(callWorkbookMethod("tables", "tableAt", [0]).result.ref, "A1:B2");

    assert.equal(
      callWorkbookMethod("tables", "updateTable", [0, { ref: "A1:B3" }]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("tables", "tableAt", [0]).result.ref, "A1:B3");
    assert.equal(callWorkbookMethod("tables", "removeTable", [0]).result.ok, true);
    assert.equal(callWorkbookMethod("tables", "tableCount", []).result, 0);
  } finally {
    closeSession("tables");
  }
});

test("round-trips a cell phonetic guide", async () => {
  await openSession(undefined, "phonetic");
  try {
    applySessionMutations("phonetic", [{ type: "text", a1: "Sheet1!A1", value: "山田" }], false);
    assert.equal(
      callWorkbookMethod("phonetic", "setCellPhonetic", [0, 0, 0, "ヤマダ"]).result.ok,
      true,
    );
    const read = callWorkbookMethod("phonetic", "getCellPhonetic", [0, 0, 0]).result;
    assert.equal(read.status.ok, true);
    assert.equal(read.value, "ヤマダ");
  } finally {
    closeSession("phonetic");
  }
});

test("supports workbook structure and metadata helpers", async () => {
  await openSession(undefined, "helpers");
  try {
    const definedName = setSessionDefinedName("helpers", "Answer", "=42");
    assert.equal(definedName.status.ok, true);

    const structure = editSessionStructure("helpers", "insertRows", 0, 0, 1);
    assert.equal(structure.status.ok, true);

    const view = setSessionSheetView("helpers", 0, {
      zoom: 120,
      freezeRows: 1,
      freezeCols: 1,
      hidden: false,
    });
    assert.equal(
      view.statuses.every((status) => status.ok),
      true,
    );

    const summary = inspectSession("helpers", false, 0);
    assert.equal(
      summary.workbook.definedNames.some((entry) => entry.name === "Answer"),
      true,
    );

    const metadata = getSessionMetadata("helpers", "functions");
    assert.equal(Array.isArray(metadata.value), true);
    assert.equal(metadata.value.includes("SUM"), true);
  } finally {
    closeSession("helpers");
  }
});

test("supports allowlisted low-level workbook calls and rejects unknown methods", async () => {
  await openSession(undefined, "low-level");
  try {
    const sheetCount = callWorkbookMethod("low-level", "sheetCount", []);
    assert.equal(sheetCount.result, 1);

    const merge = callWorkbookMethod("low-level", "addMerge", [
      0,
      { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 },
    ]);
    assert.equal(merge.result.ok, true);

    const merges = callWorkbookMethod("low-level", "getMerges", [0]);
    assert.deepEqual(merges.result, [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 }]);

    assert.throws(() => callWorkbookMethod("low-level", "constructor", []), /not allowlisted/);
  } finally {
    closeSession("low-level");
  }
});

test("supports sheet operations", async () => {
  await openSession(undefined, "sheets");
  try {
    callWorkbookMethod("sheets", "addSheet", ["Data"]);
    assert.equal(callWorkbookMethod("sheets", "sheetCount", []).result, 2);
    callWorkbookMethod("sheets", "renameSheet", [1, "Input"]);
    assert.equal(callWorkbookMethod("sheets", "sheetName", [1]).result.value, "Input");
    callWorkbookMethod("sheets", "moveSheet", [1, 0]);
    assert.equal(callWorkbookMethod("sheets", "sheetName", [0]).result.value, "Input");
    callWorkbookMethod("sheets", "removeSheet", [0]);
    assert.equal(callWorkbookMethod("sheets", "sheetCount", []).result, 1);
  } finally {
    closeSession("sheets");
  }
});

test("supports layout, comments, hyperlinks, validations, and conditional formats", async () => {
  await openSession(undefined, "objects");
  try {
    assert.equal(callWorkbookMethod("objects", "setColumnWidth", [0, 0, 1, 22]).result.ok, true);
    assert.equal(callWorkbookMethod("objects", "setColumnHidden", [0, 2, 2, true]).result.ok, true);
    const columns = callWorkbookMethod("objects", "getSheetColumns", [0]);
    assert.equal(columns.result.status.ok, true);
    assert.equal(
      columns.result.columns.some((column) => column.width === 22),
      true,
    );

    assert.equal(callWorkbookMethod("objects", "setRowHeight", [0, 0, 30]).result.ok, true);
    const rows = callWorkbookMethod("objects", "getSheetRowOverrides", [0]);
    assert.equal(rows.result.status.ok, true);
    assert.equal(
      rows.result.rows.some((row) => row.height === 30),
      true,
    );

    assert.equal(
      callWorkbookMethod("objects", "setComment", [0, 0, 0, "tester", "hello"]).result.ok,
      true,
    );
    assert.deepEqual(callWorkbookMethod("objects", "getComment", [0, 0, 0]).result, {
      author: "tester",
      text: "hello",
    });

    assert.equal(
      callWorkbookMethod("objects", "addHyperlink", [
        0,
        0,
        1,
        "https://example.com",
        "Example",
        "Example tooltip",
        "",
      ]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("objects", "addHyperlinkRange", [
        0,
        2,
        0,
        2,
        3,
        "",
        "Jump",
        "In-workbook",
        "Sheet1!A1",
      ]).result.ok,
      true,
    );
    const hyperlinks = callWorkbookMethod("objects", "getHyperlinks", [0]).result;
    assert.equal(hyperlinks.length, 2);
    assert.deepEqual(
      hyperlinks.map((link) => [link.row, link.col, link.lastRow, link.lastCol, link.location]),
      [
        [0, 1, 0, 1, ""],
        [2, 0, 2, 3, "Sheet1!A1"],
      ],
    );

    assert.equal(
      callWorkbookMethod("objects", "addValidation", [
        0,
        {
          ranges: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
          type: 1,
          op: 0,
          formula1: "1",
          formula2: "10",
        },
      ]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("objects", "getValidations", [0]).result.length, 1);

    assert.equal(
      callWorkbookMethod("objects", "addConditionalFormat", [
        0,
        {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
          type: 0,
          formula1: "=A1>0",
        },
      ]).result.status.ok,
      true,
    );
    assert.equal(callWorkbookMethod("objects", "getConditionalFormats", [0]).result.length, 1);

    assert.equal(callWorkbookMethod("objects", "removeHyperlink", [0, 0, 1]).result.ok, true);
    assert.equal(callWorkbookMethod("objects", "getHyperlinks", [0]).result.length, 1);
    assert.equal(callWorkbookMethod("objects", "clearHyperlinks", [0]).result.ok, true);
    assert.equal(callWorkbookMethod("objects", "getHyperlinks", [0]).result.length, 0);
    assert.equal(callWorkbookMethod("objects", "removeValidationAt", [0, 0]).result.ok, true);
    assert.equal(callWorkbookMethod("objects", "getValidations", [0]).result.length, 0);
    assert.equal(
      callWorkbookMethod("objects", "removeConditionalFormatAt", [0, 0]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("objects", "getConditionalFormats", [0]).result.length, 0);
    assert.equal(callWorkbookMethod("objects", "clearMerges", [0]).result.ok, true);
    assert.equal(callWorkbookMethod("objects", "getMerges", [0]).result.length, 0);
  } finally {
    closeSession("objects");
  }
});

test("round-trips comments, hyperlinks, validations, and conditional formats through xlsx", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-objects-"));
  const file = path.join(dir, "objects.xlsx");
  await openSession(undefined, "objects-roundtrip");
  try {
    assert.equal(
      callWorkbookMethod("objects-roundtrip", "setComment", [0, 0, 0, "tester", "hello"]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("objects-roundtrip", "addHyperlink", [
        0,
        0,
        1,
        "https://example.com",
        "Example",
        "Example tooltip",
        "",
      ]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("objects-roundtrip", "addValidation", [
        0,
        {
          ranges: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
          type: 1,
          op: 0,
          formula1: "1",
          formula2: "10",
        },
      ]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("objects-roundtrip", "addConditionalFormat", [
        0,
        {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
          type: 0,
          formula1: "=A1>0",
        },
      ]).result.status.ok,
      true,
    );
    await saveSession("objects-roundtrip", file);
  } finally {
    closeSession("objects-roundtrip");
  }

  const wb = await loadWorkbook(file);
  try {
    assert.deepEqual(wb.getComment(0, 0, 0), { author: "tester", text: "hello" });
    assert.equal(wb.getHyperlinks(0).length, 1);
    assert.equal(wb.getValidations(0).length, 1);
    assert.equal(wb.getConditionalFormats(0).length, 1);
  } finally {
    wb.delete();
    await rm(dir, { recursive: true, force: true });
  }
});

test("supports calc mode, profile, partial recalc, and sheet protection", async () => {
  await openSession(undefined, "settings");
  try {
    assert.equal(callWorkbookMethod("settings", "calcMode", []).result, 0);
    assert.equal(callWorkbookMethod("settings", "setCalcMode", [1]).result.ok, true);
    assert.equal(callWorkbookMethod("settings", "calcMode", []).result, 1);

    assert.equal(callWorkbookMethod("settings", "excelProfileId", []).result, "win-365-ja_JP");
    assert.equal(
      callWorkbookMethod("settings", "setExcelProfileId", ["mac-365-ja_JP"]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("settings", "excelProfileId", []).result, "mac-365-ja_JP");

    applySessionMutations(
      "settings",
      [
        { type: "number", a1: "Sheet1!A1", value: 20 },
        { type: "formula", a1: "Sheet1!B1", formula: "=A1+22" },
      ],
      false,
    );
    const partial = callWorkbookMethod("settings", "partialRecalc", [
      { sheet: 0, firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 1 },
    ]);
    assert.equal(partial.result.status.ok, true);
    assert.equal(partial.result.recomputed >= 1, true);
    assert.deepEqual(getSessionCellByA1("settings", "Sheet1!B1").value, {
      kind: "number",
      value: 42,
    });

    const protection = callWorkbookMethod("settings", "getSheetProtection", [0]).result.protection;
    const nextProtection = { ...protection, enabled: 1, sheet: 1, selectLockedCells: 1 };
    assert.equal(
      callWorkbookMethod("settings", "setSheetProtection", [0, nextProtection]).result.ok,
      true,
    );
    const protectedSheet = callWorkbookMethod("settings", "getSheetProtection", [0]);
    assert.equal(protectedSheet.result.status.ok, true);
    assert.equal(protectedSheet.result.protection.enabled, 1);
  } finally {
    closeSession("settings");
  }
});

test("supports pivot cache and pivot table creation through workbook calls", async () => {
  await openSession(undefined, "pivot");
  try {
    const cache = callWorkbookMethod("pivot", "pivotCacheCreate", [0]);
    assert.equal(cache.result.status.ok, true);
    const cacheId = cache.result.index;
    assert.equal(callWorkbookMethod("pivot", "pivotCacheCount", []).result, 1);

    assert.equal(
      callWorkbookMethod("pivot", "pivotCacheFieldAdd", [cacheId, "Region"]).result.index,
      0,
    );
    assert.equal(
      callWorkbookMethod("pivot", "pivotCacheFieldAdd", [cacheId, "Amount"]).result.index,
      1,
    );
    assert.equal(callWorkbookMethod("pivot", "pivotCacheFieldCount", [cacheId]).result, 2);
    assert.equal(
      callWorkbookMethod("pivot", "pivotCacheFieldName", [cacheId, 0]).result.value,
      "Region",
    );

    assert.equal(callWorkbookMethod("pivot", "pivotCacheRecordAdd", [cacheId]).result.index, 0);
    assert.equal(
      callWorkbookMethod("pivot", "pivotCacheRecordSetText", [cacheId, 0, 0, "East"]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("pivot", "pivotCacheRecordSetNumber", [cacheId, 0, 1, 10]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("pivot", "pivotCacheRecordCount", [cacheId]).result, 1);

    assert.equal(
      callWorkbookMethod("pivot", "pivotCreate", [0, "Pivot1", cacheId, 4, 0]).result.index,
      0,
    );
    assert.equal(callWorkbookMethod("pivot", "pivotCount", [0]).result, 1);
    assert.equal(
      callWorkbookMethod("pivot", "pivotFieldAdd", [0, 0, { sourceName: "Region", axis: 0 }]).result
        .index,
      0,
    );
    assert.equal(
      callWorkbookMethod("pivot", "pivotFieldAdd", [0, 0, { sourceName: "Amount", axis: 2 }]).result
        .index,
      1,
    );
    assert.equal(
      callWorkbookMethod("pivot", "pivotDataFieldAdd", [
        0,
        0,
        { name: "Sum of Amount", fieldIndex: 1, aggregation: 0 },
      ]).result.index,
      0,
    );

    const layout = callWorkbookMethod("pivot", "pivotLayout", [0, 0]);
    assert.equal(layout.result.status.ok, true);
    assert.equal(
      layout.result.cells.some((cell) => cell.value.kind === "number" && cell.value.value === 10),
      true,
    );

    assert.equal(
      callWorkbookMethod("pivot", "pivotFilterAdd", [
        0,
        0,
        { axis: 0, fieldName: "Region", type: 3, valueKind: 2, valueText: "Ea" },
      ]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("pivot", "pivotFilterCount", [0, 0]).result, 1);
    assert.equal(callWorkbookMethod("pivot", "pivotFilterClear", [0, 0]).result.ok, true);
    assert.equal(callWorkbookMethod("pivot", "pivotDataFieldClear", [0, 0]).result.ok, true);
    assert.equal(callWorkbookMethod("pivot", "pivotFieldClear", [0, 0]).result.ok, true);
    assert.equal(callWorkbookMethod("pivot", "pivotRemove", [0, 0]).result.ok, true);
    assert.equal(callWorkbookMethod("pivot", "pivotCount", [0]).result, 0);
    assert.equal(callWorkbookMethod("pivot", "pivotCacheRemove", [cacheId]).result.ok, true);
    assert.equal(callWorkbookMethod("pivot", "pivotCacheCount", []).result, 0);
  } finally {
    closeSession("pivot");
  }
});

test("supports empty passthrough and external-link metadata on fresh workbooks", async () => {
  await openSession(undefined, "metadata-empty");
  try {
    assert.equal(callWorkbookMethod("metadata-empty", "passthroughCount", []).result, 0);
    assert.deepEqual(callWorkbookMethod("metadata-empty", "getExternalLinks", []).result, []);
    assert.equal(callWorkbookMethod("metadata-empty", "cellStyleCount", []).result >= 0, true);
    assert.equal(callWorkbookMethod("metadata-empty", "cellStyleXfCount", []).result >= 0, true);
  } finally {
    closeSession("metadata-empty");
  }
});

test("isolates multiple sessions and rejects access after close", async () => {
  await openSession(undefined, "multi-a");
  await openSession(undefined, "multi-b");
  try {
    applySessionMutations("multi-a", [{ type: "number", a1: "Sheet1!A1", value: 1 }], true);
    applySessionMutations("multi-b", [{ type: "number", a1: "Sheet1!A1", value: 2 }], true);
    assert.deepEqual(getSessionCellByA1("multi-a", "Sheet1!A1").value, {
      kind: "number",
      value: 1,
    });
    assert.deepEqual(getSessionCellByA1("multi-b", "Sheet1!A1").value, {
      kind: "number",
      value: 2,
    });
    closeSession("multi-a");
    assert.throws(() => getSessionCellByA1("multi-a", "Sheet1!A1"), /session not found/);
  } finally {
    if (listSessions().some((session) => session.id === "multi-a")) {
      closeSession("multi-a");
    }
    if (listSessions().some((session) => session.id === "multi-b")) {
      closeSession("multi-b");
    }
  }
});

test("supports style and number format APIs", async () => {
  await openSession(undefined, "styles");
  try {
    const font = callWorkbookMethod("styles", "addFont", [
      {
        name: "Arial",
        size: 11,
        bold: true,
        italic: false,
        strike: false,
        underline: 0,
        colorArgb: 0xff112233,
      },
    ]);
    assert.equal(font.result.status.ok, true);

    const fill = callWorkbookMethod("styles", "addFill", [
      { pattern: 1, fgArgb: 0xffddeeff, bgArgb: 0xffffffff },
    ]);
    assert.equal(fill.result.status.ok, true);

    const side = { style: 1, colorArgb: 0xff000000 };
    const border = callWorkbookMethod("styles", "addBorder", [
      {
        left: side,
        right: side,
        top: side,
        bottom: side,
        diagonal: { style: 0, colorArgb: 0 },
        diagonalUp: false,
        diagonalDown: false,
      },
    ]);
    assert.equal(border.result.status.ok, true);

    const numFmt = callWorkbookMethod("styles", "addNumFmt", ["0.00"]);
    assert.equal(numFmt.result.status.ok, true);

    const xf = callWorkbookMethod("styles", "addXf", [
      {
        fontIndex: font.result.index,
        fillIndex: fill.result.index,
        borderIndex: border.result.index,
        numFmtId: numFmt.result.numFmtId,
        horizontalAlign: 2,
        verticalAlign: 1,
        wrapText: true,
      },
    ]);
    assert.equal(xf.result.status.ok, true);
    assert.equal(
      callWorkbookMethod("styles", "setCellXfIndex", [0, 0, 0, xf.result.index]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("styles", "getCellXfIndex", [0, 0, 0]).result.xfIndex,
      xf.result.index,
    );
  } finally {
    closeSession("styles");
  }
});

test("round-trips style, layout, and protection metadata through xlsx", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-style-"));
  const file = path.join(dir, "style.xlsx");
  await openSession(undefined, "style-roundtrip");
  try {
    const numFmt = callWorkbookMethod("style-roundtrip", "addNumFmt", ["0.00"]);
    const xf = callWorkbookMethod("style-roundtrip", "addXf", [
      {
        fontIndex: 0,
        fillIndex: 0,
        borderIndex: 0,
        numFmtId: numFmt.result.numFmtId,
        horizontalAlign: 2,
        verticalAlign: 1,
        wrapText: true,
      },
    ]);
    assert.equal(
      callWorkbookMethod("style-roundtrip", "setCellXfIndex", [0, 0, 0, xf.result.index]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("style-roundtrip", "setColumnWidth", [0, 0, 0, 18]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("style-roundtrip", "setRowHeight", [0, 0, 24]).result.ok, true);
    const protection = callWorkbookMethod("style-roundtrip", "getSheetProtection", [0]).result
      .protection;
    assert.equal(
      callWorkbookMethod("style-roundtrip", "setSheetProtection", [
        0,
        { ...protection, enabled: 1, sheet: 1 },
      ]).result.ok,
      true,
    );
    await saveSession("style-roundtrip", file);
  } finally {
    closeSession("style-roundtrip");
  }

  const wb = await loadWorkbook(file);
  try {
    assert.equal(wb.getCellXfIndex(0, 0, 0).status.ok, true);
    const columns = wb.getSheetColumns(0);
    assert.equal(columns.status.ok, true);
    assert.equal(columns.columns[0].width, 18);
    const rows = wb.getSheetRowOverrides(0);
    assert.equal(rows.status.ok, true);
    assert.equal(rows.rows[0].height, 24);
    assert.equal(wb.getSheetProtection(0).protection.enabled, 1);
  } finally {
    wb.delete();
    await rm(dir, { recursive: true, force: true });
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
    assert.deepEqual(callWorkbookMethod("analysis", "precedents", [0, 0, 1, 1]).result, [
      { sheet: 0, row: 0, col: 0 },
    ]);
    assert.deepEqual(callWorkbookMethod("analysis", "dependents", [0, 0, 0, 1]).result, [
      { sheet: 0, row: 0, col: 1 },
    ]);
    const spill = callWorkbookMethod("analysis", "spillInfo", [0, 0, 2]);
    assert.deepEqual(spill.result, {
      engaged: false,
      anchorRow: 0,
      anchorCol: 0,
      rows: 0,
      cols: 0,
    });

    const metadata = callWorkbookMethod("analysis", "functionMetadata", ["SUM", 0]);
    assert.equal(metadata.result.ok, true);
    assert.equal(metadata.result.name, "SUM");
    assert.equal(callWorkbookMethod("analysis", "localizeFunctionName", ["SUM", 0]).result, "SUM");
    assert.equal(
      callWorkbookMethod("analysis", "canonicalizeFunctionName", ["SUM", 0]).result,
      "SUM",
    );
  } finally {
    closeSession("analysis");
  }
});

test("recalculates sessions explicitly", async () => {
  await openSession(undefined, "recalc");
  try {
    applySessionMutations(
      "recalc",
      [
        { type: "number", a1: "Sheet1!A1", value: 5 },
        { type: "formula", a1: "Sheet1!B1", formula: "=A1*2" },
      ],
      false,
    );
    const result = recalcSession("recalc");
    assert.equal(result.status.ok, true);
    assert.deepEqual(getSessionCellByA1("recalc", "Sheet1!B1").value, {
      kind: "number",
      value: 10,
    });
  } finally {
    closeSession("recalc");
  }
});

test("recalculates formulas on open so the first read is not blank", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-open-recalc-"));
  const file = path.join(dir, "open-recalc.xlsx");
  await openSession(undefined, "open-recalc-src");
  try {
    applySessionMutations(
      "open-recalc-src",
      [
        { type: "number", a1: "Sheet1!A1", value: 21 },
        { type: "formula", a1: "Sheet1!B1", formula: "=A1*2" },
      ],
      true,
    );
    await saveSession("open-recalc-src", file);
  } finally {
    closeSession("open-recalc-src");
  }

  await openSession(file, "open-recalc-dst");
  try {
    // No explicit recalc_session: recalc-on-open must make B1 read 42, not blank.
    assert.deepEqual(getSessionCellByA1("open-recalc-dst", "Sheet1!B1").value, {
      kind: "number",
      value: 42,
    });
  } finally {
    closeSession("open-recalc-dst");
    await rm(dir, { recursive: true, force: true });
  }
});

test("annotates error cells with an Excel error literal", async () => {
  await openSession(undefined, "errname");
  try {
    const result = applySessionMutations(
      "errname",
      [
        { type: "formula", a1: "Sheet1!A1", formula: "=1/0" },
        { type: "formula", a1: "Sheet1!A2", formula: "=NOTAFUNC(1)" },
        { type: "number", a1: "Sheet1!A3", value: 5 },
      ],
      true,
    );
    assert.deepEqual(result.errorCells, [
      { index: 0, ref: "Sheet1!A1", errorName: "#DIV/0!" },
      { index: 1, ref: "Sheet1!A2", errorName: "#NAME?" },
    ]);
    const a1 = getSessionCellByA1("errname", "Sheet1!A1");
    assert.equal(a1.value.kind, "error");
    assert.equal(a1.value.errorName, "#DIV/0!");
  } finally {
    closeSession("errname");
  }
});

test("get_range clips to used range, drops blanks, and caps output", async () => {
  await openSession(undefined, "range-cap");
  try {
    applySessionMutations(
      "range-cap",
      [
        { type: "number", a1: "Sheet1!A1", value: 1 },
        { type: "number", a1: "Sheet1!C3", value: 9 },
      ],
      true,
    );
    // A generous request is clipped to the A1:C3 used range and skips blanks.
    const full = getSessionRange("range-cap", "Sheet1!A1:Z1000", {
      maxCells: 100,
      recalc: false,
      includeFormulas: false,
    });
    assert.equal(full.cells.length, 2);
    assert.equal(full.truncated, false);
    assert.deepEqual(
      full.cells.map((c) => c.a1),
      ["A1", "C3"],
    );
    // A tiny cap truncates.
    const capped = getSessionRange("range-cap", "Sheet1!A1:Z1000", {
      maxCells: 1,
      recalc: false,
      includeFormulas: false,
    });
    assert.equal(capped.cells.length, 1);
    assert.equal(capped.truncated, true);
  } finally {
    closeSession("range-cap");
  }
});

test("get_cell honors the sheet argument for a bare A1 reference", async () => {
  await openSession(undefined, "sheet-fallback");
  try {
    applySheetOperation("sheet-fallback", "add", { name: "Data" });
    applySessionMutations(
      "sheet-fallback",
      [
        { type: "text", a1: "Sheet1!A1", value: "on-Sheet1" },
        { type: "text", a1: "Data!A1", value: "on-Data" },
      ],
      true,
    );
    // Bare "A1" with sheet:1 must read the Data sheet, not sheet 0.
    assert.deepEqual(getSessionCell("sheet-fallback", 1, 0, 0).value, {
      kind: "text",
      value: "on-Data",
    });
  } finally {
    closeSession("sheet-fallback");
  }
});

test("set_range writes a 2D block and reports formula error cells", async () => {
  await openSession(undefined, "setrange");
  try {
    const result = setSessionRange(
      "setrange",
      "Sheet1!A1",
      [
        ["Item", "Qty", "Price"],
        ["Apple", 3, 1.5],
        ["Total", null, { f: "=B2*C2" }],
        [{ f: "=BADFN(1)" }, null, null],
      ],
      undefined,
      true,
    );
    assert.equal(result.cellsWritten, 9);
    assert.equal(result.range.start, "A1");
    assert.equal(result.range.end, "C4");
    assert.deepEqual(getSessionCellByA1("setrange", "Sheet1!C3").value, {
      kind: "number",
      value: 4.5,
    });
    assert.deepEqual(getSessionCellByA1("setrange", "Sheet1!B2").value, {
      kind: "number",
      value: 3,
    });
    assert.equal(result.errorCells.length, 1);
    assert.equal(result.errorCells[0].ref, "Sheet1!A4");
    assert.equal(result.errorCells[0].errorName, "#NAME?");
  } finally {
    closeSession("setrange");
  }
});

test("rejects writes past Excel grid limits", async () => {
  await openSession(undefined, "bounds");
  try {
    assert.throws(
      () =>
        applySessionMutations(
          "bounds",
          [{ type: "number", row: 2_000_000, col: 0, value: 1 }],
          false,
        ),
      /exceeds Excel sheet limits/,
    );
  } finally {
    closeSession("bounds");
  }
});

test("defined-name refersTo is stored without a leading '='", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-defname-"));
  const file = path.join(dir, "defname.xlsx");
  await openSession(undefined, "defname");
  try {
    setSessionDefinedName("defname", "MyRange", "=Sheet1!$A$1:$A$5");
    await saveSession("defname", file);
  } finally {
    closeSession("defname");
  }

  const wb = await loadWorkbook(file);
  try {
    const entry = wb.definedNameAt(0);
    assert.equal(entry.status.ok, true);
    assert.equal(entry.formula.startsWith("="), false);
    assert.equal(entry.formula, "Sheet1!$A$1:$A$5");
  } finally {
    wb.delete();
    await rm(dir, { recursive: true, force: true });
  }
});

/** Applies a custom number format to a cell via the low-level style APIs. */
function applyNumberFormat(id, sheet, row, col, formatCode) {
  const numFmtId = callWorkbookMethod(id, "addNumFmt", [formatCode]).result.numFmtId;
  const xfIndex = callWorkbookMethod(id, "addXf", [
    {
      fontIndex: 0,
      fillIndex: 0,
      borderIndex: 0,
      numFmtId,
      horizontalAlign: 0,
      verticalAlign: 0,
      wrapText: false,
    },
  ]).result.index;
  callWorkbookMethod(id, "setCellXfIndex", [sheet, row, col, xfIndex]);
}

test("decodes date, currency, and percent number formats on reads", async () => {
  await openSession(undefined, "numfmt");
  try {
    applySessionMutations(
      "numfmt",
      [
        { type: "number", row: 0, col: 0, value: 45000 },
        { type: "number", row: 1, col: 0, value: 1500 },
        { type: "number", row: 2, col: 0, value: 0.153 },
        { type: "number", row: 3, col: 0, value: 42 },
      ],
      false,
    );
    applyNumberFormat("numfmt", 0, 0, 0, "yyyy-mm-dd");
    applyNumberFormat("numfmt", 0, 1, 0, '"¥"#,##0');
    applyNumberFormat("numfmt", 0, 2, 0, "0.0%");

    const date = getSessionCell("numfmt", 0, 0, 0);
    assert.equal(date.formatKind, "date");
    assert.equal(date.formatted, "2023-03-15");
    assert.equal(date.value.value, 45000);

    const currency = getSessionCell("numfmt", 0, 1, 0);
    assert.equal(currency.formatKind, "currency");
    assert.equal(currency.formatted, undefined);

    const percent = getSessionCell("numfmt", 0, 2, 0);
    assert.equal(percent.formatKind, "percent");
    assert.equal(percent.formatted, "15.3%");

    // Unformatted number carries no format annotation.
    const plain = getSessionCell("numfmt", 0, 3, 0);
    assert.equal(plain.formatKind, undefined);
    assert.equal(plain.numberFormat, undefined);

    const range = getSessionRange("numfmt", "A1:A4", {
      maxCells: 100,
      recalc: false,
      includeFormulas: false,
    });
    const a1 = range.cells.find((cell) => cell.a1 === "A1");
    assert.equal(a1.formatted, "2023-03-15");
    const a4 = range.cells.find((cell) => cell.a1 === "A4");
    assert.equal(a4.formatKind, undefined);
  } finally {
    closeSession("numfmt");
  }
});

test("find_cells matches numeric cell values, not only text", async () => {
  await openSession(undefined, "findnum");
  try {
    applySessionMutations(
      "findnum",
      [
        { type: "number", row: 0, col: 0, value: 1500 },
        { type: "text", row: 1, col: 0, value: "budget" },
      ],
      false,
    );
    const found = findSessionCells("findnum", "1500", {
      target: "both",
      matchCase: false,
      wholeCell: false,
      regex: false,
      maxResults: 10,
    });
    assert.equal(found.count, 1);
    assert.equal(found.results[0].a1, "A1");
    assert.equal(found.results[0].text, "1500");
  } finally {
    closeSession("findnum");
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

test("dimension operation lists and sets column width and row height", async () => {
  await openSession(undefined, "dims");
  try {
    const setWidth = sessionDimension("dims", "column", "size", {
      sheet: 0,
      first: 0,
      last: 2,
      size: 120,
    });
    assert.equal(setWidth.status.ok, true);
    const setHeight = sessionDimension("dims", "row", "size", { sheet: 0, row: 0, size: 30 });
    assert.equal(setHeight.status.ok, true);
    const cols = sessionDimension("dims", "column", "list", { sheet: 0 });
    assert.equal(cols.value.status.ok, true);
    assert.throws(
      () => sessionDimension("dims", "column", "size", { sheet: 0, first: 0 }),
      /size is required/,
    );
  } finally {
    closeSession("dims");
  }
});

test("resolves sheet references by name for session operations", async () => {
  await openSession(undefined, "sheetname");
  try {
    applySheetOperation("sheetname", "rename", { index: 0, newName: "Data" });
    assert.equal(resolveSessionSheet("sheetname", "Data"), 0);
    assert.throws(() => resolveSessionSheet("sheetname", "Missing"), /sheet not found/);
  } finally {
    closeSession("sheetname");
  }
});

test("workbook_call marks pivot reads clean and rejects the removed save method", async () => {
  await openSession(undefined, "dirty");
  try {
    // A read-only allowlisted method must not flag the session dirty.
    const read = callWorkbookMethod("dirty", "getMerges", [0]);
    assert.equal(read.session.dirty, false);
    // A mutating method flags the session dirty.
    const write = callWorkbookMethod("dirty", "setNumber", [0, 0, 0, 1]);
    assert.equal(write.session.dirty, true);
    // Low-level `save` is no longer allowlisted; saveSession is the real path.
    assert.throws(() => callWorkbookMethod("dirty", "save", []), /not allowlisted/);
  } finally {
    closeSession("dirty");
  }
});
