import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { formulonModule, loadWorkbook } from "../dist/formulon.js";
import {
  applySessionMutations,
  callWorkbookMethod,
  closeSession,
  editSessionStructure,
  getSessionCell,
  getSessionMetadata,
  inspectSession,
  openSession,
  saveSession,
  setSessionDefinedName,
  setSessionRange,
  setSessionSheetView,
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
      status: { ok: true, status: 0, message: "", context: "" },
      now: {
        year: 2026,
        month: 8,
        day: 18,
        hour: 9,
        minute: 30,
        second: 0,
      },
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
    assert.deepEqual(callWorkbookMethod("clock", "pinnedNow", []).result, {
      status: { ok: true, status: 0, message: "", context: "" },
      now: null,
    });
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
    assert.equal(callWorkbookMethod("tables", "tableCount", []).result.value, 1);
    assert.equal(callWorkbookMethod("tables", "tableAt", [0]).result.ref, "A1:B2");

    assert.equal(
      callWorkbookMethod("tables", "updateTable", [0, { ref: "A1:B3" }]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("tables", "tableAt", [0]).result.ref, "A1:B3");
    assert.equal(callWorkbookMethod("tables", "removeTable", [0]).result.ok, true);
    assert.equal(callWorkbookMethod("tables", "tableCount", []).result.value, 0);
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

test("round-trips phonetic guides span by span", async () => {
  await openSession(undefined, "phonetic-runs");
  try {
    applySessionMutations(
      "phonetic-runs",
      [{ type: "text", a1: "Sheet1!A1", value: "山田太郎" }],
      false,
    );
    const runs = [
      { sb: 0, eb: 2, text: "ヤマダ" },
      { sb: 2, eb: 4, text: "タロウ" },
    ];
    assert.equal(
      callWorkbookMethod("phonetic-runs", "setCellPhoneticRuns", [0, 0, 0, runs]).result.ok,
      true,
    );
    const read = callWorkbookMethod("phonetic-runs", "getCellPhoneticRuns", [0, 0, 0]).result;
    assert.equal(read.status.ok, true);
    assert.deepEqual(read.runs, runs);
    // The flattening read concatenates the readings rather than reporting spans.
    assert.equal(
      callWorkbookMethod("phonetic-runs", "getCellPhonetic", [0, 0, 0]).result.value,
      "ヤマダタロウ",
    );

    // Rendering is set apart from the readings, and setting it leaves them be.
    assert.equal(
      callWorkbookMethod("phonetic-runs", "setCellPhoneticProperties", [
        0,
        0,
        0,
        { fontId: 0, type: 2, alignment: 1 },
      ]).result.ok,
      true,
    );
    const properties = callWorkbookMethod(
      "phonetic-runs",
      "getCellPhoneticProperties",
      [0, 0, 0],
    ).result;
    assert.equal(properties.type, 2);
    assert.equal(properties.alignment, 1);
    assert.deepEqual(
      callWorkbookMethod("phonetic-runs", "getCellPhoneticRuns", [0, 0, 0]).result.runs,
      runs,
    );
  } finally {
    closeSession("phonetic-runs");
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
    assert.equal(sheetCount.result.value, 1);

    const merge = callWorkbookMethod("low-level", "addMerge", [
      0,
      { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 },
    ]);
    assert.equal(merge.result.ok, true);

    const merges = callWorkbookMethod("low-level", "getMerges", [0]);
    assert.deepEqual(merges.result.items, [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 }]);

    assert.throws(() => callWorkbookMethod("low-level", "constructor", []), /not allowlisted/);
    assert.throws(
      () => callWorkbookMethod("low-level", "pivotFieldAddAggregation", [0, 0, 0, 0]),
      /not allowlisted/,
    );
  } finally {
    closeSession("low-level");
  }
});

/**
 * Engine methods deliberately kept out of `WORKBOOK_METHODS`. Saving is routed
 * through `formulon_save_session` so a session keeps its bookkeeping, the
 * iterative-progress hook takes a JS callback no JSON argument can express,
 * image insertion and probing take raw bytes and go through the image tool, and
 * the rest are handle lifecycle rather than workbook operations.
 */
const UNEXPOSED_WORKBOOK_METHODS = new Set([
  "constructor",
  "delete",
  "deleteLater",
  "insertImage",
  "isValid",
  "probeImage",
  "save",
  "saveAs",
  "saveWithDiagnostics",
  "setIterativeProgress",
]);

test("exposes every engine workbook method that is not deliberately withheld", async () => {
  const module = await formulonModule();
  const wb = module.Workbook.createDefault();
  try {
    // The complementary direction of the guard above: that test catches a
    // renamed or dropped method, this one catches an engine release adding
    // methods that silently stay unreachable through formulon_workbook_call.
    const unreachable = Object.getOwnPropertyNames(Object.getPrototypeOf(wb))
      .filter((name) => typeof wb[name] === "function")
      .filter((name) => !(WORKBOOK_METHODS.has(name) || UNEXPOSED_WORKBOOK_METHODS.has(name)));
    assert.deepEqual(unreachable, []);
  } finally {
    wb.delete();
  }
});
