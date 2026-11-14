import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { afterAll, beforeAll, test } from "vitest";
import { formulonModule, loadedFormulon, loadWorkbook } from "../dist/formulon.js";
import { getSession } from "../dist/session/state.js";
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
} from "../dist/sessions.js";
import { getWorkbookApiCatalog } from "../dist/workbook-api.js";
import { buildCallFixture, SHEET } from "./helpers/call-fixture.js";

const FORMULON_VERSION = createRequire(import.meta.url)("@libraz/formulon/package.json").version;

test("evaluates a formula through Formulon", async () => {
  const module = await formulonModule();
  assert.equal(module.getVersion(), FORMULON_VERSION);
  assert.equal(module.evaluateFormula("=SUM(1,2,3)"), 6);
  const typed = module.evaluateTypedFormula("=SUM(1,2,3)");
  assert.equal(typed.kind, module.ValueKind.Number);
  assert.equal(typed.number, 6);
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
    wb.recalc();
    assert.equal(wb.sheet(0).getValue(0, 0), 30);
    assert.equal(wb.sheet(0).getTypedValue(0, 0).kind, (await formulonModule()).ValueKind.Number);
  } finally {
    wb.dispose();
    await rm(dir, { recursive: true, force: true });
  }
});

test("names error codes outside the classic seven", async () => {
  await openSession(undefined, "error-names");
  try {
    assert.deepEqual(callWorkbookMethod("error-names", SHEET, "addMerge", ["A1:B1"]).result, {
      firstRow: 0,
      firstCol: 0,
      lastRow: 0,
      lastCol: 1,
    });
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

/**
 * Every function member of an engine handle, own and inherited, without `Object.prototype`'s.
 * The engine module is a plain object: its members are its own keys, less the capitalised
 * classes (`Workbook`, `ExcelError`, ...), which are not functions to call.
 */
function functionMembers(handle) {
  const names = new Set();
  if (Object.getPrototypeOf(handle) === Object.prototype) {
    for (const name of Object.keys(handle)) {
      if (typeof handle[name] === "function" && /^[a-z]/.test(name)) {
        names.add(name);
      }
    }
    return names;
  }
  for (
    let proto = Object.getPrototypeOf(handle);
    proto && proto !== Object.prototype;
    proto = Object.getPrototypeOf(proto)
  ) {
    for (const name of Object.getOwnPropertyNames(proto)) {
      if (typeof handle[name] === "function") {
        names.add(name);
      }
    }
  }
  return names;
}

/**
 * Walks a catalog kind's hops from `workbook`, the way the dispatcher does, with the fixture's
 * keys; the module kind is the engine module itself.
 */
function resolveReceiver(workbook, kind, target) {
  if (kind.kind === "Formulon") {
    return loadedFormulon();
  }
  let receiver = workbook;
  for (const hop of kind.hops) {
    receiver = "property" in hop ? receiver[hop.property] : receiver[hop.method](target[hop.key]);
  }
  return receiver;
}

/**
 * Members the catalog leaves out on purpose: the session owns closing and saving, the iterative
 * hook takes a JS callback, and snapshots, restores, image bytes and the row iterator cannot travel as JSON.
 */
const UNEXPOSED_MEMBERS = new Set([
  "constructor",
  "dispose",
  "save",
  "saveWithDiagnostics",
  "setIterativeProgress",
  "snapshotRangeXfs",
  "restoreRangeXfs",
  "remapRangeXfs",
  "snapshot",
  "iterRows",
  "addPicture",
  "restoreDrawing",
  // Theme files are bytes; formulon_theme_operation reads and writes them.
  "applyTheme",
  "saveTheme",
  // The engine module's image probe takes bytes and its log sink a callback.
  "getImageInfo",
  "setLogSink",
  // Package-level exports the runtime module also carries but the `Formulon` interface does not
  // declare, so the catalog cannot see them.
  "toNative",
  "enumName",
  "mergeFunctionMetadata",
]);

let fixtureDir;
let fixtureTargets;

beforeAll(async () => {
  fixtureDir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-engine-"));
  await openSession(undefined, "engine-fixture");
  try {
    fixtureTargets = await buildCallFixture(
      "engine-fixture",
      fixtureDir,
      path.join(fixtureDir, "fixture.xlsx"),
    );
  } finally {
    closeSession("engine-fixture");
  }
});

afterAll(async () => {
  await rm(fixtureDir, { recursive: true, force: true });
});

test("keeps the method catalog callable on the engine", async () => {
  const catalog = getWorkbookApiCatalog();
  await openSession(path.join(fixtureDir, "fixture.xlsx"), "engine-callable");
  try {
    const { workbook } = getSession("engine-callable");
    const missing = [];
    for (const kind of catalog.kinds) {
      const receiver = resolveReceiver(workbook, kind, fixtureTargets[kind.kind]);
      for (const method of catalog.methods) {
        if (method.target.kind === kind.kind && typeof receiver[method.name] !== "function") {
          missing.push(`${kind.kind}.${method.name}`);
        }
      }
    }
    assert.deepEqual(missing, []);
  } finally {
    closeSession("engine-callable");
  }
});

test("exposes every engine handle method that is not deliberately withheld", async () => {
  // The complementary direction of the guard above: that test catches a renamed or dropped
  // method, this one catches an engine release adding methods that silently stay unreachable
  // through formulon_workbook_call.
  const catalog = getWorkbookApiCatalog();
  await openSession(path.join(fixtureDir, "fixture.xlsx"), "engine-reachable");
  try {
    const { workbook } = getSession("engine-reachable");
    const unreachable = [];
    for (const kind of catalog.kinds) {
      // `Sheet.drawing(id)` returns the object's own class, so a Drawing receiver has its members.
      if (kind.kind === "Drawing") {
        continue;
      }
      const receiver = resolveReceiver(workbook, kind, fixtureTargets[kind.kind]);
      const cataloged = new Set(
        catalog.methods.filter((m) => m.target.kind === kind.kind).map((m) => m.name),
      );
      for (const name of functionMembers(receiver)) {
        if (!(cataloged.has(name) || UNEXPOSED_MEMBERS.has(name))) {
          unreachable.push(`${kind.kind}.${name}`);
        }
      }
    }
    assert.deepEqual(unreachable, []);
  } finally {
    closeSession("engine-reachable");
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
    callWorkbookMethod("clock", undefined, "setPinnedNow", [
      { year: 2026, month: 8, day: 18, hour: 9, minute: 30, second: 0 },
    ]);
    assert.deepEqual(callWorkbookMethod("clock", undefined, "getPinnedNow", []).result, {
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

    callWorkbookMethod("clock", undefined, "setPinnedNow", [null]);
    assert.equal(callWorkbookMethod("clock", undefined, "getPinnedNow", []).result, null);
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
    const call = (target, method, args = []) =>
      callWorkbookMethod("tables", target, method, args).result;
    const table = { kind: "Table", table: "Items" };
    const created = call(SHEET, "addTable", [
      { range: "A1:B2", name: "Items", columns: ["item", "qty"] },
    ]);
    assert.equal(created.handle, "Table");
    assert.equal(created.properties.name, "Items");
    assert.equal(call(SHEET, "listTables").length, 1);
    assert.deepEqual(call(table, "getProperties").range, {
      sheet: 0,
      rect: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
    });

    call(table, "update", [{ range: "A1:B3" }]);
    assert.equal(call(table, "getProperties").range.rect.lastRow, 2);
    call(table, "insertRows");
    call(table, "insertColumns");
    const grown = call(table, "getProperties");
    assert.deepEqual(grown.range.rect, { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 2 });
    assert.deepEqual(
      grown.columns.map((column) => column.name),
      ["item", "qty", "Column1"],
    );
    call(SHEET, "removeTable", ["Items"]);
    assert.equal(call(SHEET, "listTables").length, 0);
  } finally {
    closeSession("tables");
  }
});

test("round-trips a cell phonetic guide", async () => {
  await openSession(undefined, "phonetic");
  try {
    applySessionMutations("phonetic", [{ type: "text", a1: "Sheet1!A1", value: "山田" }], false);
    callWorkbookMethod("phonetic", SHEET, "setCellPhonetic", [0, 0, "ヤマダ"]);
    assert.equal(callWorkbookMethod("phonetic", SHEET, "getCellPhonetic", [0, 0]).result, "ヤマダ");
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
    const call = (method, args) => callWorkbookMethod("phonetic-runs", SHEET, method, args).result;
    const runs = [
      { start: 0, end: 2, text: "ヤマダ" },
      { start: 2, end: 4, text: "タロウ" },
    ];
    call("setCellPhoneticRuns", [0, 0, runs]);
    assert.deepEqual(call("getCellPhoneticRuns", [0, 0]), runs);
    // The flattening read concatenates the readings rather than reporting spans.
    assert.equal(call("getCellPhonetic", [0, 0]), "ヤマダタロウ");

    // Rendering is set apart from the readings, and setting it leaves them be.
    call("setCellPhoneticProperties", [0, 0, { fontIndex: 0, type: 2, alignment: 1 }]);
    const properties = call("getCellPhoneticProperties", [0, 0]);
    assert.equal(properties.type, 2);
    assert.equal(properties.alignment, 1);
    assert.equal(properties.stated, true);
    assert.deepEqual(call("getCellPhoneticRuns", [0, 0]), runs);
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

    const metadata = await getSessionMetadata("helpers", "functions");
    assert.equal(Array.isArray(metadata.value), true);
    assert.equal(metadata.value.includes("SUM"), true);
  } finally {
    closeSession("helpers");
  }
});

test("supports catalog-listed workbook calls and rejects unknown methods", async () => {
  await openSession(undefined, "low-level");
  try {
    const sheets = callWorkbookMethod("low-level", undefined, "listSheets", []);
    assert.equal(sheets.result.length, 1);
    assert.equal(sheets.method, "listSheets");
    assert.deepEqual(sheets.target, { kind: "Workbook" });

    const merge = callWorkbookMethod("low-level", SHEET, "addMerge", ["A1:B1"]);
    assert.deepEqual(merge.result, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 });
    assert.equal(merge.session.dirty, true);

    const merges = callWorkbookMethod("low-level", SHEET, "listMerges", []);
    assert.deepEqual(merges.result, [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 }]);

    assert.throws(
      () => callWorkbookMethod("low-level", undefined, "constructor", []),
      /not a callable method/,
    );
    // A method of another kind is not callable on the workbook.
    assert.throws(
      () => callWorkbookMethod("low-level", undefined, "addDataField", []),
      /Workbook\.addDataField is not a callable method/,
    );
  } finally {
    closeSession("low-level");
  }
});
