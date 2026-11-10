import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { afterAll, beforeAll, test } from "vitest";
import { formulonModule, resultToJson } from "../../dist/formulon.js";
import { callLabeled } from "../../dist/session/workbook-call.js";
import {
  callWorkbookMethod,
  closeSession,
  insertSessionImage,
  listSessions,
  openSession,
  saveSession,
} from "../../dist/sessions.js";
import { getWorkbookApiCatalog, methodAccess } from "../../dist/workbook-api.js";
import { buildCallFixture, SHEET, SHEET_OOXML, WORKBOOK_OOXML } from "../helpers/call-fixture.js";
import { solidPng } from "../helpers/png.js";

const dirtyOf = (id) => listSessions().find((session) => session.id === id)?.dirty;

let fixtureDir;
let fixtureFile;
/** One target per catalog kind, addressing real objects of the fixture workbook. */
let targets;

beforeAll(async () => {
  fixtureDir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-call-"));
  fixtureFile = path.join(fixtureDir, "fixture.xlsx");
  await openSession(undefined, "call-fixture");
  try {
    targets = await buildCallFixture("call-fixture", fixtureDir, fixtureFile);
  } finally {
    closeSession("call-fixture");
  }
});

afterAll(async () => {
  await rm(fixtureDir, { recursive: true, force: true });
});

/** Opens the fixture workbook as a clean session, runs `fn`, then closes the session. */
async function withFixture(id, fn) {
  await openSession(fixtureFile, id);
  try {
    assert.equal(dirtyOf(id), false);
    return await fn();
  } finally {
    closeSession(id);
  }
}

/** Opens a clean empty workbook, runs `fn(call)` and closes it. */
async function withCleanSession(id, prepare, fn) {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-call-clean-"));
  await openSession(undefined, id);
  try {
    const call = (target, method, args = []) => callWorkbookMethod(id, target, method, args);
    await prepare?.(call, dir);
    await saveSession(id, path.join(dir, "clean.xlsx"));
    assert.equal(dirtyOf(id), false);
    await fn(call, dir);
  } finally {
    closeSession(id);
    await rm(dir, { recursive: true, force: true });
  }
}

/** A cell xf that differs from the default only by number format 10 (0.00%). */
function percentXf(call) {
  const record = call(WORKBOOK_OOXML, "getCellXf", [0]).result;
  return call(WORKBOOK_OOXML, "addCellXf", [{ ...record, numFmtId: 10, applyNumberFormat: true }])
    .result;
}

test("classifies every cataloged method as read or write", () => {
  const catalog = getWorkbookApiCatalog();
  assert.ok(catalog.methods.length > 400);
  const seen = new Set();
  for (const method of catalog.methods) {
    const label = `${method.target.kind}.${method.name}`;
    assert.ok(method.access === "read" || method.access === "write", `${label}: ${method.access}`);
    assert.equal(seen.has(label), false, `${label} is cataloged twice`);
    seen.add(label);
  }
  // Every kind the sweeps below address is in the catalog, and every cataloged kind has a target.
  assert.deepEqual(catalog.kinds.map((kind) => kind.kind).sort(), Object.keys(targets).sort());
});

test("the engine module's functions never mark a session dirty", async () => {
  const catalog = getWorkbookApiCatalog();
  const module = catalog.methods.filter((method) => method.target.kind === "Formulon");
  assert.ok(module.length > 10);
  assert.deepEqual(
    module.filter((method) => method.access !== "read").map((method) => method.name),
    [],
  );
  await openSession(undefined, "module-clean");
  try {
    const target = { kind: "Formulon" };
    // Verbs that would read as writes elsewhere: a conversion and process-level logging.
    callWorkbookMethod("module-clean", target, "convertCivilToSerial", [
      { year: 2024, month: 1, day: 2 },
    ]);
    // Off is the engine default, so the process-wide level is left as it was.
    callWorkbookMethod("module-clean", target, "setLogMinLevel", [4]);
    assert.equal(dirtyOf("module-clean"), false);
  } finally {
    closeSession("module-clean");
  }
});

test("methodAccess leaves a method neither the ledger nor a prefix names undecided", () => {
  assert.equal(methodAccess("someUnknownVerb", new Set()), undefined);
  assert.equal(methodAccess("frobnicateAll", new Set(["frobnicateAll"])), "read");
  assert.equal(methodAccess("getSomething", new Set()), "read");
  assert.equal(methodAccess("setSomething", new Set()), "write");
  // The prefix must end at a word boundary: `settle` is not `set`.
  assert.equal(methodAccess("settleAccounts", new Set()), undefined);
});

test("every read method leaves the session clean", async () => {
  const catalog = getWorkbookApiCatalog();
  const reads = catalog.methods.filter((method) => method.access === "read");
  assert.ok(reads.length > 150);
  await withFixture("read-sweep", () => {
    for (const method of reads) {
      // A call that fails on its arguments must not mark the session dirty either.
      try {
        callWorkbookMethod("read-sweep", targets[method.target.kind], method.name, []);
      } catch {
        // The classification is under test, not the arguments.
      }
      assert.equal(dirtyOf("read-sweep"), false, `${method.target.kind}.${method.name} dirtied`);
    }
  });
});

test("every write method marks the session dirty, even when the call throws", async () => {
  const catalog = getWorkbookApiCatalog();
  const writes = catalog.methods.filter((method) => method.access === "write");
  assert.ok(writes.length > 200);
  for (const required of [
    ["SheetOoxml", "setColumnXfIndex"],
    ["SheetOoxml", "resetColumnXfIndex"],
    ["SheetOoxml", "setRowXfIndex"],
    ["SheetOoxml", "resetRowXfIndex"],
    ["SheetOoxml", "resetCellXfIndex"],
    ["Sheet", "resetColumnWidth"],
    ["Workbook", "setThemeScriptFont"],
    ["Workbook", "resetTheme"],
    ["Picture", "update"],
    ["Picture", "setZOrder"],
  ]) {
    assert.ok(
      writes.some((method) => method.target.kind === required[0] && method.name === required[1]),
      `${required.join(".")} should be a write`,
    );
  }
  for (const [index, method] of writes.entries()) {
    const id = `write-sweep-${index}`;
    await withFixture(id, () => {
      try {
        callWorkbookMethod(id, targets[method.target.kind], method.name, []);
      } catch {
        // A failed or malformed mutation stays dirty: it may have changed native state first.
      }
      assert.equal(dirtyOf(id), true, `${method.target.kind}.${method.name} left it clean`);
    });
  }
}, 180_000);

test("refuses session-owned and uncallable members on every kind", async () => {
  const catalog = getWorkbookApiCatalog();
  const refused = [
    "save",
    "saveWithDiagnostics",
    "dispose",
    "snapshotRangeXfs",
    "restoreRangeXfs",
    "constructor",
    "__proto__",
  ];
  for (const name of refused) {
    assert.equal(
      catalog.methods.some((method) => method.name === name),
      false,
      `${name} is cataloged`,
    );
  }
  await withFixture("refused", () => {
    for (const kind of catalog.kinds) {
      for (const name of refused) {
        assert.throws(
          () => callWorkbookMethod("refused", targets[kind.kind], name, []),
          /is not a callable method/,
          `${kind.kind}.${name} was dispatched`,
        );
      }
    }
    // A method of another kind is refused on this one.
    assert.throws(
      () => callWorkbookMethod("refused", SHEET, "addSheet", ["x"]),
      /Sheet\.addSheet is not a callable method/,
    );
    assert.throws(
      () => callWorkbookMethod("refused", { kind: "Nope" }, "getValue", []),
      /unknown target kind: Nope/,
    );
    assert.equal(dirtyOf("refused"), false);
  });
});

test("refuses setIterativeProgress, whose callback no JSON argument can carry", async () => {
  assert.equal(
    getWorkbookApiCatalog().methods.some((method) => method.name === "setIterativeProgress"),
    false,
    "setIterativeProgress is cataloged",
  );
  await openSession(undefined, "iterative-progress");
  try {
    assert.throws(
      () => callWorkbookMethod("iterative-progress", undefined, "setIterativeProgress", [null]),
      /is not a callable method/,
    );
  } finally {
    closeSession("iterative-progress");
  }
});

test("catalogs the handle kinds with the keys that reach them", () => {
  const kinds = new Map(getWorkbookApiCatalog().kinds.map((kind) => [kind.kind, kind.keys]));
  assert.deepEqual(kinds.get("Workbook"), []);
  assert.deepEqual(kinds.get("Sheet"), ["sheet"]);
  assert.deepEqual(kinds.get("Pivot"), ["sheet", "pivot"]);
  assert.deepEqual(kinds.get("Chart"), ["sheet", "chart"]);
  assert.deepEqual(kinds.get("Table"), ["table"]);
  assert.deepEqual(kinds.get("WorkbookOoxml"), []);
  assert.deepEqual(kinds.get("SheetOoxml"), ["sheet"]);
});

test("resolves a target through its keys and describes handles in results", async () => {
  await withFixture("targets", () => {
    const called = callWorkbookMethod("targets", targets.Pivot, "getProperties", []);
    assert.equal(called.result.name, "Sales");
    assert.deepEqual(called.target, targets.Pivot);
    assert.equal(called.method, "getProperties");
    assert.equal(called.session.dirty, false);

    const sheets = callWorkbookMethod("targets", undefined, "listSheets", []).result;
    assert.deepEqual(sheets, [{ handle: "Sheet", index: 0, name: "Sheet1" }]);
    const table = callWorkbookMethod("targets", SHEET, "listTables", []).result;
    assert.equal(table[0].handle, "Table");
    assert.equal(table[0].properties.name, "Items");
    const drawing = callWorkbookMethod("targets", targets.Chart, "getDrawingProperties", []);
    assert.equal(drawing.result.objectId, targets.Chart.chart);

    // A missing key and an object that does not exist are both errors, and neither dirties.
    assert.throws(
      () => callWorkbookMethod("targets", { kind: "Pivot", sheet: 0 }, "getProperties", []),
      /target Pivot needs the key pivot/,
    );
    assert.throws(
      () => callWorkbookMethod("targets", { kind: "Sheet", sheet: 99 }, "listMerges", []),
      /out of range/,
    );
    assert.equal(dirtyOf("targets"), false);
  });
});

test("throws engine errors for failed calls and keeps failed mutations dirty", async () => {
  await openSession(undefined, "failed-call");
  try {
    assert.throws(
      () => callWorkbookMethod("failed-call", SHEET, "setNumber", [-1, 0, 1]),
      /outside the grid/,
    );
    assert.equal(dirtyOf("failed-call"), true);
  } finally {
    closeSession("failed-call");
  }
  await openSession(undefined, "failed-read");
  try {
    assert.throws(
      () => callWorkbookMethod("failed-read", { kind: "Sheet", sheet: 99 }, "listMerges", []),
      /out of range/,
    );
    assert.equal(dirtyOf("failed-read"), false);
  } finally {
    closeSession("failed-read");
  }
});

test("supports empty passthrough and external-link metadata on fresh workbooks", async () => {
  await openSession(undefined, "metadata-empty");
  try {
    const call = (target, method) => callWorkbookMethod("metadata-empty", target, method, []);
    assert.deepEqual(call(WORKBOOK_OOXML, "listPassthroughs").result, []);
    assert.deepEqual(call(undefined, "listExternalLinks").result, []);
    assert.ok(call(undefined, "listCellStyles").result.length >= 1);
    assert.ok(call(WORKBOOK_OOXML, "listCellStyleXfs").result.length >= 1);
    assert.equal(dirtyOf("metadata-empty"), false);
  } finally {
    closeSession("metadata-empty");
  }
});

test("workbook_call marks reads clean, writes dirty, and rejects the session-owned save", async () => {
  await openSession(undefined, "dirty");
  try {
    // A read must not flag the session dirty.
    const read = callWorkbookMethod("dirty", SHEET, "listMerges", []);
    assert.equal(read.session.dirty, false);
    // A mutating method flags the session dirty.
    const write = callWorkbookMethod("dirty", SHEET, "setNumber", [0, 0, 1]);
    assert.equal(write.session.dirty, true);
    // `save` is owned by the session; saveSession is the real path.
    assert.throws(
      () => callWorkbookMethod("dirty", undefined, "save", []),
      /not a callable method/,
    );
  } finally {
    closeSession("dirty");
  }
});

test("result conversion returns typed collation class ids as a plain array and stays clean", async () => {
  const engine = await formulonModule();
  await openSession(undefined, "collation");
  try {
    const { result, session } = callLabeled(
      "collation",
      "Formulon.getCollationClasses",
      "read",
      () =>
        engine.getCollationClasses(
          [
            { kind: 3, text: "a" },
            { kind: 3, text: "A" },
            { kind: 1, number: 1 },
          ],
          engine.CollationStrength.Compare,
          "mac-365-ja_JP",
        ),
    );
    assert.deepEqual(result.classIds, [1, 1, 0]);
    assert.equal(result.classCount, 2);
    assert.equal(session.dirty, false);
    assert.deepEqual(resultToJson(new Uint32Array([0, 1])), [0, 1]);
  } finally {
    closeSession("collation");
  }
});

test("a labeled write marks the session dirty once it succeeds and not when the engine refuses it", async () => {
  await openSession(undefined, "labeled-write");
  try {
    // One engine call changes nothing when it throws, so the session stays clean.
    assert.throws(
      () =>
        callLabeled("labeled-write", "Test.fail", "write", () => {
          throw new Error("boom");
        }),
      /boom/,
    );
    assert.equal(dirtyOf("labeled-write"), false);
    callLabeled("labeled-write", "Sheet.setNumber", "write", (wb) =>
      wb.sheet(0).setNumber(0, 0, 1),
    );
    assert.equal(dirtyOf("labeled-write"), true);
  } finally {
    closeSession("labeled-write");
  }
});

test("style reads and locale reads leave the session clean", async () => {
  const engine = await formulonModule();
  await withCleanSession("call-read-new", undefined, (call) => {
    const clean = (target, name, args) => {
      const called = call(target, name, args);
      assert.equal(called.session.dirty, false, `${name} dirtied the session`);
      return called.result;
    };
    assert.deepEqual(clean(SHEET_OOXML, "getRangeXfBases", ["A1:D4"]), [0]);
    assert.equal(clean(SHEET_OOXML, "getEffectiveStyle", [0, 0]).numberFormat, "General");
    assert.equal(clean(SHEET, "getCellFormat", [0, 0]).numberFormat, "General");

    const labeled = (method, operation) => {
      const called = callLabeled("call-read-new", method, "read", operation);
      assert.equal(called.session.dirty, false, `${method} dirtied the session`);
      return called.result;
    };
    assert.equal(
      labeled("Formulon.localizeFormula", () =>
        engine.localizeFormula("=SUM(1.5)", "win-365-de_DE"),
      ),
      "=SUMME(1,5)",
    );
    assert.equal(
      labeled("Formulon.canonicalizeFormula", () =>
        engine.canonicalizeFormula("=SUMME(1,5)", "win-365-de_DE"),
      ),
      "=SUM(1.5)",
    );
    assert.equal(
      labeled("Formulon.getLocaleFacts", () => engine.getLocaleFacts("win-365-de_DE"))
        .decimalSeparator,
      ",",
    );
  });
});

test("column and row styles are set and reset as mutations", async () => {
  let xf;
  await withCleanSession(
    "call-line-styles",
    (call) => {
      xf = percentXf(call);
    },
    async (call, dir) => {
      const columns = () => call(SHEET, "listColumnOverrides").result;
      const rows = () => call(SHEET, "listRowOverrides").result;
      const columnXfs = () =>
        [0, 1, 2, 3, 4].map((col) => call(SHEET_OOXML, "getColumnXfIndex", [col]).result);
      const rowXfs = () =>
        [3, 4, 5, 6, 7].map((row) => call(SHEET_OOXML, "getRowXfIndex", [row]).result);
      const numberFormat = (row, col) =>
        call(SHEET_OOXML, "getEffectiveStyle", [row, col]).result.numberFormat;

      assert.equal(call(SHEET_OOXML, "setColumnXfIndex", [1, 3, xf]).session.dirty, true);
      assert.deepEqual(
        columns().map(({ first, last }) => ({ first, last })),
        [{ first: 1, last: 3 }],
      );
      assert.deepEqual(columnXfs(), [null, xf, xf, xf, null]);
      assert.equal(numberFormat(50, 2), "0.00%");
      await saveSession("call-line-styles", path.join(dir, "after-set.xlsx"));
      assert.equal(dirtyOf("call-line-styles"), false);
      assert.equal(call(SHEET_OOXML, "resetColumnXfIndex", [1, 3]).session.dirty, true);
      assert.deepEqual(columnXfs(), [null, null, null, null, null]);
      assert.equal(numberFormat(50, 2), "General");

      await saveSession("call-line-styles", path.join(dir, "after-clear.xlsx"));
      assert.equal(call(SHEET_OOXML, "setRowXfIndex", [4, 6, xf]).session.dirty, true);
      assert.deepEqual(
        rows().map(({ row }) => row),
        [4, 5, 6],
      );
      assert.deepEqual(rowXfs(), [null, xf, xf, xf, null]);
      assert.equal(numberFormat(5, 40), "0.00%");
      await saveSession("call-line-styles", path.join(dir, "after-row.xlsx"));
      assert.equal(call(SHEET_OOXML, "resetRowXfIndex", [4, 6]).session.dirty, true);
      assert.deepEqual(rowXfs(), [null, null, null, null, null]);
      assert.deepEqual(rows(), []);
    },
  );
});

test("resetting a column width keeps what else the column carries", async () => {
  let xf;
  await withCleanSession(
    "call-clear-width",
    (call) => {
      xf = percentXf(call);
      call(SHEET, "setColumnWidth", [1, 3, 20]);
      call(SHEET, "setColumnHidden", [2, 2, true]);
      call(SHEET_OOXML, "setColumnXfIndex", [1, 3, xf]);
    },
    (call) => {
      const columns = () => call(SHEET, "listColumnOverrides").result;
      assert.ok(columns().every((column) => column.widthChars === 20));
      const cleared = call(SHEET, "resetColumnWidth", [1, 3]);
      assert.equal(cleared.session.dirty, true);
      const after = columns();
      assert.ok(after.length >= 1);
      for (const column of after) {
        assert.equal(column.widthChars, null);
      }
      assert.equal(
        after.some((column) => column.first <= 2 && 2 <= column.last && column.hidden === true),
        true,
      );
      for (const col of [1, 2, 3]) {
        assert.equal(call(SHEET_OOXML, "getColumnXfIndex", [col]).result, xf);
      }
      // Resetting a column that has no width is a successful no-op.
      assert.equal(call(SHEET, "resetColumnWidth", [10, 12]).session.dirty, true);
      assert.equal(columns().length, after.length);
    },
  );
});

test("resetting a cell xf returns the cell to the default style and keeps its value", async () => {
  let xf;
  await withCleanSession(
    "call-clear-xf",
    (call) => {
      xf = percentXf(call);
      call(SHEET, "setNumber", [0, 0, 5]);
      call(SHEET_OOXML, "setCellXfIndex", [0, 0, xf]);
    },
    (call) => {
      assert.equal(call(SHEET_OOXML, "getCellXfIndex", [0, 0]).result, xf);
      const cleared = call(SHEET_OOXML, "resetCellXfIndex", [0, 0]);
      assert.equal(cleared.session.dirty, true);
      assert.equal(call(SHEET_OOXML, "getCellXfIndex", [0, 0]).result, 0);
      assert.equal(call(SHEET_OOXML, "getEffectiveStyle", [0, 0]).result.numberFormat, "General");
      assert.equal(call(SHEET, "getValue", [0, 0]).result, 5);
    },
  );
});

test("remapRangeXfs, whose Map argument JSON cannot carry, is refused before it runs", async () => {
  let xf;
  await withCleanSession(
    "call-remap",
    (call) => {
      xf = percentXf(call);
    },
    (call) => {
      assert.deepEqual(call(SHEET_OOXML, "getRangeXfBases", ["A1:B2"]).result, [0]);
      // The engine takes a Map; a JSON argument is never one, so the method is not callable.
      for (const mapping of [{ 0: xf }, [[0, xf]]]) {
        assert.throws(
          () => callWorkbookMethod("call-remap", SHEET_OOXML, "remapRangeXfs", ["A1:B2", mapping]),
          /is not a callable method/,
        );
      }
      assert.equal(call(SHEET_OOXML, "getCellXfIndex", [1, 1]).result, null);
      assert.equal(call(SHEET_OOXML, "getEffectiveStyle", [1, 1]).result.numberFormat, "General");
      // A refused call never ran, so the session stays clean.
      assert.equal(dirtyOf("call-remap"), false);
    },
  );
});

test("theme script fonts are set and reset as mutations", async () => {
  await withCleanSession("call-theme", undefined, async (call, dir) => {
    const MAJOR = 1;
    const MINOR = 2;
    const jpan = (scheme) =>
      call(undefined, "listThemeScriptFonts", [scheme]).result.find(
        (entry) => entry.script === "Jpan",
      ).typeface;
    const original = { major: jpan(MAJOR), minor: jpan(MINOR) };

    const set = call(undefined, "setThemeScriptFont", [
      { scheme: MAJOR, script: "Jpan", typeface: "Meiryo" },
    ]);
    assert.equal(set.session.dirty, true);
    assert.equal(jpan(MAJOR), "Meiryo");
    assert.equal(jpan(MINOR), original.minor, "the minor set is untouched");
    assert.equal(call(undefined, "getThemeFonts").result.majorEastAsian, "Meiryo");

    await saveSession("call-theme", path.join(dir, "themed.xlsx"));
    assert.equal(
      call(undefined, "setThemeScriptFont", [
        { scheme: MINOR, script: "Jpan", typeface: "Yu Mincho" },
      ]).session.dirty,
      true,
    );
    assert.equal(jpan(MINOR), "Yu Mincho");
    assert.throws(
      () =>
        callWorkbookMethod("call-theme", undefined, "setThemeScriptFont", [
          { scheme: MAJOR, script: "Jpan", typeface: "" },
        ]),
      /typeface must not be empty/,
    );

    await saveSession("call-theme", path.join(dir, "themed-again.xlsx"));
    const reset = call(undefined, "resetTheme");
    assert.equal(reset.session.dirty, true);
    assert.deepEqual({ major: jpan(MAJOR), minor: jpan(MINOR) }, original);
  });
});

test("image anchor and stacking methods are mutations", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-call-images-"));
  const wide = path.join(dir, "wide.png");
  const narrow = path.join(dir, "narrow.png");
  await writeFile(wide, solidPng(30, 10));
  await writeFile(narrow, solidPng(20, 10));
  await openSession(undefined, "call-images");
  try {
    const first = (await insertSessionImage("call-images", 0, { path: narrow, anchor: "B2" }))
      .objectId;
    await insertSessionImage("call-images", 0, { path: wide, anchor: "D4" });
    await saveSession("call-images", path.join(dir, "images.xlsx"));
    assert.equal(dirtyOf("call-images"), false);

    const call = (target, method, args = []) =>
      callWorkbookMethod("call-images", target, method, args);
    const picture = { kind: "Picture", sheet: 0, picture: first };
    // A listed picture is described by its properties; the widths tell the two apart.
    const stacking = () =>
      call(SHEET, "listDrawings").result.map((item) => item.properties.widthPx);
    const anchorOf = () => call(picture, "getDrawingProperties").result.anchor;
    const listed = call(SHEET, "listDrawings");
    assert.equal(listed.session.dirty, false);
    assert.deepEqual(stacking(), [20, 30]);
    assert.deepEqual({ row: anchorOf().fromRow, col: anchorOf().fromCol }, { row: 1, col: 1 });

    const anchored = call(picture, "update", [{ anchor: { fromRow: 7, fromCol: 5 } }]);
    assert.equal(anchored.session.dirty, true);
    assert.deepEqual({ row: anchorOf().fromRow, col: anchorOf().fromCol }, { row: 7, col: 5 });

    await saveSession("call-images", path.join(dir, "images-2.xlsx"));
    const stacked = call(picture, "setZOrder", [1]);
    assert.equal(stacked.session.dirty, true);
    assert.deepEqual(stacking(), [30, 20]);
  } finally {
    closeSession("call-images");
    await rm(dir, { recursive: true, force: true });
  }
});

const COMPARE_CASES = [
  // [a, b, order] under a mac profile, where text compares case-insensitively.
  ["a", "A", 0],
  ["A", "a", 0],
  ["a", "a", 0],
  ["", "", 0],
  ["a", "b", -1],
  ["b", "a", 1],
  ["A", "b", -1],
  ["b", "A", 1],
  ["a", "", 1],
  ["", "a", -1],
  ["abc", "abd", -1],
  ["abc", "ab", 1],
];

test("compareText orders text case-insensitively under a mac profile at every strength", async () => {
  const engine = await formulonModule();
  await openSession(undefined, "compare-mac");
  try {
    const compare = (a, b, strength, profile) => {
      const called = callLabeled("compare-mac", "Formulon.compareText", "read", () =>
        engine.compareText(a, b, strength, profile),
      );
      assert.equal(called.session.dirty, false);
      return called.result;
    };
    for (const profile of ["mac-365-en_US", "mac-365-ja_JP", "mac-365-de_DE"]) {
      for (const strength of [0, 1, 2]) {
        for (const [a, b, order] of COMPARE_CASES) {
          assert.equal(
            compare(a, b, strength, profile),
            order,
            `${profile}/${strength}: ${a} vs ${b}`,
          );
        }
      }
    }
    // Strength picks the fold: Mac and Windows collation both equate sharp s
    // with ss except for exact classic matching.
    assert.equal(compare("ß", "ss", 0, "mac-365-en_US"), 0);
    assert.equal(compare("ß", "ss", 2, "mac-365-en_US"), 1);
    assert.equal(compare("ß", "ss", 0, "win-365-en_US"), 0);
    assert.equal(compare("ß", "ss", 2, "win-365-en_US"), 1);
    assert.throws(() => compare("a", "b", 0, "mac-365-xx_XX"), /./);
  } finally {
    closeSession("compare-mac");
  }
});
