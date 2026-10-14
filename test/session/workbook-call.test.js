import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { READ_ONLY_METHODS } from "../../dist/session/workbook-call.js";
import {
  callWorkbookMethod,
  closeSession,
  insertSessionImage,
  listSessions,
  openSession,
  saveSession,
  WORKBOOK_METHODS,
} from "../../dist/sessions.js";
import { solidPng } from "../helpers/png.js";

test("throws failed workbook statuses and keeps failed mutations dirty", async () => {
  await openSession(undefined, "failed-call");
  try {
    assert.throws(
      () => callWorkbookMethod("failed-call", "setNumber", [99, 0, 0, 1]),
      /workbook setNumber failed/,
    );
    assert.equal(listSessions().find((session) => session.id === "failed-call")?.dirty, true);
    assert.throws(
      () => callWorkbookMethod("failed-call", "getMerges", [99]),
      /workbook getMerges failed/,
    );
  } finally {
    closeSession("failed-call");
  }
});

test("supports empty passthrough and external-link metadata on fresh workbooks", async () => {
  await openSession(undefined, "metadata-empty");
  try {
    assert.equal(callWorkbookMethod("metadata-empty", "passthroughCount", []).result.value, 0);
    assert.deepEqual(callWorkbookMethod("metadata-empty", "getExternalLinks", []).result.items, []);
    assert.equal(
      callWorkbookMethod("metadata-empty", "cellStyleCount", []).result.value >= 0,
      true,
    );
    assert.equal(
      callWorkbookMethod("metadata-empty", "cellStyleXfCount", []).result.value >= 0,
      true,
    );
  } finally {
    closeSession("metadata-empty");
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

test("workbook_call returns collation class ids as a plain array and stays clean", async () => {
  await openSession(undefined, "collation");
  try {
    const { result, session } = callWorkbookMethod("collation", "collationClasses", [
      [
        { kind: 3, text: "a" },
        { kind: 3, text: "A" },
        { kind: 1, number: 1 },
      ],
      0,
      "mac-365-ja_JP",
    ]);
    assert.deepEqual(result.classIds, [1, 1, 0]);
    assert.equal(result.classCount, 2);
    assert.equal(session.dirty, false);
  } finally {
    closeSession("collation");
  }
});

const dirtyOf = (id) => listSessions().find((session) => session.id === id)?.dirty;

/** Runs `fn` against a saved (clean) session, then closes it and removes its files. */
async function withCleanSession(id, prepare, fn) {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-call-"));
  await openSession(undefined, id);
  try {
    const call = (method, args = []) => callWorkbookMethod(id, method, args);
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
  const record = { ...call("getCellXf", [0]).result, status: undefined };
  return call("addXf", [{ ...record, numFmtId: 10, applyNumberFormat: true }]).result.index;
}

test("every read-only method is allowlisted and leaves the session clean", async () => {
  for (const method of READ_ONLY_METHODS) {
    assert.equal(WORKBOOK_METHODS.has(method), true, `${method} is read-only but not allowlisted`);
  }
  await openSession(undefined, "readonly-sweep");
  try {
    for (const method of READ_ONLY_METHODS) {
      // A call that fails on its arguments must not mark the session dirty either.
      try {
        callWorkbookMethod("readonly-sweep", method, []);
      } catch {
        // The classification is under test, not the arguments.
      }
      assert.equal(dirtyOf("readonly-sweep"), false, `${method} dirtied the session`);
    }
  } finally {
    closeSession("readonly-sweep");
  }
});

test("every allowlisted method outside the read-only set marks the session dirty", async () => {
  const mutating = [...WORKBOOK_METHODS].filter((method) => !READ_ONLY_METHODS.has(method));
  assert.ok(mutating.length > 100);
  for (const required of [
    "setColumnStyle",
    "clearColumnStyle",
    "setRowStyle",
    "clearRowStyle",
    "clearColumnWidth",
    "clearCellXfIndex",
    "remapRangeStyle",
    "setThemeScriptFont",
    "resetTheme",
    "setImageAnchor",
    "setImageZOrder",
  ]) {
    assert.equal(mutating.includes(required), true, `${required} should be classified as mutating`);
  }
  for (const method of mutating) {
    const id = `mutating-sweep-${method}`;
    await openSession(undefined, id);
    try {
      assert.equal(dirtyOf(id), false);
      try {
        callWorkbookMethod(id, method, []);
      } catch {
        // A failed or malformed mutation stays dirty: it may have changed native state first.
      }
      assert.equal(dirtyOf(id), true, `${method} left the session clean`);
    } finally {
      closeSession(id);
    }
  }
});

test("rejects the engine methods that are withheld from the allowlist", async () => {
  await openSession(undefined, "withheld");
  try {
    for (const method of [
      "snapshotImage",
      "restoreImage",
      "snapshotRangeStyle",
      "restoreRangeStyle",
      "save",
      "saveAs",
      "saveWithDiagnostics",
      "setIterativeProgress",
      "probeImage",
      "insertImage",
      "constructor",
      "delete",
      "__proto__",
    ]) {
      assert.equal(WORKBOOK_METHODS.has(method), false);
      assert.throws(
        () => callWorkbookMethod("withheld", method, []),
        /not allowlisted/,
        `${method} was dispatched`,
      );
    }
    assert.equal(dirtyOf("withheld"), false);
  } finally {
    closeSession("withheld");
  }
});

test("style reads and locale reads leave the session clean", async () => {
  await withCleanSession("call-read-new", undefined, (call) => {
    const clean = (name, args) => {
      const called = call(name, args);
      assert.equal(called.session.dirty, false, `${name} dirtied the session`);
      return called.result;
    };
    assert.deepEqual(clean("rangeStyleBases", [0, 0, 0, 3, 3]).xfIndices, [0]);
    assert.equal(clean("localizeFormula", ["=SUM(1.5)", "win-365-de_DE"]).value, "=SUMME(1,5)");
    assert.equal(clean("canonicalizeFormula", ["=SUMME(1,5)", "win-365-de_DE"]).value, "=SUM(1.5)");
    assert.equal(clean("localeFacts", ["win-365-de_DE"]).facts.decimalSeparator, ",");
    assert.equal(clean("compareText", ["a", "b", 0, "win-365-en_US"]).order, -1);
    assert.deepEqual(
      clean("collationClasses", [
        [
          { kind: 3, text: "x" },
          { kind: 3, text: "y" },
        ],
        0,
        "win-365-en_US",
      ]).classIds,
      [0, 1],
    );
  });
});

test("column and row styles are set and cleared as mutations", async () => {
  let xf;
  await withCleanSession(
    "call-line-styles",
    (call) => {
      xf = percentXf(call);
    },
    async (call, dir) => {
      const columns = () => call("getSheetColumns", [0]).result.columns;
      const rows = () => call("getSheetRowOverrides", [0]).result.rows;

      const set = call("setColumnStyle", [0, 1, 3, xf]);
      assert.equal(set.session.dirty, true);
      assert.deepEqual(
        columns().map(({ first, last, hasStyle, styleXf }) => ({ first, last, hasStyle, styleXf })),
        [{ first: 1, last: 3, hasStyle: 1, styleXf: xf }],
      );
      assert.equal(call("getEffectiveStyle", [0, 50, 2]).result.numFmtCode, "0.00%");
      await saveSession("call-line-styles", path.join(dir, "after-set.xlsx"));
      assert.equal(dirtyOf("call-line-styles"), false);
      assert.equal(call("clearColumnStyle", [0, 1, 3]).session.dirty, true);
      assert.deepEqual(columns(), []);
      assert.equal(call("getEffectiveStyle", [0, 50, 2]).result.numFmtCode, "General");

      await saveSession("call-line-styles", path.join(dir, "after-clear.xlsx"));
      assert.equal(call("setRowStyle", [0, 4, 6, xf]).session.dirty, true);
      assert.deepEqual(
        rows().map(({ row, hasStyle, styleXf }) => ({ row, hasStyle, styleXf })),
        [4, 5, 6].map((row) => ({ row, hasStyle: 1, styleXf: xf })),
      );
      assert.equal(call("getEffectiveStyle", [0, 5, 40]).result.numFmtCode, "0.00%");
      await saveSession("call-line-styles", path.join(dir, "after-row.xlsx"));
      assert.equal(call("clearRowStyle", [0, 4, 6]).session.dirty, true);
      assert.deepEqual(rows(), []);
    },
  );
});

test("clearing a column width keeps what else the column carries", async () => {
  let xf;
  await withCleanSession(
    "call-clear-width",
    (call) => {
      xf = percentXf(call);
      call("setColumnWidth", [0, 1, 3, 20]);
      call("setColumnHidden", [0, 2, 2, true]);
      call("setColumnStyle", [0, 1, 3, xf]);
    },
    (call) => {
      const columns = () => call("getSheetColumns", [0]).result.columns;
      assert.ok(columns().every((column) => column.hasWidth === 1));
      const cleared = call("clearColumnWidth", [0, 1, 3]);
      assert.equal(cleared.session.dirty, true);
      const after = columns();
      assert.ok(after.length >= 1);
      for (const column of after) {
        assert.equal(column.hasWidth, 0);
        assert.equal(column.hasStyle, 1);
        assert.equal(column.styleXf, xf);
      }
      assert.equal(
        after.some((column) => column.first <= 2 && 2 <= column.last && column.hidden === 1),
        true,
      );
      // Clearing a column that has no width is a successful no-op.
      assert.equal(call("clearColumnWidth", [0, 10, 12]).result.ok, true);
    },
  );
});

test("clearing a cell xf returns the cell to the default style and keeps its value", async () => {
  let xf;
  await withCleanSession(
    "call-clear-xf",
    (call) => {
      xf = percentXf(call);
      call("setNumber", [0, 0, 0, 5]);
      call("setCellXfIndex", [0, 0, 0, xf]);
    },
    (call) => {
      assert.equal(call("getCellXfIndex", [0, 0, 0]).result.xfIndex, xf);
      const cleared = call("clearCellXfIndex", [0, 0, 0]);
      assert.equal(cleared.session.dirty, true);
      assert.equal(call("getCellXfIndex", [0, 0, 0]).result.xfIndex, 0);
      assert.equal(call("getEffectiveStyle", [0, 0, 0]).result.numFmtCode, "General");
      assert.equal(call("getValue", [0, 0, 0]).result.value.value, 5);
    },
  );
});

test("remapRangeStyle rewrites styles in place and reports bad input without changing cells", async () => {
  let xf;
  await withCleanSession(
    "call-remap",
    (call) => {
      xf = percentXf(call);
    },
    async (call, dir) => {
      const bases = call("rangeStyleBases", [0, 0, 0, 1, 1]);
      assert.deepEqual(bases.result.xfIndices, [0]);
      const remapped = call("remapRangeStyle", [0, 0, 0, 1, 1, [0], [xf]]);
      assert.equal(remapped.session.dirty, true);
      for (const [row, col] of [
        [0, 0],
        [0, 1],
        [1, 0],
        [1, 1],
      ]) {
        assert.equal(call("getCellXfIndex", [0, row, col]).result.xfIndex, xf);
      }
      assert.equal(call("rangeStyleBases", [0, 0, 0, 1, 1]).result.xfIndices.includes(xf), true);
      assert.equal(call("getEffectiveStyle", [0, 1, 1]).result.numFmtCode, "0.00%");
      assert.equal(call("getEffectiveStyle", [0, 2, 2]).result.numFmtCode, "General");

      // Mismatched arrays, and a target past the xf table, fail and change nothing.
      await saveSession("call-remap", path.join(dir, "remapped.xlsx"));
      const before = call("getSheetColumns", [0]).result;
      assert.throws(
        () => callWorkbookMethod("call-remap", "remapRangeStyle", [0, 0, 0, 1, 1, [xf], [0, 1]]),
        /workbook remapRangeStyle failed: .*same length/,
      );
      assert.throws(
        () => callWorkbookMethod("call-remap", "remapRangeStyle", [0, 0, 0, 1, 1, [xf], [99999]]),
        /workbook remapRangeStyle failed/,
      );
      assert.equal(call("getCellXfIndex", [0, 1, 1]).result.xfIndex, xf);
      assert.equal(call("getEffectiveStyle", [0, 2, 2]).result.numFmtCode, "General");
      assert.deepEqual(call("getSheetColumns", [0]).result, before);
      // A failed mutation still leaves the session marked dirty.
      assert.equal(dirtyOf("call-remap"), true);
    },
  );
});

test("theme script fonts are set and reset as mutations", async () => {
  await withCleanSession("call-theme", undefined, async (call, dir) => {
    const jpan = (major) =>
      call("getTheme", []).result[major ? "majorScripts" : "minorScripts"].find(
        (entry) => entry.script === "Jpan",
      ).typeface;
    const original = { major: jpan(true), minor: jpan(false) };

    const set = call("setThemeScriptFont", [true, "Jpan", "Meiryo"]);
    assert.equal(set.session.dirty, true);
    assert.equal(jpan(true), "Meiryo");
    assert.equal(jpan(false), original.minor, "the minor set is untouched");
    assert.equal(call("getTheme", []).result.fonts.majorEastAsian, "Meiryo");

    await saveSession("call-theme", path.join(dir, "themed.xlsx"));
    assert.equal(call("setThemeScriptFont", [false, "Jpan", "Yu Mincho"]).session.dirty, true);
    assert.equal(jpan(false), "Yu Mincho");
    assert.throws(
      () => callWorkbookMethod("call-theme", "setThemeScriptFont", [true, "Jpan", ""]),
      /typeface must not be empty/,
    );

    await saveSession("call-theme", path.join(dir, "themed-again.xlsx"));
    const reset = call("resetTheme", []);
    assert.equal(reset.session.dirty, true);
    assert.deepEqual({ major: jpan(true), minor: jpan(false) }, original);
  });
});

test("image anchor and stacking methods are mutations", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-call-images-"));
  const picture = path.join(dir, "pixel.png");
  await writeFile(picture, solidPng(20, 10));
  await openSession(undefined, "call-images");
  try {
    const first = insertSessionImage("call-images", 0, { path: picture, anchor: "B2" }).objectId;
    const second = insertSessionImage("call-images", 0, { path: picture, anchor: "D4" }).objectId;
    await saveSession("call-images", path.join(dir, "images.xlsx"));
    assert.equal(dirtyOf("call-images"), false);

    const listed = () =>
      callWorkbookMethod("call-images", "listDrawingObjects", [0]).result.items.map((item) => ({
        objectId: item.objectId,
        row: item.fromRow,
        col: item.fromCol,
      }));
    assert.equal(callWorkbookMethod("call-images", "listDrawingObjects", [0]).session.dirty, false);
    assert.deepEqual(
      listed().map((item) => item.objectId),
      [first, second],
    );

    const anchored = callWorkbookMethod("call-images", "setImageAnchor", [
      0,
      first,
      { anchorKind: 0, row: 7, col: 5, colOffEmu: 0, rowOffEmu: 0, widthEmu: 0, heightEmu: 0 },
    ]);
    assert.equal(anchored.session.dirty, true);
    assert.deepEqual(
      listed().find((item) => item.objectId === first),
      { objectId: first, row: 7, col: 5 },
    );

    await saveSession("call-images", path.join(dir, "images-2.xlsx"));
    const stacked = callWorkbookMethod("call-images", "setImageZOrder", [0, first, 1]);
    assert.equal(stacked.session.dirty, true);
    assert.deepEqual(
      listed().map((item) => item.objectId),
      [second, first],
    );
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
  await openSession(undefined, "compare-mac");
  try {
    for (const profile of ["mac-365-en_US", "mac-365-ja_JP", "mac-365-de_DE"]) {
      for (const strength of [0, 1, 2]) {
        for (const [a, b, order] of COMPARE_CASES) {
          const called = callWorkbookMethod("compare-mac", "compareText", [
            a,
            b,
            strength,
            profile,
          ]);
          assert.equal(called.result.order, order, `${profile}/${strength}: ${a} vs ${b}`);
          assert.equal(called.session.dirty, false);
        }
      }
    }
    // Strength picks the fold: Mac and Windows collation both equate sharp s
    // with ss except for exact classic matching.
    const compare = (a, b, strength, profile) =>
      callWorkbookMethod("compare-mac", "compareText", [a, b, strength, profile]).result.order;
    assert.equal(compare("ß", "ss", 0, "mac-365-en_US"), 0);
    assert.equal(compare("ß", "ss", 2, "mac-365-en_US"), 1);
    assert.equal(compare("ß", "ss", 0, "win-365-en_US"), 0);
    assert.equal(compare("ß", "ss", 2, "win-365-en_US"), 1);
    assert.throws(
      () => callWorkbookMethod("compare-mac", "compareText", ["a", "b", 0, "mac-365-xx_XX"]),
      /workbook compareText failed/,
    );
  } finally {
    closeSession("compare-mac");
  }
});
