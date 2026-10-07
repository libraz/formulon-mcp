import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import {
  applySessionMutations,
  applySheetOperation,
  callWorkbookMethod,
  closeSession,
  getSession,
  getSessionCell,
  getSessionCellByA1,
  getSessionRange,
  listSessions,
  openSession,
  recalcSession,
  replaceSessionCells,
  saveSession,
  setSessionRange,
  styleSessionRange,
} from "../../dist/sessions.js";

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

test("keeps completed writes dirty when iterative recalculation aborts", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-abort-"));
  await openSession(undefined, "iterative-abort");
  const workbook = getSession("iterative-abort").workbook;
  let callbacks = 0;
  try {
    callWorkbookMethod("iterative-abort", "setIterative", [true, 100, 0.0001]);
    applySessionMutations(
      "iterative-abort",
      [{ type: "formula", a1: "A1", formula: "=(A1+10)/2" }],
      false,
    );
    await saveSession("iterative-abort", path.join(dir, "book.xlsx"));
    assert.equal(getSession("iterative-abort").dirty, false);
    assert.equal(
      workbook.setIterativeProgress(() => {
        callbacks += 1;
        throw new Error("stop iterative solve");
      }).ok,
      true,
    );
    assert.throws(
      () =>
        applySessionMutations("iterative-abort", [{ type: "number", a1: "B1", value: 7 }], true),
      /recalc workbook failed/,
    );
    assert.equal(callbacks, 1);
    assert.equal(getSessionCellByA1("iterative-abort", "B1").value.value, 7);
    assert.equal(getSession("iterative-abort").dirty, true);

    applySessionMutations(
      "iterative-abort",
      [{ type: "formula", a1: "A1", formula: "=(A1+20)/2" }],
      false,
    );
    await saveSession("iterative-abort", path.join(dir, "book.xlsx"));
    assert.equal(getSession("iterative-abort").dirty, false);
    assert.throws(() => recalcSession("iterative-abort"), /recalc workbook failed/);
    assert.equal(callbacks, 2);
    assert.equal(getSession("iterative-abort").dirty, true);
  } finally {
    workbook.setIterativeProgress(null);
    closeSession("iterative-abort");
    await rm(dir, { recursive: true, force: true });
  }
});

test("keeps completed replacements dirty when recalculation aborts", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-replace-failure-"));
  await openSession(undefined, "replace-failure");
  const workbook = getSession("replace-failure").workbook;
  let callbacks = 0;
  try {
    callWorkbookMethod("replace-failure", "setIterative", [true, 100, 0.0001]);
    applySessionMutations(
      "replace-failure",
      [
        { type: "text", a1: "A1", value: "SUM" },
        { type: "formula", a1: "B1", formula: "=(B1+10)/2" },
      ],
      false,
    );
    await saveSession("replace-failure", path.join(dir, "book.xlsx"));
    assert.equal(getSession("replace-failure").dirty, false);
    assert.equal(
      workbook.setIterativeProgress(() => {
        callbacks += 1;
        throw new Error("stop iterative solve");
      }).ok,
      true,
    );
    const replacement = "Replaced";
    assert.throws(
      () =>
        replaceSessionCells("replace-failure", "SUM", {
          replacement,
          target: "texts",
          matchCase: true,
          wholeCell: false,
          regex: false,
          maxResults: 100,
          maxReplacements: 100,
          recalc: true,
        }),
      /recalc workbook failed/,
    );
    assert.equal(callbacks, 1);
    assert.equal(getSessionCellByA1("replace-failure", "A1").value.value, replacement);
    assert.equal(getSessionCellByA1("replace-failure", "B1").formula, "=(B1+10)/2");
    assert.equal(getSession("replace-failure").dirty, true);
  } finally {
    workbook.setIterativeProgress(null);
    closeSession("replace-failure");
    await rm(dir, { recursive: true, force: true });
  }
});

test("cell and range reads mark recalculated cached values dirty", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-read-recalc-"));
  await openSession(undefined, "read-recalc");
  const workbook = getSession("read-recalc").workbook;
  try {
    for (const read of [
      () => getSessionCellByA1("read-recalc", "A1", undefined, true),
      () =>
        getSessionRange("read-recalc", "A1:B1", {
          maxCells: 100,
          recalc: true,
          includeFormulas: true,
        }),
    ]) {
      applySessionMutations(
        "read-recalc",
        [{ type: "formula", a1: "A1", formula: "=SUM(1,2)" }],
        false,
      );
      await saveSession("read-recalc", path.join(dir, "book.xlsx"));
      assert.equal(getSession("read-recalc").dirty, false);
      read();
      assert.equal(getSession("read-recalc").dirty, true);
      assert.equal(getSessionCellByA1("read-recalc", "A1").value.value, 3);

      assert.equal(workbook.setIterative(true, 100, 0.0001).ok, true);
      assert.equal(workbook.setFormula(0, 0, 0, "=(A1+10)/2").ok, true);
      await saveSession("read-recalc", path.join(dir, "book.xlsx"));
      let callbacks = 0;
      workbook.setIterativeProgress(() => {
        callbacks += 1;
        throw new Error("abort read recalc");
      });
      assert.throws(read, /recalc workbook failed/);
      assert.equal(callbacks, 1);
      assert.equal(getSession("read-recalc").dirty, true);
      workbook.setIterativeProgress(null);
    }
  } finally {
    workbook.setIterativeProgress(null);
    closeSession("read-recalc");
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

test("prevalidates non-finite numbers before writing any mutation", async () => {
  await openSession(undefined, "finite-mutations");
  try {
    assert.throws(
      () =>
        applySessionMutations(
          "finite-mutations",
          [
            { type: "number", a1: "Sheet1!A1", value: 7 },
            { type: "number", a1: "Sheet1!B1", value: Number.NaN },
          ],
          false,
        ),
      /number value must be finite/,
    );
    assert.deepEqual(getSessionCellByA1("finite-mutations", "Sheet1!A1").value, {
      kind: "blank",
    });
    assert.equal(listSessions().find((session) => session.id === "finite-mutations")?.dirty, false);
  } finally {
    closeSession("finite-mutations");
  }
});

test("reads a subrange of a large sheet with exact formulas", async () => {
  await openSession(undefined, "range-large");
  try {
    const mutations = [];
    for (let row = 1; row <= 2000; row += 1) {
      mutations.push(
        { type: "number", a1: `A${row}`, value: row },
        { type: "formula", a1: `B${row}`, formula: `=A${row}*2` },
      );
    }
    applySessionMutations("range-large", mutations, true);

    const range = getSessionRange("range-large", "A1000:B1001", {
      maxCells: 100,
      recalc: false,
      includeFormulas: true,
    });
    assert.deepEqual(
      range.cells.map((cell) => [cell.a1, cell.formula]),
      [
        ["A1000", ""],
        ["B1000", "=A1000*2"],
        ["A1001", ""],
        ["B1001", "=A1001*2"],
      ],
    );
    assert.equal(getSessionCellByA1("range-large", "B1500").formula, "=A1500*2");

    const capped = getSessionRange("range-large", "A1:B2000", {
      maxCells: 3,
      recalc: false,
      includeFormulas: false,
    });
    assert.equal(capped.cellCount, 3);
    assert.equal(capped.truncated, true);
    const whole = getSessionRange("range-large", "A1:B2", {
      maxCells: 4,
      recalc: false,
      includeFormulas: false,
    });
    assert.equal(whole.truncated, false);
  } finally {
    closeSession("range-large");
  }
});

test("formatted numbers carry the engine display text", async () => {
  await openSession(undefined, "range-text");
  try {
    setSessionRange("range-text", "A1", [[1234.5], [0.75], [45000], [7]], undefined, false);
    styleSessionRange(
      "range-text",
      "A1:A1",
      { numberFormat: '"¥"#,##0.00' },
      undefined,
      "existing",
    );
    styleSessionRange("range-text", "A2:A2", { numberFormat: "# ?/?" }, undefined, "existing");
    styleSessionRange("range-text", "A3:A3", { numberFormat: "yyyy/mm/dd" }, undefined, "existing");

    const range = getSessionRange("range-text", "A1:A4", {
      maxCells: 10,
      recalc: false,
      includeFormulas: false,
    });
    assert.deepEqual(
      range.cells.map((cell) => cell.text),
      ["¥1,234.50", " 3/4", "2023/03/15", undefined],
    );
    assert.equal(range.cells[1].formatKind, undefined);
    assert.equal(range.cells[2].formatted, "2023-03-15");
    assert.equal(getSessionCellByA1("range-text", "A1").text, "¥1,234.50");
    assert.equal(getSessionCellByA1("range-text", "A4").text, undefined);
  } finally {
    closeSession("range-text");
  }
});
