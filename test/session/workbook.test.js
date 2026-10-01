import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { loadWorkbook } from "../../dist/formulon.js";
import {
  applySheetOperation,
  callWorkbookMethod,
  closeSession,
  getSession,
  inspectSession,
  openSession,
  resolveSessionSheet,
  saveSession,
  setSessionDefinedName,
} from "../../dist/sessions.js";

test("scopes a defined name to one sheet and reports the scope back", async () => {
  await openSession(undefined, "defined-name-scope");
  try {
    applySheetOperation("defined-name-scope", "add", { name: "Sheet2" });

    const workbookScoped = setSessionDefinedName("defined-name-scope", "Rate", "0.08");
    assert.equal(workbookScoped.localSheetId, -1);

    // Print settings only take effect sheet-locally, so the scope has to reach
    // the engine rather than being flattened to workbook scope.
    const sheetScoped = setSessionDefinedName(
      "defined-name-scope",
      "_xlnm.Print_Area",
      "'Sheet2'!$A$1:$C$3",
      "Sheet2",
    );
    assert.equal(sheetScoped.localSheetId, 1);

    const { workbook } = inspectSession("defined-name-scope");
    const byName = Object.fromEntries(workbook.definedNames.map((entry) => [entry.name, entry]));
    assert.equal(byName.Rate.localSheetId, -1);
    assert.equal(byName["_xlnm.Print_Area"].localSheetId, 1);
    assert.equal(byName["_xlnm.Print_Area"].formula, "'Sheet2'!$A$1:$C$3");
  } finally {
    closeSession("defined-name-scope");
  }
});

test("supports sheet operations", async () => {
  await openSession(undefined, "sheets");
  try {
    callWorkbookMethod("sheets", "addSheet", ["Data"]);
    assert.equal(callWorkbookMethod("sheets", "sheetCount", []).result.value, 2);
    callWorkbookMethod("sheets", "renameSheet", [1, "Input"]);
    assert.equal(callWorkbookMethod("sheets", "sheetName", [1]).result.value, "Input");
    callWorkbookMethod("sheets", "moveSheet", [1, 0]);
    assert.equal(callWorkbookMethod("sheets", "sheetName", [0]).result.value, "Input");
    callWorkbookMethod("sheets", "removeSheet", [0]);
    assert.equal(callWorkbookMethod("sheets", "sheetCount", []).result.value, 1);
  } finally {
    closeSession("sheets");
  }
});

test("sheet-operation argument failures leave a clean session unchanged", async () => {
  await openSession(undefined, "sheet-preflight");
  try {
    for (const operation of ["add", "remove", "rename", "move"]) {
      assert.throws(() => applySheetOperation("sheet-preflight", operation, {}), /is required/);
      assert.equal(getSession("sheet-preflight").dirty, false);
    }
    assert.throws(
      () => applySheetOperation("sheet-preflight", "rename", { index: 0 }),
      /newName is required/,
    );
    assert.equal(getSession("sheet-preflight").dirty, false);
    assert.equal(callWorkbookMethod("sheet-preflight", "sheetCount", []).result.value, 1);
  } finally {
    closeSession("sheet-preflight");
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
