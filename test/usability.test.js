import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { loadWorkbook } from "../dist/formulon.js";
import {
  applySessionMutations,
  closeSession,
  getSessionCellByA1,
  openSession,
  saveSession,
} from "../dist/sessions.js";
import { callTool, errorPayload, textPayload, withClient } from "./helpers/mcp.js";

test("set_range distinguishes an explicit blank from a skipped cell", async () => {
  await withClient(async (client) => {
    const sessionId = "usability-range";
    await callTool(client, "formulon_open_workbook", { sessionId });
    try {
      const initial = textPayload(
        await callTool(client, "formulon_set_range", {
          sessionId,
          start: "Sheet1!A1",
          values: [[1, 2]],
        }),
      );
      assert.equal(initial.cellsWritten, 2);

      const updated = textPayload(
        await callTool(client, "formulon_set_range", {
          sessionId,
          start: "Sheet1!A1",
          values: [[null, { blank: true }]],
        }),
      );
      assert.equal(updated.cellsWritten, 1);
      assert.deepEqual(
        textPayload(
          await callTool(client, "formulon_get_cell", {
            sessionId,
            a1: "Sheet1!A1",
          }),
        ).value,
        { kind: "number", value: 1 },
      );
      assert.deepEqual(
        textPayload(
          await callTool(client, "formulon_get_cell", {
            sessionId,
            a1: "Sheet1!B1",
          }),
        ).value,
        { kind: "blank" },
      );

      const invalid = await callTool(client, "formulon_set_range", {
        sessionId,
        start: "Sheet1!A1",
        values: [[{ blank: false }]],
      });
      assert.match(errorPayload(invalid), /invalid|blank/i);
    } finally {
      await callTool(client, "formulon_close_workbook", { sessionId });
    }
  });
});

test("set_range preflights ragged trailing cells without dirtying a clean session", async () => {
  await withClient(async (client) => {
    const sessionId = "usability-range-bounds";
    await callTool(client, "formulon_open_workbook", { sessionId });
    try {
      const noOp = textPayload(
        await callTool(client, "formulon_set_range", {
          sessionId,
          start: "Sheet1!B2",
          values: [[null, null], [null]],
        }),
      );
      assert.equal(noOp.cellsWritten, 0);
      const missingSheet = await callTool(client, "formulon_get_cell", {
        sessionId,
        sheet: "Missing",
        a1: "A1",
      });
      assert.match(errorPayload(missingSheet), /sheet not found/i);
      const outOfGrid = await callTool(client, "formulon_get_cell", {
        sessionId,
        a1: "A1048577",
      });
      assert.match(errorPayload(outOfGrid), /exceeds Excel sheet limits/i);
      for (const [start, values] of [
        ["Sheet1!XFD1", [[1, null]]],
        ["Sheet1!A1048576", [[1], [null]]],
      ]) {
        const outOfGrid = await callTool(client, "formulon_set_range", {
          sessionId,
          start,
          values,
        });
        assert.match(errorPayload(outOfGrid), /exceeds Excel sheet limits/i);
      }
      const afterRejectedRanges = textPayload(
        await callTool(client, "formulon_inspect_session", { sessionId }),
      );
      assert.equal(afterRejectedRanges.session.dirty, false);
    } finally {
      await callTool(client, "formulon_close_workbook", { sessionId });
    }
  });
});

test("a mutation concurrent with save keeps the session dirty while the artifact stays consistent", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-usability-race-"));
  const outputPath = path.join(dir, "race.xlsx");
  const sessionId = "usability-save-race";
  try {
    await openSession(undefined, sessionId);
    applySessionMutations(sessionId, [{ type: "number", a1: "Sheet1!A1", value: 1 }], true);
    const pendingSave = saveSession(sessionId, outputPath);
    applySessionMutations(sessionId, [{ type: "number", a1: "Sheet1!A1", value: 2 }], true);
    const saved = await pendingSave;
    assert.equal(saved.session.dirty, true);
    assert.deepEqual(getSessionCellByA1(sessionId, "Sheet1!A1").value, {
      kind: "number",
      value: 2,
    });

    const workbook = await loadWorkbook(outputPath);
    try {
      const artifactCell = workbook.getValue(0, 0, 0);
      assert.equal(artifactCell.status.ok, true);
      assert.equal(artifactCell.value.number, 1);
    } finally {
      workbook.delete();
    }
  } finally {
    closeSession(sessionId);
    await rm(dir, { recursive: true, force: true });
  }
});

test("path cell reads match session reads for named sheets, currency formats, and formulas", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-usability-read-"));
  const outputPath = path.join(dir, "parity.xlsx");
  try {
    await withClient(async (client) => {
      const sessionId = "usability-read";
      await callTool(client, "formulon_open_workbook", { sessionId });
      try {
        await callTool(client, "formulon_sheet_operation", {
          sessionId,
          operation: "add",
          name: "Data",
        });
        await callTool(client, "formulon_set_range", {
          sessionId,
          start: "Data!A1",
          values: [[1234.5, { f: "=A1*2" }]],
        });
        await callTool(client, "formulon_style_range", {
          sessionId,
          range: "Data!A1",
          style: { numberFormat: '"¥"#,##0' },
        });
        const saved = textPayload(
          await callTool(client, "formulon_save_session", { sessionId, outputPath }),
        );
        assert.equal(saved.format, "xlsx");

        const sessionCurrency = textPayload(
          await callTool(client, "formulon_get_cell", {
            sessionId,
            a1: "Data!A1",
          }),
        );
        const pathCurrency = textPayload(
          await callTool(client, "formulon_get_cell", {
            path: outputPath,
            a1: "'Data'!A1",
          }),
        );
        assert.deepEqual(pathCurrency.address, sessionCurrency.address);
        assert.deepEqual(pathCurrency.value, sessionCurrency.value);
        assert.equal(pathCurrency.numberFormat, sessionCurrency.numberFormat);
        assert.equal(pathCurrency.formatKind, "currency");

        const sessionFormula = textPayload(
          await callTool(client, "formulon_get_cell", {
            sessionId,
            a1: "Data!B1",
          }),
        );
        const pathFormula = textPayload(
          await callTool(client, "formulon_get_cell", {
            path: outputPath,
            sheet: "Data",
            row: 0,
            col: 1,
          }),
        );
        assert.deepEqual(pathFormula.address, sessionFormula.address);
        assert.deepEqual(pathFormula.value, sessionFormula.value);
        assert.equal(pathFormula.formula, sessionFormula.formula);
      } finally {
        await callTool(client, "formulon_close_workbook", { sessionId });
      }
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("get_cell requires exactly one source and formula evaluation resolves session sheet names", async () => {
  await withClient(async (client) => {
    const sessionId = "usability-eval";
    await callTool(client, "formulon_open_workbook", { sessionId });
    try {
      const neither = await callTool(client, "formulon_get_cell", {
        a1: "A1",
      });
      assert.match(errorPayload(neither), /exactly one of sessionId or path/i);

      const both = await callTool(client, "formulon_get_cell", {
        sessionId,
        path: "missing.xlsx",
        a1: "A1",
      });
      assert.match(errorPayload(both), /exactly one of sessionId or path/i);

      await callTool(client, "formulon_sheet_operation", {
        sessionId,
        operation: "add",
        name: "Data",
      });
      await callTool(client, "formulon_set_cells", {
        sessionId,
        mutations: [{ type: "number", a1: "Data!A1", value: 7 }],
      });
      const named = textPayload(
        await callTool(client, "formulon_eval_formula", {
          sessionId,
          sheet: "Data",
          formula: "=A1*2",
        }),
      );
      assert.deepEqual(named.result.value, { kind: "number", value: 14 });

      const globalNamed = await callTool(client, "formulon_eval_formula", {
        sheet: "Sheet1",
        formula: "=SUM(1,2)",
      });
      assert.match(errorPayload(globalNamed), /sheet names require sessionId/i);

      const malformed = textPayload(
        await callTool(client, "formulon_eval_formula", { formula: "=(" }),
      );
      assert.equal(malformed.value.kind, "error");
    } finally {
      await callTool(client, "formulon_close_workbook", { sessionId });
    }
  });
});

test("lossy XLSB saves stay dirty while a clean XLSX save clears dirty", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-usability-save-"));
  const xlsbPath = path.join(dir, "table.xlsb");
  const xlsxPath = path.join(dir, "table.xlsx");
  try {
    await withClient(async (client) => {
      const sessionId = "usability-save";
      await callTool(client, "formulon_open_workbook", { sessionId });
      try {
        const unsupported = await callTool(client, "formulon_save_session", {
          sessionId,
          outputPath: path.join(dir, "table.csv"),
        });
        assert.match(errorPayload(unsupported), /\.xlsx or \.xlsb/i);
        await callTool(client, "formulon_set_range", {
          sessionId,
          start: "Sheet1!A1",
          values: [
            ["Item", "Qty"],
            ["Pen", 2],
          ],
        });
        const created = textPayload(
          await callTool(client, "formulon_workbook_call", {
            sessionId,
            method: "createTable",
            args: [{ sheetIndex: 0, ref: "A1:B2", name: "Items", columns: ["Item", "Qty"] }],
          }),
        );
        assert.equal(created.result.status.ok, true);

        const xlsb = textPayload(
          await callTool(client, "formulon_save_session", {
            sessionId,
            outputPath: xlsbPath,
          }),
        );
        assert.equal(xlsb.format, "xlsb");
        assert.equal(xlsb.session.dirty, true);
        assert.ok((xlsb.losses?.deferredFeatures ?? 0) > 0);

        const reloadedXlsb = textPayload(
          await callTool(client, "formulon_inspect_workbook", {
            path: xlsbPath,
          }),
        );
        assert.equal(reloadedXlsb.tables.length, 0);

        const xlsx = textPayload(
          await callTool(client, "formulon_save_session", {
            sessionId,
            outputPath: xlsxPath,
          }),
        );
        assert.equal(xlsx.losses, undefined);
        assert.equal(xlsx.session.dirty, false);
        const reloadedXlsx = textPayload(
          await callTool(client, "formulon_inspect_workbook", {
            path: xlsxPath,
          }),
        );
        assert.equal(reloadedXlsx.tables.length, 1);
        assert.equal(reloadedXlsx.tables[0].index, 0);
      } finally {
        await callTool(client, "formulon_close_workbook", { sessionId });
      }
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
