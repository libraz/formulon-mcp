import assert from "node:assert/strict";
import { test } from "vitest";
import {
  applySessionMutations,
  auditSessionFormulas,
  closeSession,
  getSession,
  openSession,
} from "../dist/sessions.js";
import { textPayload, withClient } from "./helpers/mcp.js";

function tableMutations(firstRow, lastRow) {
  const mutations = [];
  for (let row = firstRow; row <= lastRow; row += 1) {
    mutations.push(
      { type: "number", a1: `Sheet1!B${row}`, value: row },
      { type: "number", a1: `Sheet1!C${row}`, value: 2 },
      { type: "formula", a1: `Sheet1!D${row}`, formula: `=B${row}*C${row}` },
    );
  }
  return mutations;
}

test("formula patterns compare copied formulas by their relative R1C1 text", async () => {
  await openSession(undefined, "formula-audit-patterns");
  try {
    const mutations = [];
    for (let row = 2; row <= 8; row += 1) {
      // Literals, names, absolute refs and casing must not split a copied run.
      const call = row % 2 === 0 ? "LOG10" : "log10";
      mutations.push(
        { type: "number", a1: `Sheet1!B${row}`, value: row },
        {
          type: "formula",
          a1: `Sheet1!D${row}`,
          formula: `=${call}(B${row})+$B$2+LEN("A1")+SUM(A:A)`,
        },
      );
    }
    mutations.push({
      type: "formula",
      a1: "Sheet1!D6",
      formula: '=LOG10(B5)+$B$2+LEN("A1")+SUM(A:A)',
    });
    applySessionMutations("formula-audit-patterns", mutations, true);
    const result = auditSessionFormulas("formula-audit-patterns", {
      range: "B2:D8",
      direction: "vertical",
      window: 5,
      minPeers: 3,
    });
    assert.equal(result.unsupportedFormulaCells, 0);
    assert.deepEqual(
      result.findings.map((finding) => [finding.a1, finding.type, finding.expectedPattern]),
      [["D6", "formula_pattern_outlier", 'LOG10(RC[-2])+R2C2+LEN("A1")+SUM(C[-3])']],
    );
  } finally {
    closeSession("formula-audit-patterns");
  }
});

test("formula audit reports constants, blanks, and outlier formulas in one run", async () => {
  await openSession(undefined, "formula-audit-findings");
  try {
    applySessionMutations("formula-audit-findings", tableMutations(2, 12), true);
    applySessionMutations(
      "formula-audit-findings",
      [
        { type: "number", a1: "Sheet1!D5", value: 99 },
        { type: "blank", a1: "Sheet1!D7" },
        { type: "formula", a1: "Sheet1!D9", formula: "=B9+C9" },
      ],
      true,
    );

    const result = auditSessionFormulas("formula-audit-findings", {
      range: "B2:D12",
      direction: "vertical",
      window: 5,
      minPeers: 3,
    });
    assert.equal(result.complete, true);
    assert.equal(result.findingCount, 3);
    assert.equal(result.findingsTruncated, false);
    assert.deepEqual(
      result.findings.map((finding) => [finding.a1, finding.type]),
      [
        ["D5", "constant_in_formula_run"],
        ["D7", "blank_in_formula_run"],
        ["D9", "formula_pattern_outlier"],
      ],
    );
    assert.equal(
      result.findings.every((finding) => finding.reviewNote.includes("Review only")),
      true,
    );
  } finally {
    closeSession("formula-audit-findings");
  }
});

test("formula audit treats blank rows as boundaries and reports formula errors independently", async () => {
  await openSession(undefined, "formula-audit-boundaries");
  try {
    applySessionMutations(
      "formula-audit-boundaries",
      [...tableMutations(2, 4), { type: "formula", a1: "Sheet1!D6", formula: "=B6*C6" }],
      true,
    );
    const bounded = auditSessionFormulas("formula-audit-boundaries", {
      range: "B2:D6",
      direction: "vertical",
      window: 3,
      minPeers: 3,
    });
    assert.equal(bounded.findingCount, 0);

    applySessionMutations(
      "formula-audit-boundaries",
      [
        { type: "number", a1: "Sheet1!B8", value: 1 },
        { type: "number", a1: "Sheet1!C8", value: 2 },
        { type: "formula", a1: "Sheet1!D8", formula: "=1/0" },
        { type: "formula", a1: "Sheet1!D9", formula: "=B9*C9" },
        { type: "number", a1: "Sheet1!B9", value: 1 },
        { type: "number", a1: "Sheet1!C9", value: 2 },
      ],
      true,
    );
    const errors = auditSessionFormulas("formula-audit-boundaries", {
      range: "B8:D9",
      direction: "vertical",
      window: 2,
      minPeers: 3,
    });
    const errorFinding = errors.findings.find((finding) => finding.a1 === "D8");
    assert.equal(errorFinding?.type, "formula_error");
    assert.equal(errorFinding?.value.kind, "error");
    assert.equal(errorFinding?.expectedPattern, undefined);
  } finally {
    closeSession("formula-audit-boundaries");
  }
});

test("formula audit preflights an oversized range before scanning or recalculating", async () => {
  await openSession(undefined, "formula-audit-limit");
  try {
    assert.throws(
      () =>
        auditSessionFormulas("formula-audit-limit", {
          range: "A1:A50001",
          maxCells: 50_000,
          recalc: true,
        }),
      /50001.*limit 50000/,
    );
    assert.equal(getSession("formula-audit-limit").dirty, false);
  } finally {
    closeSession("formula-audit-limit");
  }
});

test("formula audit scans a full 50,000-cell area and keeps late findings", async () => {
  await openSession(undefined, "formula-audit-late");
  try {
    const mutations = [];
    for (let row = 24_992; row <= 25_000; row += 1) {
      mutations.push(
        { type: "number", a1: `Sheet1!B${row}`, value: row },
        { type: "formula", a1: `Sheet1!C${row}`, formula: `=B${row}*2` },
      );
    }
    mutations.push({ type: "number", a1: "Sheet1!C24996", value: 0 });
    applySessionMutations("formula-audit-late", mutations, true);

    const result = auditSessionFormulas("formula-audit-late", {
      sheet: "Sheet1",
      range: "Sheet1!B1:C25000",
      direction: "vertical",
      window: 4,
      minPeers: 3,
      maxCells: 50_000,
      maxFindings: 1,
    });
    assert.equal(result.scannedCells, 50_000);
    assert.equal(result.findingCount, 1);
    assert.equal(result.findingsTruncated, false);
    assert.equal(result.findings[0].a1, "C24996");
    assert.equal(result.findings[0].type, "constant_in_formula_run");
  } finally {
    closeSession("formula-audit-late");
  }
});

test("MCP registers and returns the bounded formula audit", async () => {
  await withClient(async (client) => {
    const tools = await client.listTools();
    assert.equal(
      tools.tools.some((tool) => tool.name === "formulon_audit_formulas"),
      true,
    );
    const sessionId = "formula-audit-mcp";
    const opened = await client.callTool({
      name: "formulon_open_workbook",
      arguments: { sessionId },
    });
    assert.equal(textPayload(opened).session.id, sessionId);
    try {
      const written = await client.callTool({
        name: "formulon_set_cells",
        arguments: {
          sessionId,
          mutations: [
            { type: "number", a1: "Sheet1!B2", value: 2 },
            { type: "number", a1: "Sheet1!C2", value: 3 },
            { type: "formula", a1: "Sheet1!D2", formula: "=B2*C2" },
          ],
          recalc: true,
        },
      });
      assert.equal(textPayload(written).applied.length, 3);
      const audited = await client.callTool({
        name: "formulon_audit_formulas",
        arguments: { sessionId, range: "Sheet1!B2:D2" },
      });
      const payload = textPayload(audited);
      assert.equal(payload.complete, true);
      assert.equal(payload.scannedCells, 3);
      assert.equal(Array.isArray(payload.findings), true);
    } finally {
      const closed = await client.callTool({
        name: "formulon_close_workbook",
        arguments: { sessionId },
      });
      assert.equal(textPayload(closed).session.id, sessionId);
    }
  });
});

test("formula audit with an explicit range reads only that range", async () => {
  await openSession(undefined, "formula-audit-range");
  try {
    applySessionMutations("formula-audit-range", tableMutations(2, 12), true);
    const result = auditSessionFormulas("formula-audit-range", { range: "Sheet1!B2:D12" });
    assert.equal(result.range, "B2:D12");
    assert.equal(result.formulaCells, 11);
    assert.equal(result.findingCount, 0);
  } finally {
    closeSession("formula-audit-range");
  }
});
