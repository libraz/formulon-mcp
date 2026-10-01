import assert from "node:assert/strict";
import { test } from "vitest";
import {
  applySessionMutations,
  closeSession,
  findSessionCells,
  getSessionCellByA1,
  openSession,
  replaceSessionCells,
} from "../../dist/sessions.js";

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
