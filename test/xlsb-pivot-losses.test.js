import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import {
  callWorkbookMethod,
  closeSession,
  createSessionPivot,
  openSession,
  saveSession,
  setSessionRange,
  styleSessionRange,
} from "../dist/sessions.js";

const sheet0 = { kind: "Sheet", sheet: 0 };
const pivot0 = { kind: "Pivot", sheet: 0, pivot: 0 };

/** The serial of a day plus a time of day in whole seconds, as the cache stores a date. */
const serial = (day, hours, minutes, seconds) =>
  day + (hours * 3600 + minutes * 60 + seconds) / 86400;

/** Serials for 2026-01-15 06:30:15, 2026-04-20 12:45:30 and 2026-08-05 18:15:45. */
const SOURCE = [
  ["Region", "Qty", "When", "Amount"],
  ["East", 3, serial(46037, 6, 30, 15), 20],
  ["East", 14, serial(46132, 12, 45, 30), 10],
  ["West", 27, serial(46239, 18, 15, 45), 5],
];

/** Writes the columns the pivot uses as its source (rows, then pages, then Amount), builds a pivot
 *  summing Amount over them and applies `edit`, which gets the source position of a column name. */
const buildPivot = (id, { rows = ["Region"], pages = [], edit } = {}) => {
  const columns = [...rows, ...pages, "Amount"];
  const at = (name) => columns.indexOf(name);
  const pick = columns.map((name) => SOURCE[0].indexOf(name));
  setSessionRange(
    id,
    "Sheet1!A1",
    SOURCE.map((row) => pick.map((c) => row[c])),
    0,
    true,
  );
  if (columns.includes("When")) {
    const col = String.fromCharCode(65 + at("When"));
    styleSessionRange(
      id,
      `Sheet1!${col}2:${col}4`,
      { numberFormat: "yyyy-mm-dd hh:mm:ss" },
      0,
      "existing",
    );
  }
  const created = createSessionPivot(id, {
    sourceRange: `A1:${String.fromCharCode(64 + columns.length)}4`,
    target: "G1",
    name: "P1",
    rows,
    pages,
    values: [{ field: "Amount" }],
  });
  edit?.(id, at, created);
  return created;
};

/** Saves `id` to .xlsb, reopens the file and returns the save's losses and the pivots read back. */
const saveAndReopen = async (id, dir) => {
  const file = path.join(dir, `${id}.xlsb`);
  const saved = await saveSession(id, file);
  await openSession(file, `${id}-back`);
  try {
    return {
      losses: saved.losses,
      pivots: callWorkbookMethod(`${id}-back`, sheet0, "listPivots", []).result.length,
    };
  } finally {
    closeSession(`${id}-back`);
  }
};

const call = (id, method, args) => callWorkbookMethod(id, pivot0, method, args).result;

/** PivotTables the engine writes to .xlsb: each is saved with no loss and read back. */
const WRITTEN = {
  plain: {},
  "a page filter": { pages: ["Region"], rows: ["Qty"] },
  "a date group by days, months and years": {
    rows: ["When"],
    edit: (id, at) => call(id, "setFieldDateGroup", [at("When"), { periods: 8 | 16 | 64 }]),
  },
  "a date group by quarters": {
    rows: ["When"],
    edit: (id, at) => call(id, "setFieldDateGroup", [at("When"), { periods: 32 | 64 }]),
  },
  "a date group by hours, minutes and seconds": {
    rows: ["When"],
    edit: (id, at) => call(id, "setFieldDateGroup", [at("When"), { periods: 1 | 2 | 4 }]),
  },
  "a numeric range group": {
    rows: ["Qty"],
    edit: (id, at) => call(id, "setFieldNumericGroup", [at("Qty"), { interval: 10 }]),
  },
  "a blank source cell": {
    edit: (id) => setSessionRange(id, "Sheet1!B3", [[{ blank: true }]], 0, true),
  },
  "a descending sort": {
    edit: (id, at) => call(id, "updateField", [at("Region"), { sortOrder: 2 }]),
  },
  "a sort by value": {
    edit: (id, at) => {
      const [data] = call(id, "listDataFields", []);
      call(id, "updateField", [at("Region"), { sortOrder: 2, sortByField: data.name }]);
    },
  },
  "values shown as rank": {
    edit: (id, at) => call(id, "setDataField", [0, { fieldIndex: at("Amount"), showAs: 15 }]),
  },
  "values shown as percent of running total": {
    edit: (id, at) =>
      call(id, "setDataField", [
        0,
        { fieldIndex: at("Amount"), showAs: 13, showAsBaseField: at("Region") },
      ]),
  },
  "a calculated field": {
    edit: (id, _at, created) => {
      const cache = { kind: "PivotCache", pivotCache: created.cacheId };
      const field = callWorkbookMethod(id, cache, "addCalculatedField", ["Double", "=Amount*2"]);
      call(id, "addDataField", [{ fieldIndex: field.result }]);
    },
  },
  "page wrap and empty rows and columns shown": {
    pages: ["Region"],
    rows: ["Qty"],
    edit: (id) =>
      call(id, "updateOptions", [{ pageWrap: 2, showEmptyRow: true, showEmptyColumn: true }]),
  },
};

/** PivotTables using a feature the engine does not write to .xlsb yet: left out and counted. */
const LEFT_OUT = {
  "alt text": {
    edit: (id) => call(id, "updateOptions", [{ altText: "Sales", altTextSummary: "By region" }]),
  },
  "a cache saved without its data": {
    edit: (id, _at, created) =>
      callWorkbookMethod(id, { kind: "PivotCache", pivotCache: created.cacheId }, "updateOptions", [
        { saveData: false },
      ]),
  },
};

for (const [label, spec] of Object.entries(WRITTEN)) {
  test(`an .xlsb save writes a PivotTable with ${label}`, async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-xlsb-pivot-"));
    const id = `written-${label.replace(/\W+/g, "-")}`;
    await openSession(undefined, id);
    try {
      buildPivot(id, spec);
      const back = await saveAndReopen(id, dir);
      assert.equal(back.losses, undefined);
      assert.equal(back.pivots, 1);
    } finally {
      closeSession(id);
      await rm(dir, { recursive: true, force: true });
    }
  });
}

for (const [label, spec] of Object.entries(LEFT_OUT)) {
  test(`an .xlsb save leaves out a PivotTable with ${label} and counts it`, async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-xlsb-pivot-"));
    const id = `left-out-${label.replace(/\W+/g, "-")}`;
    await openSession(undefined, id);
    try {
      buildPivot(id, spec);
      const back = await saveAndReopen(id, dir);
      assert.equal(back.losses?.deferredFeatures, 1);
      assert.equal(back.pivots, 0);
    } finally {
      closeSession(id);
      await rm(dir, { recursive: true, force: true });
    }
  });
}
