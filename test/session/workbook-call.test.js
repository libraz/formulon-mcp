import assert from "node:assert/strict";
import { test } from "vitest";
import {
  callWorkbookMethod,
  closeSession,
  listSessions,
  openSession,
} from "../../dist/sessions.js";

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
