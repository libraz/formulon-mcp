import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import {
  applySessionMutations,
  closeSession,
  getSessionCellByA1,
  listSessions,
  openSession,
  recalcSession,
  saveSession,
} from "../../dist/sessions.js";

test("manages session lifecycle", async () => {
  const session = await openSession(undefined, "lifecycle");
  assert.equal(session.id, "lifecycle");
  assert.equal(
    listSessions().some((item) => item.id === "lifecycle"),
    true,
  );
  const closed = closeSession("lifecycle");
  assert.equal(closed.id, "lifecycle");
  assert.equal(
    listSessions().some((item) => item.id === "lifecycle"),
    false,
  );
});

test("concurrent opens register one session and reject the duplicate id", async () => {
  const results = await Promise.allSettled([
    openSession(undefined, "concurrent-open"),
    openSession(undefined, "concurrent-open"),
  ]);
  try {
    assert.equal(results.filter((result) => result.status === "fulfilled").length, 1);
    assert.equal(results.filter((result) => result.status === "rejected").length, 1);
    assert.equal(listSessions().filter((session) => session.id === "concurrent-open").length, 1);
  } finally {
    if (listSessions().some((session) => session.id === "concurrent-open")) {
      closeSession("concurrent-open");
    }
  }
});

test("isolates multiple sessions and rejects access after close", async () => {
  await openSession(undefined, "multi-a");
  await openSession(undefined, "multi-b");
  try {
    applySessionMutations("multi-a", [{ type: "number", a1: "Sheet1!A1", value: 1 }], true);
    applySessionMutations("multi-b", [{ type: "number", a1: "Sheet1!A1", value: 2 }], true);
    assert.deepEqual(getSessionCellByA1("multi-a", "Sheet1!A1").value, {
      kind: "number",
      value: 1,
    });
    assert.deepEqual(getSessionCellByA1("multi-b", "Sheet1!A1").value, {
      kind: "number",
      value: 2,
    });
    closeSession("multi-a");
    assert.throws(() => getSessionCellByA1("multi-a", "Sheet1!A1"), /session not found/);
  } finally {
    if (listSessions().some((session) => session.id === "multi-a")) {
      closeSession("multi-a");
    }
    if (listSessions().some((session) => session.id === "multi-b")) {
      closeSession("multi-b");
    }
  }
});

test("recalculates sessions explicitly", async () => {
  await openSession(undefined, "recalc");
  try {
    applySessionMutations(
      "recalc",
      [
        { type: "number", a1: "Sheet1!A1", value: 5 },
        { type: "formula", a1: "Sheet1!B1", formula: "=A1*2" },
      ],
      false,
    );
    const result = recalcSession("recalc");
    assert.equal(result.status.ok, true);
    assert.deepEqual(getSessionCellByA1("recalc", "Sheet1!B1").value, {
      kind: "number",
      value: 10,
    });
  } finally {
    closeSession("recalc");
  }
});

test("recalculates formulas on open so the first read is not blank", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-open-recalc-"));
  const file = path.join(dir, "open-recalc.xlsx");
  await openSession(undefined, "open-recalc-src");
  try {
    applySessionMutations(
      "open-recalc-src",
      [
        { type: "number", a1: "Sheet1!A1", value: 21 },
        { type: "formula", a1: "Sheet1!B1", formula: "=A1*2" },
      ],
      true,
    );
    await saveSession("open-recalc-src", file);
  } finally {
    closeSession("open-recalc-src");
  }

  await openSession(file, "open-recalc-dst");
  try {
    // No explicit recalc_session: recalc-on-open must make B1 read 42, not blank.
    assert.deepEqual(getSessionCellByA1("open-recalc-dst", "Sheet1!B1").value, {
      kind: "number",
      value: 42,
    });
  } finally {
    closeSession("open-recalc-dst");
    await rm(dir, { recursive: true, force: true });
  }
});
