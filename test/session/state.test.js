import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { createOrLoadWorkbook, valueToJson } from "../../dist/formulon.js";
import {
  applySessionMutations,
  callWorkbookMethod,
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

test("opens a session under an Excel profile", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-profile-"));
  const file = path.join(dir, "profile.xlsx");
  const profileOf = (id) => callWorkbookMethod(id, "excelProfileId", []).result.value;
  const ascOf = (id) => {
    applySessionMutations(id, [{ type: "formula", a1: "A1", formula: '=ASC("ＡＢ")' }], true);
    return getSessionCellByA1(id, "A1").value.value;
  };
  try {
    await openSession(undefined, "profile-default");
    assert.equal(profileOf("profile-default"), "win-365-en_US");
    // ASC folds width only on the double-byte profiles.
    assert.equal(ascOf("profile-default"), "ＡＢ");
    await saveSession("profile-default", file);
    closeSession("profile-default");

    await openSession(undefined, "profile-new", "win-365-ja_JP");
    assert.equal(profileOf("profile-new"), "win-365-ja_JP");
    assert.equal(ascOf("profile-new"), "AB");
    closeSession("profile-new");

    await openSession(file, "profile-loaded", "mac-365-ja_JP");
    assert.equal(profileOf("profile-loaded"), "mac-365-ja_JP");
    assert.equal(getSessionCellByA1("profile-loaded", "A1").value.value, "AB");
    closeSession("profile-loaded");

    await assert.rejects(
      openSession(undefined, "profile-bad", "win-365-xx_XX"),
      /unknown excel profile/,
    );
    await assert.rejects(openSession(file, "profile-bad", "win-365-xx_XX"), /excel profile/);
    assert.equal(
      listSessions().some((session) => session.id === "profile-bad"),
      false,
    );
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("falls back to FORMULON_MCP_PROFILE when no profile is requested", async () => {
  const previous = process.env.FORMULON_MCP_PROFILE;
  try {
    process.env.FORMULON_MCP_PROFILE = "win-365-ja_JP";
    await openSession(undefined, "profile-env");
    assert.equal(
      callWorkbookMethod("profile-env", "excelProfileId", []).result.value,
      "win-365-ja_JP",
    );
    closeSession("profile-env");
    // An explicit profile still wins over the environment.
    await openSession(undefined, "profile-env-explicit", "mac-365-de_DE");
    assert.equal(
      callWorkbookMethod("profile-env-explicit", "excelProfileId", []).result.value,
      "mac-365-de_DE",
    );
    closeSession("profile-env-explicit");

    process.env.FORMULON_MCP_PROFILE = "win-365-xx_XX";
    await assert.rejects(
      openSession(undefined, "profile-env-bad"),
      /unknown excel profile: win-365-xx_XX \(from FORMULON_MCP_PROFILE\)/,
    );
  } finally {
    if (previous === undefined) {
      delete process.env.FORMULON_MCP_PROFILE;
    } else {
      process.env.FORMULON_MCP_PROFILE = previous;
    }
  }
});

// ---------------------------------------------------------------------------
// Excel profile resolution matrix: requested x environment x source x entry
// point (20 rows covering every value pair).
// ---------------------------------------------------------------------------

const PROFILE_ENV = "FORMULON_MCP_PROFILE";
const DEFAULT_PROFILE = "win-365-en_US";
const ASC_FORMULA = '=ASC("ＡＢ")';

const REQUESTED = {
  omitted: undefined,
  ja: "win-365-ja_JP",
  macde: "mac-365-de_DE",
  unknown: "win-365-xx_XX",
};
const ENV_VALUES = {
  unset: undefined,
  valid: "mac-365-ja_JP",
  unknown: "win-365-xx_XX",
  empty: "",
  whitespace: " \t ",
};

/** What a new workbook looks like in each profile the matrix can land on. */
const NEW_WORKBOOK_FACTS = {
  default: { sheet: "Sheet1", fontSize: 11 },
  "win-365-ja_JP": { sheet: "Sheet1", fontSize: 11, fontName: "游ゴシック" },
  "mac-365-ja_JP": { sheet: "Sheet1", fontSize: 12, fontName: "游ゴシック" },
  "mac-365-de_DE": { sheet: "Tabelle1", fontSize: 12 },
};

const PROFILE_CASES = [
  ["createOrLoadWorkbook", "macde", "empty", "loaded"],
  ["openSession", "ja", "unknown", "new"],
  ["createOrLoadWorkbook", "ja", "whitespace", "loaded"],
  ["openSession", "macde", "whitespace", "new"],
  ["createOrLoadWorkbook", "unknown", "valid", "new"],
  ["openSession", "omitted", "valid", "loaded"],
  ["createOrLoadWorkbook", "omitted", "unknown", "loaded"],
  ["createOrLoadWorkbook", "ja", "unset", "loaded"],
  ["openSession", "unknown", "empty", "new"],
  ["createOrLoadWorkbook", "unknown", "unknown", "loaded"],
  ["openSession", "macde", "unset", "loaded"],
  ["openSession", "omitted", "unset", "new"],
  ["openSession", "macde", "unknown", "new"],
  ["openSession", "macde", "valid", "new"],
  ["openSession", "unknown", "unset", "loaded"],
  ["openSession", "ja", "empty", "loaded"],
  ["createOrLoadWorkbook", "unknown", "whitespace", "new"],
  ["createOrLoadWorkbook", "omitted", "empty", "new"],
  ["openSession", "ja", "valid", "loaded"],
  ["createOrLoadWorkbook", "omitted", "whitespace", "loaded"],
  // Beyond the pairwise rows: the remaining source/entry pairings of the
  // environment-only paths.
  ["openSession", "omitted", "unknown", "new"],
  ["openSession", "omitted", "unknown", "loaded"],
  ["createOrLoadWorkbook", "omitted", "valid", "new"],
  ["createOrLoadWorkbook", "omitted", "valid", "loaded"],
  ["openSession", "omitted", "whitespace", "new"],
  ["openSession", "omitted", "empty", "loaded"],
];

/** Runs `fn` with FORMULON_MCP_PROFILE set (or deleted) and restores it afterwards. */
async function withProfileEnv(value, fn) {
  const previous = process.env[PROFILE_ENV];
  try {
    if (value === undefined) {
      delete process.env[PROFILE_ENV];
    } else {
      process.env[PROFILE_ENV] = value;
    }
    return await fn();
  } finally {
    if (previous === undefined) {
      delete process.env[PROFILE_ENV];
    } else {
      process.env[PROFILE_ENV] = previous;
    }
  }
}

/** An explicit profile wins; otherwise the trimmed environment value, if any. */
function resolveProfile(requestedKey, envKey) {
  const requested = REQUESTED[requestedKey];
  const fromEnvironment = ENV_VALUES[envKey]?.trim() || undefined;
  const chosen = requested ?? fromEnvironment;
  return {
    chosen,
    fromEnv: requested === undefined && fromEnvironment !== undefined,
    id: chosen ?? DEFAULT_PROFILE,
    known: chosen === undefined || !chosen.endsWith("xx_XX"),
  };
}

/** Writes a workbook holding `=ASC("ＡＢ")` (cached under the default profile) to `file`. */
async function writeAscWorkbook(file) {
  await openSession(undefined, "profile-fixture");
  try {
    applySessionMutations(
      "profile-fixture",
      [{ type: "formula", a1: "A1", formula: ASC_FORMULA }],
      true,
    );
    await saveSession("profile-fixture", file);
  } finally {
    closeSession("profile-fixture");
  }
}

function assertNewWorkbookFacts({ sheet, font }, chosen) {
  const facts = NEW_WORKBOOK_FACTS[chosen ?? "default"];
  assert.equal(sheet, facts.sheet);
  assert.equal(font.size, facts.fontSize);
  if (facts.fontName) {
    assert.equal(font.name, facts.fontName);
  }
  if (chosen !== undefined) {
    // A profile-shaped workbook never carries the engine's profile-less default font.
    assert.notEqual(font.name, "Calibri");
  } else {
    assert.equal(font.name, "Calibri");
  }
}

for (const row of PROFILE_CASES) {
  const [api, requestedKey, envKey, source] = row;
  test(`resolves the Excel profile: ${row.join(" / ")}`, async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-profile-matrix-"));
    const file = path.join(dir, "asc.xlsx");
    const id = `profile-matrix-${row.join("-")}`;
    const expected = resolveProfile(requestedKey, envKey);
    const requested = REQUESTED[requestedKey];
    const input = source === "loaded" ? file : undefined;
    const folds = expected.id.endsWith("ja_JP");
    try {
      if (source === "loaded") {
        await writeAscWorkbook(file);
      }
      await withProfileEnv(ENV_VALUES[envKey], async () => {
        const open =
          api === "openSession"
            ? () => openSession(input, id, requested)
            : () => createOrLoadWorkbook(input, requested);

        if (!expected.known) {
          const message =
            source === "new"
              ? `unknown excel profile: ${expected.chosen}`
              : `set excel profile ${expected.chosen}`;
          const error = await open().then(
            () => assert.fail("an unknown profile must be rejected"),
            (caught) => caught,
          );
          assert.ok(
            error.message.startsWith(
              expected.fromEnv ? `${message} (from ${PROFILE_ENV})` : message,
            ),
            error.message,
          );
          assert.equal(error.message.includes(PROFILE_ENV), expected.fromEnv, error.message);
          // A failed open registers nothing, and the id stays usable.
          assert.equal(
            listSessions().some((session) => session.id === id),
            false,
          );
          if (api === "openSession") {
            await openSession(undefined, id, "win-365-en_US");
            closeSession(id);
          }
          return;
        }

        if (api === "openSession") {
          const session = await openSession(input, id, requested);
          assert.equal(session.id, id);
          try {
            assert.equal(callWorkbookMethod(id, "excelProfileId", []).result.value, expected.id);
            if (source === "new") {
              applySessionMutations(
                id,
                [{ type: "formula", a1: "A1", formula: ASC_FORMULA }],
                true,
              );
              assertNewWorkbookFacts(
                {
                  sheet: callWorkbookMethod(id, "sheetName", [0]).result.value,
                  font: callWorkbookMethod(id, "getFont", [0]).result,
                },
                expected.chosen,
              );
            }
            // Opening recalculates, so a loaded workbook folds under ja_JP only.
            assert.equal(getSessionCellByA1(id, "A1").value.value, folds ? "AB" : "ＡＢ");
          } finally {
            closeSession(id);
          }
          return;
        }

        const workbook = await open();
        try {
          assert.equal(workbook.excelProfileId().value, expected.id);
          if (source === "new") {
            assert.equal(workbook.setFormula(0, 0, 0, ASC_FORMULA).ok, true);
            assertNewWorkbookFacts(
              { sheet: workbook.sheetName(0).value, font: workbook.getFont(0) },
              expected.chosen,
            );
          }
          assert.equal(workbook.recalc().ok, true);
          assert.equal(valueToJson(workbook.getValue(0, 0, 0).value).value, folds ? "AB" : "ＡＢ");
        } finally {
          workbook.delete();
        }
      });
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
}

test("trims the environment profile and treats a blank one as unset", async () => {
  await withProfileEnv("  win-365-fr_FR\t", async () => {
    await openSession(undefined, "profile-env-trim");
    try {
      assert.equal(
        callWorkbookMethod("profile-env-trim", "excelProfileId", []).result.value,
        "win-365-fr_FR",
      );
      assert.equal(callWorkbookMethod("profile-env-trim", "sheetName", [0]).result.value, "Feuil1");
    } finally {
      closeSession("profile-env-trim");
    }
  });
  for (const blank of ["", " ", "\t\n "]) {
    await withProfileEnv(blank, async () => {
      await openSession(undefined, "profile-env-blank");
      try {
        assert.equal(
          callWorkbookMethod("profile-env-blank", "excelProfileId", []).result.value,
          DEFAULT_PROFILE,
        );
        assert.equal(
          callWorkbookMethod("profile-env-blank", "getFont", [0]).result.name,
          "Calibri",
        );
      } finally {
        closeSession("profile-env-blank");
      }
    });
  }
});

test("reports a malformed environment profile as unknown rather than ignoring it", async () => {
  await withProfileEnv("ja_JP", async () => {
    await assert.rejects(
      openSession(undefined, "profile-env-malformed"),
      /^Error: unknown excel profile: ja_JP \(from FORMULON_MCP_PROFILE\)$/,
    );
    // An explicit well-formed profile still wins over the broken value.
    await openSession(undefined, "profile-env-malformed", "win-365-de_DE");
    assert.equal(
      callWorkbookMethod("profile-env-malformed", "excelProfileId", []).result.value,
      "win-365-de_DE",
    );
    closeSession("profile-env-malformed");
  });
});

test("does not persist the profile in a saved workbook", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-profile-persist-"));
  const file = path.join(dir, "persist.xlsx");
  try {
    await withProfileEnv(undefined, async () => {
      await openSession(undefined, "profile-persist", "win-365-ja_JP");
      applySessionMutations("profile-persist", [{ type: "number", a1: "A1", value: 1 }], true);
      await saveSession("profile-persist", file);
      closeSession("profile-persist");

      await openSession(file, "profile-persist-reopen");
      assert.equal(
        callWorkbookMethod("profile-persist-reopen", "excelProfileId", []).result.value,
        DEFAULT_PROFILE,
      );
      closeSession("profile-persist-reopen");
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
