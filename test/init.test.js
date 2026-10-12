import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdir, mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { platform, tmpdir } from "node:os";
import path from "node:path";
import { PassThrough } from "node:stream";
import { fileURLToPath } from "node:url";
import { test } from "vitest";
import { writeJsonConfig } from "../dist/init/config.js";
import { createPrompter } from "../dist/init/prompt.js";
import {
  assertKnownProfile,
  defaultLocaleKey,
  PROFILE_LOCALES,
  parseProfileChoice,
  parseTargetChoice,
  previewRemoveImpact,
  previewWriteImpact,
  removeFromClaudeConfig,
  removeFromCodexConfig,
  writeClaudeConfig,
  writeClaudeDesktopConfig,
  writeCodexConfig,
} from "../dist/init.js";

const makeTmp = async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-init-"));
  return {
    dir,
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
};

test("writeClaudeConfig creates a fresh JSON config with the formulon entry", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, "nested", ".claude.json");
    assert.equal(await previewWriteImpact(file), "(new)");

    await writeClaudeConfig(file);
    const parsed = JSON.parse(await readFile(file, "utf8"));
    assert.deepEqual(parsed.mcpServers.formulon, {
      command: "npx",
      args: ["-y", "@libraz/formulon-mcp"],
    });
  } finally {
    await cleanup();
  }
});

test("writeClaudeConfig preserves existing keys and other MCP servers", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, ".claude.json");
    const initial = {
      somethingElse: { keepMe: true },
      mcpServers: {
        other: { command: "node", args: ["./other.js"] },
      },
    };
    await writeFile(file, JSON.stringify(initial, null, 2));
    assert.equal(await previewWriteImpact(file), "(merge)");

    await writeClaudeConfig(file);
    const parsed = JSON.parse(await readFile(file, "utf8"));
    assert.deepEqual(parsed.somethingElse, { keepMe: true });
    assert.deepEqual(parsed.mcpServers.other, { command: "node", args: ["./other.js"] });
    assert.deepEqual(parsed.mcpServers.formulon, {
      command: "npx",
      args: ["-y", "@libraz/formulon-mcp"],
    });

    assert.equal(await previewWriteImpact(file), "(replace formulon)");
  } finally {
    await cleanup();
  }
});

test("writeCodexConfig appends a TOML block without touching other sections", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, ".codex", "config.toml");

    await writeCodexConfig(file);
    let content = await readFile(file, "utf8");
    assert.match(content, /\[mcp_servers\.formulon\]/);
    assert.match(content, /command = "npx"/);
    assert.match(content, /args = \["-y", "@libraz\/formulon-mcp"\]/);

    const seeded = '[mcp_servers.other]\ncommand = "node"\nargs = ["./other.js"]\n';
    await writeFile(file, seeded);
    assert.equal(await previewWriteImpact(file), "(merge)");
    await writeCodexConfig(file);
    content = await readFile(file, "utf8");
    assert.match(content, /\[mcp_servers\.other\]/);
    assert.match(content, /\[mcp_servers\.formulon\]/);
    assert.equal(await previewWriteImpact(file), "(replace formulon)");

    await writeCodexConfig(file);
    const occurrences = (await readFile(file, "utf8")).match(/\[mcp_servers\.formulon\]/g) ?? [];
    assert.equal(occurrences.length, 1);
  } finally {
    await cleanup();
  }
});

test("removeFromClaudeConfig drops only the formulon entry", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, ".claude.json");
    await writeFile(
      file,
      JSON.stringify({
        mcpServers: {
          other: { command: "node", args: ["./other.js"] },
          formulon: { command: "npx", args: ["-y", "@libraz/formulon-mcp"] },
        },
      }),
    );
    assert.equal(await previewRemoveImpact(file), "(remove formulon)");

    assert.equal(await removeFromClaudeConfig(file), "removed");
    const parsed = JSON.parse(await readFile(file, "utf8"));
    assert.equal(parsed.mcpServers.formulon, undefined);
    assert.deepEqual(parsed.mcpServers.other, { command: "node", args: ["./other.js"] });

    assert.equal(await removeFromClaudeConfig(file), "absent");
    assert.equal(await removeFromClaudeConfig(path.join(dir, "missing.json")), "no-file");
  } finally {
    await cleanup();
  }
});

test("removeFromCodexConfig keeps other TOML sections intact", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, "config.toml");
    await writeFile(
      file,
      [
        "[mcp_servers.other]",
        'command = "node"',
        'args = ["./other.js"]',
        "",
        "[mcp_servers.formulon]",
        'command = "npx"',
        'args = ["-y", "@libraz/formulon-mcp"]',
        "",
        "[other_section]",
        'value = "kept"',
        "",
      ].join("\n"),
    );

    assert.equal(await removeFromCodexConfig(file), "removed");
    const content = await readFile(file, "utf8");
    assert.match(content, /\[mcp_servers\.other\]/);
    assert.match(content, /\[other_section\]/);
    assert.doesNotMatch(content, /\[mcp_servers\.formulon\]/);

    assert.equal(await removeFromCodexConfig(file), "absent");
    assert.equal(await removeFromCodexConfig(path.join(dir, "missing.toml")), "no-file");
  } finally {
    await cleanup();
  }
});

test("removeFromCodexConfig drops sub-tables under the formulon section", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, "config.toml");
    // A hand-added [mcp_servers.formulon.env] must go with its parent: left
    // behind, it is an orphan sub-table of a server that no longer exists.
    await writeFile(
      file,
      [
        "[mcp_servers.other]",
        'command = "node"',
        "",
        "[mcp_servers.formulon]",
        'command = "npx"',
        'args = ["-y", "@libraz/formulon-mcp"]',
        "",
        "[mcp_servers.formulon.env]",
        'FORMULON_MCP_PRETTY = "1"',
        "",
        "[other_section]",
        'value = "kept"',
        "",
      ].join("\n"),
    );
    assert.equal(await previewRemoveImpact(file), "(remove formulon)");

    assert.equal(await removeFromCodexConfig(file), "removed");
    const content = await readFile(file, "utf8");
    assert.doesNotMatch(content, /mcp_servers\.formulon/);
    assert.doesNotMatch(content, /FORMULON_MCP_PRETTY/);
    assert.match(content, /\[mcp_servers\.other\]/);
    assert.match(content, /\[other_section\]/);
  } finally {
    await cleanup();
  }
});

test("config writes leave no temp file behind", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, ".claude.json");
    await writeClaudeConfig(file);
    await removeFromClaudeConfig(file);
    const leftovers = (await readdir(dir)).filter((name) => name.endsWith(".tmp"));
    assert.deepEqual(leftovers, []);
  } finally {
    await cleanup();
  }
});

test("parseTargetChoice handles empty/explicit/invalid selections", () => {
  assert.deepEqual(parseTargetChoice("", "1,3"), ["claude-user", "codex"]);
  assert.deepEqual(parseTargetChoice("2", "1"), ["claude-project"]);
  assert.deepEqual(parseTargetChoice("4,1,4", "1"), ["claude-desktop", "claude-user"]);
  assert.throws(() => parseTargetChoice("9", "1"), /Invalid choice: 9/);
  assert.throws(() => parseTargetChoice(" , ", "1"), /No targets selected/);
});

test("init writes the default Excel profile into each client's env", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const json = path.join(dir, ".claude.json");
    await writeClaudeConfig(json, { profile: "win-365-ja_JP" });
    assert.deepEqual(JSON.parse(await readFile(json, "utf8")).mcpServers.formulon, {
      command: "npx",
      args: ["-y", "@libraz/formulon-mcp"],
      env: { FORMULON_MCP_PROFILE: "win-365-ja_JP" },
    });
    // Re-running without a profile writes the entry without one.
    await writeClaudeConfig(json);
    assert.equal(JSON.parse(await readFile(json, "utf8")).mcpServers.formulon.env, undefined);

    const toml = path.join(dir, "config.toml");
    await writeCodexConfig(toml, { profile: "mac-365-ja_JP" });
    assert.match(
      await readFile(toml, "utf8"),
      /\[mcp_servers\.formulon\]\ncommand = "npx"\nargs = \["-y", "@libraz\/formulon-mcp"\]\nenv = \{ FORMULON_MCP_PROFILE = "mac-365-ja_JP" \}\n$/,
    );
  } finally {
    await cleanup();
  }
});

test("init rejects a profile the engine does not know", async () => {
  await assertKnownProfile("win-365-ja_JP");
  await assert.rejects(assertKnownProfile("win-365-xx_XX"), /Unknown Excel profile/);
  await assert.rejects(assertKnownProfile("ja_JP"), /Unknown Excel profile/);
});

test("every profile the init menu offers is known to the engine", async () => {
  for (const { code } of PROFILE_LOCALES) {
    for (const host of ["win", "mac"]) {
      await assertKnownProfile(`${host}-365-${code}`);
    }
  }
});

test("maps the locale and host menu answers to a profile id", () => {
  const defaults = { locale: "2", host: "1" };
  assert.equal(parseProfileChoice("", "", defaults), "win-365-ja_JP");
  assert.equal(parseProfileChoice("1", "2", defaults), "mac-365-en_US");
  assert.equal(parseProfileChoice("0", "", defaults), undefined);
  assert.equal(parseProfileChoice("", "", { locale: "0", host: "1" }), undefined);
  assert.throws(() => parseProfileChoice("99", "", defaults), /Invalid choice: 99/);
  assert.throws(() => parseProfileChoice("ja", "", defaults), /Invalid choice: ja/);
  assert.throws(() => parseProfileChoice("2", "3", defaults), /Invalid choice: 3/);
});

test("defaults the locale menu to the system locale", () => {
  assert.equal(defaultLocaleKey("ja-JP"), "2");
  assert.equal(defaultLocaleKey("en-US"), "1");
  // A region the engine has no profile for falls back to the language.
  assert.equal(defaultLocaleKey("de-AT"), "3");
  assert.equal(defaultLocaleKey("ar-SA"), "0");
});

test("answers piped in at once reach every question", async () => {
  const { PassThrough } = await import("node:stream");
  const { createPrompter } = await import("../dist/init/prompt.js");
  const input = new PassThrough();
  const output = new PassThrough();
  const prompter = createPrompter(input, output);
  input.end("1\n2\ny\n");
  try {
    assert.equal(await prompter.ask("A: "), "1");
    assert.equal(await prompter.ask("B: "), "2");
    assert.equal(await prompter.ask("C: "), "y");
    // Input that runs out fails the question instead of leaving it pending.
    await assert.rejects(prompter.ask("D: "), /Input ended before an answer to: D:/);
  } finally {
    prompter.close();
  }
});

// ---------------------------------------------------------------------------
// Profile menu answers
// ---------------------------------------------------------------------------

const LOCALE_CODES = [
  "en_US",
  "ja_JP",
  "de_DE",
  "fr_FR",
  "es_ES",
  "es_MX",
  "pt_BR",
  "it_IT",
  "nl_NL",
  "ru_RU",
  "zh_CN",
  "zh_TW",
  "ko_KR",
  "th_TH",
];

test("the locale menu lists every supported locale in order", () => {
  assert.deepEqual(
    PROFILE_LOCALES.map((locale) => locale.code),
    LOCALE_CODES,
  );
});

test("maps every locale key and host key to its profile id", () => {
  const hosts = { 1: "win", 2: "mac" };
  for (const [index, code] of LOCALE_CODES.entries()) {
    for (const [hostKey, host] of Object.entries(hosts)) {
      const expected = `${host}-365-${code}`;
      const defaults = { locale: "0", host: "1" };
      assert.equal(parseProfileChoice(String(index + 1), hostKey, defaults), expected);
      // Surrounding whitespace is ignored on both answers.
      assert.equal(parseProfileChoice(` ${index + 1}\t`, ` ${hostKey} `, defaults), expected);
      // Blank answers take the defaults.
      assert.equal(
        parseProfileChoice("", "", { locale: String(index + 1), host: hostKey }),
        expected,
      );
      assert.equal(
        parseProfileChoice(String(index + 1), "", { ...defaults, host: hostKey }),
        expected,
      );
      assert.equal(
        parseProfileChoice("", hostKey, { ...defaults, locale: String(index + 1) }),
        expected,
      );
    }
  }
});

const PROFILE_ANSWERS = {
  localeRaw: {
    blank: "",
    ws: " \t",
    3: "3",
    padded3: " 3 ",
    0: "0",
    15: "15",
    "-1": "-1",
    "1.0": "1.0",
    ja: "ja",
    "01": "01",
  },
  hostRaw: { blank: "", 1: "1", padded2: " 2 ", 3: "3", x: "x" },
};

/**
 * Pairwise rows of (locale answer, host answer, default locale, default host);
 * 50 rows covering every value pair.
 */
const PROFILE_CHOICE_CASES = [
  ["0", "1", "99", "2"],
  ["ws", "x", "0", "1"],
  ["15", "1", "0", "2"],
  ["0", "padded2", "2", "1"],
  ["-1", "3", "2", "2"],
  ["padded3", "padded2", "0", "2"],
  ["1.0", "blank", "99", "1"],
  ["blank", "padded2", "99", "1"],
  ["15", "3", "99", "1"],
  ["ja", "x", "99", "1"],
  ["1.0", "x", "2", "2"],
  ["ws", "blank", "2", "2"],
  ["ja", "padded2", "0", "2"],
  ["blank", "blank", "0", "2"],
  ["01", "1", "2", "1"],
  ["-1", "1", "99", "1"],
  ["3", "padded2", "2", "2"],
  ["padded3", "blank", "99", "1"],
  ["3", "3", "0", "1"],
  ["-1", "padded2", "0", "1"],
  ["blank", "3", "2", "2"],
  ["0", "3", "0", "2"],
  ["padded3", "1", "2", "2"],
  ["ws", "padded2", "99", "1"],
  ["1.0", "padded2", "0", "1"],
  ["01", "x", "0", "2"],
  ["01", "3", "99", "1"],
  ["ja", "1", "2", "2"],
  ["3", "x", "99", "2"],
  ["1.0", "1", "2", "2"],
  ["-1", "blank", "0", "2"],
  ["padded3", "3", "2", "2"],
  ["padded3", "x", "99", "1"],
  ["3", "blank", "99", "2"],
  ["ja", "3", "0", "1"],
  ["1.0", "3", "99", "1"],
  ["3", "1", "2", "2"],
  ["15", "padded2", "0", "1"],
  ["ja", "blank", "0", "2"],
  ["0", "x", "99", "1"],
  ["ws", "3", "2", "1"],
  ["blank", "1", "0", "2"],
  ["-1", "x", "0", "1"],
  ["01", "padded2", "2", "1"],
  ["15", "x", "2", "1"],
  ["blank", "x", "0", "1"],
  ["15", "blank", "99", "2"],
  ["ws", "1", "0", "2"],
  ["0", "blank", "99", "1"],
  ["01", "blank", "99", "2"],
];

/** The documented rules, written out without the implementation's helpers. */
function expectedProfileChoice(localeRaw, hostRaw, defaults) {
  const localeKey = localeRaw.trim() || defaults.locale;
  if (localeKey === "0") {
    return { profile: undefined };
  }
  const number = /^[0-9]+$/.test(localeKey) ? Number(localeKey) : 0;
  if (number < 1 || number > LOCALE_CODES.length) {
    return { error: `Invalid choice: ${localeKey}` };
  }
  const hostKey = hostRaw.trim() || defaults.host;
  if (hostKey !== "1" && hostKey !== "2") {
    return { error: `Invalid choice: ${hostKey}` };
  }
  return { profile: `${hostKey === "1" ? "win" : "mac"}-365-${LOCALE_CODES[number - 1]}` };
}

for (const [localeKey, hostKey, defaultLocale, defaultHost] of PROFILE_CHOICE_CASES) {
  test(`parses the profile answers: locale ${localeKey}, host ${hostKey}, defaults ${defaultLocale}/${defaultHost}`, () => {
    const localeRaw = PROFILE_ANSWERS.localeRaw[localeKey];
    const hostRaw = PROFILE_ANSWERS.hostRaw[hostKey];
    const defaults = { locale: defaultLocale, host: defaultHost };
    const expected = expectedProfileChoice(localeRaw, hostRaw, defaults);
    if (expected.error) {
      assert.throws(
        () => parseProfileChoice(localeRaw, hostRaw, defaults),
        new RegExp(`^Error: ${expected.error.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`),
      );
    } else {
      assert.equal(parseProfileChoice(localeRaw, hostRaw, defaults), expected.profile);
    }
  });
}

test("the engine-default locale ignores the host answer, even an invalid one", () => {
  for (const host of ["", "1", "2", "3", "mac", "99"]) {
    assert.equal(parseProfileChoice("0", host, { locale: "2", host: "1" }), undefined);
    assert.equal(parseProfileChoice(" 0 ", host, { locale: "2", host: "9" }), undefined);
  }
});

test("rejects locale answers that are not a whole menu number", () => {
  const defaults = { locale: "2", host: "1" };
  for (const answer of ["15", "-1", "1.0", "ja", "+3", "0x3", "1e0", "00", "3,4", "٣"]) {
    assert.throws(
      () => parseProfileChoice(answer, "1", defaults),
      new RegExp(`^Error: Invalid choice: ${answer.replace(/[.+*?^${}()|[\]\\]/g, "\\$&")}$`),
      answer,
    );
  }
  // Leading zeros are still a number.
  assert.equal(parseProfileChoice("01", "1", defaults), "win-365-en_US");
  assert.equal(parseProfileChoice("014", "2", defaults), "mac-365-th_TH");
  for (const host of ["0", "3", "01", "win", "x", "1.0"]) {
    assert.throws(
      () => parseProfileChoice("2", host, defaults),
      new RegExp(`^Error: Invalid choice: ${host.replace(".", "\\.")}$`),
      host,
    );
  }
});

test("an invalid default is reported when the answer is blank", () => {
  assert.throws(
    () => parseProfileChoice("", "1", { locale: "99", host: "1" }),
    /Invalid choice: 99/,
  );
  assert.throws(() => parseProfileChoice("2", "", { locale: "2", host: "9" }), /Invalid choice: 9/);
});

test("accepts every menu locale on both hosts as a known profile", async () => {
  for (const code of LOCALE_CODES) {
    for (const host of ["win", "mac"]) {
      await assertKnownProfile(`${host}-365-${code}`);
    }
  }
  for (const bad of [
    "",
    "ja_JP",
    "win-365-ja-JP",
    "win-365-xx_XX",
    "win-364-ja_JP",
    "linux-365-ja_JP",
    "WIN-365-ja_JP",
    " win-365-ja_JP",
  ]) {
    await assert.rejects(assertKnownProfile(bad), /Unknown Excel profile/, JSON.stringify(bad));
  }
});

const SYSTEM_LOCALES = [
  ["ja-JP", "2"],
  ["ja_JP", "2"],
  ["ja", "2"],
  ["en-US", "1"],
  ["en-GB", "1"],
  ["en", "1"],
  ["de-DE", "3"],
  ["de-AT", "3"],
  ["de-CH", "3"],
  ["fr-FR", "4"],
  ["fr-CA", "4"],
  ["es-ES", "5"],
  ["es-AR", "5"],
  ["es-419", "5"],
  ["es-MX", "6"],
  ["pt-BR", "7"],
  ["pt-PT", "7"],
  ["it-IT", "8"],
  ["it-CH", "8"],
  ["nl-NL", "9"],
  ["nl-BE", "9"],
  ["ru-RU", "10"],
  ["ru-UA", "10"],
  ["zh-CN", "11"],
  ["zh-TW", "12"],
  // Without a region match the script decides: Hong Kong and Macau write Traditional.
  ["zh-HK", "12"],
  ["zh-MO", "12"],
  ["zh-SG", "11"],
  ["zh-Hant", "12"],
  ["zh-Hant-TW", "12"],
  ["zh-Hans", "11"],
  ["zh", "11"],
  ["ko-KR", "13"],
  ["ko", "13"],
  ["th-TH", "14"],
  ["th", "14"],
  // POSIX spellings, as LANG carries them, match like their BCP 47 forms.
  ["ja_JP.UTF-8", "2"],
  ["de_DE.UTF-8@euro", "3"],
  ["zh_HK", "12"],
  // Unicode extensions do not change the match.
  ["en-US-u-ca-gregory", "1"],
  ["ja-JP-u-ca-japanese", "2"],
  // No profile for these: the engine default is offered.
  ["ar-SA", "0"],
  ["hi-IN", "0"],
  ["sv-SE", "0"],
  ["C", "0"],
  ["POSIX", "0"],
  ["und", "0"],
  ["", "0"],
  ["x", "0"],
];

test("picks the menu default from a BCP 47 system locale", () => {
  for (const [tag, key] of SYSTEM_LOCALES) {
    assert.equal(defaultLocaleKey(tag), key, JSON.stringify(tag));
  }
  // Every menu locale is its own default, spelled with either separator.
  for (const [index, code] of LOCALE_CODES.entries()) {
    assert.equal(defaultLocaleKey(code), String(index + 1));
    assert.equal(defaultLocaleKey(code.replace("_", "-")), String(index + 1));
  }
});

// ---------------------------------------------------------------------------
// Config writers with a profile
// ---------------------------------------------------------------------------

const PROFILE_OPTIONS = {
  none: undefined,
  winJa: "win-365-ja_JP",
  macDe: "mac-365-de_DE",
  // An empty profile is no profile.
  empty: "",
};

const JSON_WRITERS = {
  claude: writeClaudeConfig,
  desktop: writeClaudeDesktopConfig,
  writeJson: writeJsonConfig,
};

const FORMULON_ARGS = ["-y", "@libraz/formulon-mcp"];

const OTHER_SERVER = { command: "node", args: ["./other.js"], env: { KEEP: "me" } };

/** Existing file contents (as text) and what must survive a rewrite of them. */
const JSON_STARTS = {
  absent: { text: undefined, kept: {} },
  empty: { text: "", kept: {} },
  noServers: {
    text: JSON.stringify({ theme: "dark", history: [1, 2, { nested: true }] }),
    kept: { theme: "dark", history: [1, 2, { nested: true }] },
  },
  otherServers: {
    text: JSON.stringify({ keep: { a: 1 }, mcpServers: { other: OTHER_SERVER } }),
    kept: { keep: { a: 1 }, other: OTHER_SERVER },
  },
  formulonPlain: {
    text: JSON.stringify({
      mcpServers: {
        formulon: { command: "npx", args: FORMULON_ARGS },
        other: OTHER_SERVER,
      },
      top: 1,
    }),
    kept: { other: OTHER_SERVER, top: 1 },
  },
  formulonWithEnv: {
    text: JSON.stringify({
      mcpServers: {
        other: OTHER_SERVER,
        formulon: {
          command: "npx",
          args: FORMULON_ARGS,
          env: { FORMULON_MCP_PROFILE: "mac-365-ja_JP", EXTRA: "x" },
        },
      },
    }),
    kept: { other: OTHER_SERVER },
  },
};

/** Pairwise rows of (writer, profile, existing file); 25 rows covering every value pair. */
const JSON_WRITER_CASES = [
  ["desktop", "none", "empty"],
  ["writeJson", "none", "otherServers"],
  ["claude", "macDe", "otherServers"],
  ["desktop", "winJa", "otherServers"],
  ["desktop", "macDe", "formulonWithEnv"],
  ["claude", "empty", "absent"],
  ["desktop", "empty", "absent"],
  ["writeJson", "winJa", "empty"],
  ["writeJson", "macDe", "formulonPlain"],
  ["claude", "winJa", "formulonWithEnv"],
  ["writeJson", "macDe", "absent"],
  ["desktop", "macDe", "noServers"],
  ["claude", "none", "absent"],
  ["writeJson", "empty", "noServers"],
  ["claude", "winJa", "noServers"],
  ["claude", "empty", "formulonPlain"],
  ["desktop", "none", "formulonPlain"],
  ["writeJson", "none", "noServers"],
  ["desktop", "winJa", "absent"],
  ["desktop", "empty", "otherServers"],
  ["claude", "none", "formulonWithEnv"],
  ["writeJson", "empty", "formulonWithEnv"],
  ["claude", "empty", "empty"],
  ["desktop", "winJa", "formulonPlain"],
  ["writeJson", "macDe", "empty"],
];

for (const [writerKey, profileKey, startKey] of JSON_WRITER_CASES) {
  test(`JSON config writer ${writerKey} with profile ${profileKey} over ${startKey}`, async () => {
    const { dir, cleanup } = await makeTmp();
    try {
      const file = path.join(dir, "nested", "config.json");
      const start = JSON_STARTS[startKey];
      if (start.text !== undefined) {
        await mkdir(path.dirname(file), { recursive: true });
        await writeFile(file, start.text);
      }
      const profile = PROFILE_OPTIONS[profileKey];
      await JSON_WRITERS[writerKey](file, profile === undefined ? undefined : { profile });

      const text = await readFile(file, "utf8");
      assert.ok(text.endsWith("}\n"), "ends with a newline");
      assert.equal(text, `${JSON.stringify(JSON.parse(text), null, 2)}\n`, "two-space JSON");
      const parsed = JSON.parse(text);
      // The entry is replaced as a whole: the launch command plus the profile, nothing older.
      assert.deepEqual(parsed.mcpServers.formulon, {
        command: "npx",
        args: FORMULON_ARGS,
        ...(profile ? { env: { FORMULON_MCP_PROFILE: profile } } : {}),
      });
      // Everything else is exactly as it was.
      const { formulon: _entry, ...otherServers } = parsed.mcpServers;
      const { mcpServers: _servers, ...otherKeys } = parsed;
      const { other, ...keptKeys } = start.kept;
      assert.deepEqual(otherKeys, keptKeys);
      assert.deepEqual(otherServers, other === undefined ? {} : { other });

      // Writing the same options again changes nothing.
      await JSON_WRITERS[writerKey](file, profile === undefined ? undefined : { profile });
      assert.equal(await readFile(file, "utf8"), text);
      assert.deepEqual(
        (await readdir(path.dirname(file))).filter((name) => name.endsWith(".tmp")),
        [],
      );

      // Uninstalling removes the entry with its env and nothing else.
      assert.equal(await removeFromClaudeConfig(file), "removed");
      const after = await readFile(file, "utf8");
      assert.doesNotMatch(after, /FORMULON_MCP_PROFILE/);
      const reparsed = JSON.parse(after);
      assert.equal(reparsed.mcpServers.formulon, undefined);
      assert.deepEqual(reparsed.mcpServers, other === undefined ? {} : { other });
    } finally {
      await cleanup();
    }
  });
}

test("a rewrite with a different profile replaces the old one", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, ".claude.json");
    await writeClaudeConfig(file, { profile: "win-365-ja_JP" });
    await writeClaudeConfig(file, { profile: "mac-365-de_DE" });
    assert.deepEqual(JSON.parse(await readFile(file, "utf8")).mcpServers.formulon.env, {
      FORMULON_MCP_PROFILE: "mac-365-de_DE",
    });
    await writeClaudeConfig(file, {});
    assert.equal(JSON.parse(await readFile(file, "utf8")).mcpServers.formulon.env, undefined);
  } finally {
    await cleanup();
  }
});

test("a JSON config that does not parse is left untouched", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, ".claude.json");
    await writeFile(file, '{ "mcpServers": { broken');
    await assert.rejects(writeClaudeConfig(file, { profile: "win-365-ja_JP" }), /Failed to parse/);
    await assert.rejects(removeFromClaudeConfig(file), /Failed to parse/);
    assert.equal(await readFile(file, "utf8"), '{ "mcpServers": { broken');
    assert.deepEqual(
      (await readdir(dir)).filter((name) => name.endsWith(".tmp")),
      [],
    );
  } finally {
    await cleanup();
  }
});

const CODEX_OTHER = [
  'model = "gpt"',
  "",
  "[mcp_servers.other]",
  'command = "node"',
  'args = ["./other.js"]',
  "",
  "[profiles.fast]",
  'model = "small"',
];

const CODEX_PLAIN = [
  "[mcp_servers.formulon]",
  'command = "npx"',
  'args = ["-y", "@libraz/formulon-mcp"]',
];

/** Existing TOML as `[kept text, formulon text]`, where the formulon part may sit anywhere. */
const TOML_STARTS = {
  absent: { before: undefined, formulon: [], after: [] },
  empty: { before: [], formulon: [], after: [] },
  otherTables: { before: CODEX_OTHER, formulon: [], after: [] },
  formulonPlain: { before: CODEX_OTHER, formulon: CODEX_PLAIN, after: [] },
  formulonInlineEnv: {
    before: CODEX_OTHER,
    formulon: [...CODEX_PLAIN, 'env = { FORMULON_MCP_PROFILE = "mac-365-ja_JP" }'],
    after: [],
  },
  formulonEnvSubtable: {
    before: CODEX_OTHER,
    formulon: [
      ...CODEX_PLAIN,
      "",
      "[mcp_servers.formulon.env]",
      'FORMULON_MCP_PROFILE = "mac-365-ja_JP"',
      'EXTRA = "1"',
    ],
    after: ["", "[other_section]", 'value = "kept"'],
  },
  formulonBetweenOthers: {
    before: ["[first]", 'a = "1"', ""],
    formulon: CODEX_PLAIN,
    after: ["", "[second]", 'b = "2"', "", "[[array_of_tables]]", 'c = "3"'],
  },
  lookalikeTables: {
    before: [
      "[mcp_servers.formulon-extra]",
      'command = "keep-1"',
      "",
      "[mcp_servers.formulon2]",
      'command = "keep-2"',
      "",
      "[mcp_servers.other.formulon]",
      'command = "keep-3"',
      "",
      "[mcp_servers.formulonx.env]",
      'KEEP = "4"',
    ],
    formulon: [],
    after: [],
  },
  crlf: {
    before: CODEX_OTHER,
    formulon: CODEX_PLAIN,
    after: ["", "[tail]", 'x = "1"'],
    eol: "\r\n",
  },
};

/** Pairwise rows of (profile, existing, writes); 36 rows covering every value pair. */
const TOML_WRITER_CASES = [
  ["empty", "otherTables", "profileChange"],
  ["macDe", "crlf", "profileChange"],
  ["winJa", "formulonEnvSubtable", "twice"],
  ["winJa", "crlf", "once"],
  ["macDe", "formulonPlain", "once"],
  ["none", "formulonBetweenOthers", "profileChange"],
  ["empty", "absent", "once"],
  ["winJa", "absent", "profileChange"],
  ["none", "formulonEnvSubtable", "once"],
  ["none", "empty", "twice"],
  ["empty", "lookalikeTables", "once"],
  ["macDe", "formulonBetweenOthers", "once"],
  ["macDe", "absent", "twice"],
  ["none", "lookalikeTables", "profileChange"],
  ["none", "crlf", "twice"],
  ["empty", "empty", "once"],
  ["winJa", "empty", "profileChange"],
  ["macDe", "lookalikeTables", "twice"],
  ["macDe", "empty", "twice"],
  ["empty", "formulonBetweenOthers", "twice"],
  ["macDe", "formulonInlineEnv", "twice"],
  ["winJa", "formulonInlineEnv", "profileChange"],
  ["macDe", "otherTables", "twice"],
  ["winJa", "lookalikeTables", "profileChange"],
  ["winJa", "otherTables", "once"],
  ["none", "absent", "once"],
  ["empty", "crlf", "once"],
  ["empty", "formulonEnvSubtable", "profileChange"],
  ["none", "formulonInlineEnv", "once"],
  ["macDe", "formulonEnvSubtable", "twice"],
  ["empty", "formulonPlain", "twice"],
  ["none", "otherTables", "twice"],
  ["none", "formulonPlain", "profileChange"],
  ["winJa", "formulonBetweenOthers", "profileChange"],
  ["empty", "formulonInlineEnv", "twice"],
  ["winJa", "formulonPlain", "once"],
];

const significantLines = (text) =>
  text
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

const formulonBlock = (profile) => [
  "[mcp_servers.formulon]",
  'command = "npx"',
  'args = ["-y", "@libraz/formulon-mcp"]',
  ...(profile ? [`env = { FORMULON_MCP_PROFILE = "${profile}" }`] : []),
];

for (const [profileKey, startKey, writes] of TOML_WRITER_CASES) {
  test(`Codex TOML writer with profile ${profileKey} over ${startKey}, ${writes}`, async () => {
    const { dir, cleanup } = await makeTmp();
    try {
      const file = path.join(dir, ".codex", "config.toml");
      const start = TOML_STARTS[startKey];
      if (start.before !== undefined) {
        await mkdir(path.dirname(file), { recursive: true });
        await writeFile(
          file,
          [...start.before, ...start.formulon, ...start.after].join(start.eol ?? "\n"),
        );
      }
      const kept = [...(start.before ?? []), ...start.after];
      const profile = PROFILE_OPTIONS[profileKey];
      const options = profile === undefined ? undefined : { profile };

      await writeCodexConfig(file, options);
      const first = await readFile(file, "utf8");
      // Exactly one formulon table, last, with the env line inside it and no sub-table.
      const lines = significantLines(first);
      assert.deepEqual(lines.slice(-formulonBlock(profile).length), formulonBlock(profile));
      assert.equal(first.endsWith("\n"), true);
      assert.equal(first.match(/^\[mcp_servers\.formulon\]/gm)?.length, 1);
      assert.doesNotMatch(first, /mcp_servers\.formulon\.env/);
      if (!profile) {
        assert.doesNotMatch(first, /FORMULON_MCP_PROFILE/);
      }
      // Every other table, key and comment is kept in order.
      assert.deepEqual(
        lines.slice(0, -formulonBlock(profile).length),
        significantLines(kept.join("\n")),
      );

      if (writes === "twice") {
        await writeCodexConfig(file, options);
        assert.equal(await readFile(file, "utf8"), first, "a repeated write changes nothing");
      } else if (writes === "profileChange") {
        const changed = profile === "mac-365-de_DE" ? "win-365-ja_JP" : "mac-365-de_DE";
        await writeCodexConfig(file, { profile: changed });
        const second = significantLines(await readFile(file, "utf8"));
        assert.deepEqual(second.slice(-formulonBlock(changed).length), formulonBlock(changed));
        assert.deepEqual(
          second.slice(0, -formulonBlock(changed).length),
          significantLines(kept.join("\n")),
        );
        await writeCodexConfig(file, options);
        assert.equal(
          await readFile(file, "utf8"),
          first,
          "switching back restores the first write",
        );
      }

      assert.equal(await removeFromCodexConfig(file), "removed");
      const removed = await readFile(file, "utf8");
      assert.doesNotMatch(removed, /mcp_servers\.formulon\]/);
      assert.doesNotMatch(removed, /FORMULON_MCP_PROFILE/);
      assert.deepEqual(significantLines(removed), significantLines(kept.join("\n")));
      assert.equal(await removeFromCodexConfig(file), "absent");
      assert.deepEqual(
        (await readdir(path.dirname(file))).filter((name) => name.endsWith(".tmp")),
        [],
      );
    } finally {
      await cleanup();
    }
  });
}

test("a hand-added env sub-table is replaced along with its parent", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, "config.toml");
    await writeFile(
      file,
      [
        "[mcp_servers.formulon]",
        'command = "npx"',
        "",
        "[mcp_servers.formulon.env]",
        'FORMULON_MCP_PROFILE = "mac-365-ja_JP"',
        'OTHER = "x"',
        "",
        "[keep]",
        "v = 1",
        "",
      ].join("\n"),
    );
    await writeCodexConfig(file, { profile: "win-365-de_DE" });
    const content = await readFile(file, "utf8");
    assert.equal(content.match(/FORMULON_MCP_PROFILE/g)?.length, 1);
    assert.doesNotMatch(content, /OTHER/);
    assert.doesNotMatch(content, /mac-365-ja_JP/);
    assert.match(content, /\[keep\]\nv = 1\n/);
    assert.match(content, /env = \{ FORMULON_MCP_PROFILE = "win-365-de_DE" \}\n$/);
  } finally {
    await cleanup();
  }
});

// A header with a trailing comment is a valid TOML table header, so rewriting must
// replace it rather than append a second `[mcp_servers.formulon]` table.
test("replaces a formulon table whose header carries a trailing comment", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, "config.toml");
    await writeFile(
      file,
      ["[mcp_servers.formulon] # added by hand", 'command = "npx"', 'args = ["-y", "x"]', ""].join(
        "\n",
      ),
    );
    await writeCodexConfig(file, { profile: "win-365-ja_JP" });
    const content = await readFile(file, "utf8");
    assert.equal(content.match(/^\[mcp_servers\.formulon\]/gm)?.length ?? 0, 1);
    assert.doesNotMatch(content, /added by hand/);
  } finally {
    await cleanup();
  }
});

// The formulon table ends at the next header of any shape. A commented table or an
// array of tables after it must survive the rewrite and the removal intact.
test("keeps the tables after the formulon table whatever their header shape", async () => {
  const { dir, cleanup } = await makeTmp();
  try {
    const file = path.join(dir, "config.toml");
    const after = [
      "[other] # keep me",
      "key = 1",
      "",
      "[[arr]] # x",
      "v = 2",
      "",
      "  [indented.table]",
      'w = "3"',
    ];
    await writeFile(
      file,
      ["[mcp_servers.formulon]", 'command = "old"', "", ...after, ""].join("\n"),
    );
    await writeCodexConfig(file, { profile: "win-365-ja_JP" });
    let content = await readFile(file, "utf8");
    for (const line of after.filter((line) => line !== "")) {
      assert.equal(content.includes(line), true, line);
    }
    assert.doesNotMatch(content, /command = "old"/);
    assert.equal(content.match(/^\[mcp_servers\.formulon\]/gm)?.length ?? 0, 1);

    assert.equal(await removeFromCodexConfig(file), "removed");
    content = await readFile(file, "utf8");
    for (const line of after.filter((line) => line !== "")) {
      assert.equal(content.includes(line), true, line);
    }
    assert.doesNotMatch(content, /mcp_servers\.formulon/);
  } finally {
    await cleanup();
  }
});

// ---------------------------------------------------------------------------
// createPrompter
// ---------------------------------------------------------------------------

const LINE_ENDINGS = { LF: "\n", CRLF: "\r\n", CR: "\r" };
const PROMPT_ANSWERS = {
  normal: ["1", "2", "y"],
  emptyAnswers: ["", "x", "", "y"],
  spaced: ["  two words  ", "\ttab", "end "],
};
const tick = () => new Promise((resolve) => setImmediate(resolve));

/** Pairwise rows of (line ending, trailing terminator, delivery, content); 12 rows. */
const PROMPTER_CASES = [
  ["LF", "yes", "atOnce", "emptyAnswers"],
  ["CRLF", "no", "byteByByte", "normal"],
  ["CR", "yes", "midLine", "normal"],
  ["CR", "no", "delayed", "spaced"],
  ["LF", "yes", "delayed", "normal"],
  ["CRLF", "yes", "midLine", "spaced"],
  ["LF", "no", "midLine", "emptyAnswers"],
  ["CR", "no", "atOnce", "spaced"],
  ["CRLF", "yes", "atOnce", "normal"],
  ["CRLF", "yes", "delayed", "emptyAnswers"],
  ["LF", "yes", "byteByByte", "spaced"],
  ["CR", "no", "byteByByte", "emptyAnswers"],
];

for (const [endingKey, trailing, delivery, contentKey] of PROMPTER_CASES) {
  test(`prompter reads ${contentKey} answers: ${endingKey}, trailing ${trailing}, ${delivery}`, async () => {
    const eol = LINE_ENDINGS[endingKey];
    const answers = PROMPT_ANSWERS[contentKey];
    const payload = answers.join(eol) + (trailing === "yes" ? eol : "");
    const input = new PassThrough();
    const output = new PassThrough();
    let written = "";
    output.on("data", (chunk) => {
      written += chunk.toString("utf8");
    });
    const prompter = createPrompter(input, output);
    const questions = answers.map((_, index) => `Question ${index + 1}: `);
    try {
      const asked = (async () => {
        const received = [];
        for (const question of questions) {
          received.push(await prompter.ask(question));
        }
        return received;
      })();
      if (delivery === "atOnce") {
        input.end(payload);
      } else if (delivery === "byteByByte") {
        for (const char of payload) {
          input.write(char);
          await tick();
        }
        input.end();
      } else if (delivery === "midLine") {
        // Split inside a line, and inside a CRLF pair when there is one.
        const cut = eol.length === 2 ? payload.indexOf("\r") + 1 : Math.ceil(payload.length / 2);
        input.write(payload.slice(0, cut));
        await tick();
        input.write(payload.slice(cut));
        input.end();
      } else {
        await new Promise((resolve) => setTimeout(resolve, 25));
        input.write(payload);
        await new Promise((resolve) => setTimeout(resolve, 25));
        input.end();
      }

      const received = await asked;
      assert.deepEqual(received, answers);
      for (const answer of received) {
        assert.doesNotMatch(answer, /[\r\n]/, "no line terminator leaks into an answer");
      }
      await assert.rejects(
        prompter.ask("One more: "),
        /^Error: Input ended before an answer to: One more:$/,
      );
      // Questions go to the output in the order asked; answers are not echoed.
      assert.equal(written, `${questions.join("")}One more: `);
    } finally {
      prompter.close();
    }
  });
}

test("a prompter reports the unanswered question without its trailing space", async () => {
  const input = new PassThrough();
  const output = new PassThrough();
  const prompter = createPrompter(input, output);
  input.end("only\n");
  try {
    assert.equal(await prompter.ask("First [1]: "), "only");
    await assert.rejects(
      prompter.ask("Second [2]: "),
      (error) => error.message === "Input ended before an answer to: Second [2]:",
    );
  } finally {
    prompter.close();
  }
});

test("a prompter on empty input fails the first question", async () => {
  const input = new PassThrough();
  const prompter = createPrompter(input, new PassThrough());
  input.end();
  try {
    await assert.rejects(prompter.ask("Anything: "), /Input ended before an answer to: Anything:/);
  } finally {
    prompter.close();
  }
});

test("a blank line is an empty answer, distinct from missing input", async () => {
  const input = new PassThrough();
  const prompter = createPrompter(input, new PassThrough());
  input.end("\n\n");
  try {
    assert.equal(await prompter.ask("A: "), "");
    assert.equal(await prompter.ask("B: "), "");
    await assert.rejects(prompter.ask("C: "), /Input ended/);
  } finally {
    prompter.close();
  }
});

// ---------------------------------------------------------------------------
// The installer end to end: `node dist/index.js init|uninstall` with piped answers
// ---------------------------------------------------------------------------

const CLI_PATH = fileURLToPath(new URL("../dist/index.js", import.meta.url));
const CLI_TIMEOUT_MS = 20_000;
const DEFAULT_HOST_KEY = platform() === "darwin" ? "2" : "1";

/**
 * Runs the CLI with HOME and the working directory inside `home`, so no real
 * config file is ever read or written. A hang is killed and fails the test.
 */
function runCli(home, args, input, env = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [CLI_PATH, ...args], {
      cwd: home,
      env: { HOME: home, USERPROFILE: home, LANG: "en_US.UTF-8", ...env },
      stdio: ["pipe", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      reject(new Error(`formulon-mcp ${args.join(" ")} did not exit within ${CLI_TIMEOUT_MS}ms`));
    }, CLI_TIMEOUT_MS);
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });
    child.on("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.on("close", (code, signal) => {
      clearTimeout(timer);
      resolve({ code, signal, stdout, stderr });
    });
    child.stdin.on("error", () => {
      // The child may exit before reading all of a long input.
    });
    child.stdin.end(input);
  });
}

async function withHome(fn) {
  const home = await mkdtemp(path.join(tmpdir(), "formulon-mcp-cli-"));
  try {
    return await fn(home);
  } finally {
    await rm(home, { recursive: true, force: true });
  }
}

const readJson = async (file) => JSON.parse(await readFile(file, "utf8"));
const CLAUDE_FILE = (home) => path.join(home, ".claude.json");
const CODEX_FILE = (home) => path.join(home, ".codex", "config.toml");

test("init with every answer piped at once writes the chosen profile and exits 0", async () => {
  await withHome(async (home) => {
    // Targets, locale (Japanese), host (Windows), proceed.
    const result = await runCli(home, ["init"], "1\n2\n1\ny\n");
    assert.equal(result.code, 0, result.stderr);
    assert.equal(result.stderr, "");
    assert.deepEqual((await readJson(CLAUDE_FILE(home))).mcpServers.formulon, {
      command: "npx",
      args: ["-y", "@libraz/formulon-mcp"],
      env: { FORMULON_MCP_PROFILE: "win-365-ja_JP" },
    });
    assert.match(result.stdout, /Excel profile: win-365-ja_JP/);
    assert.match(result.stdout, /Wrote ~\/\.claude\.json/);
    assert.match(result.stdout, /Done\. Restart your MCP client/);
  });
});

test("init with the engine-default locale writes no env and skips the host question", async () => {
  await withHome(async (home) => {
    // Locale 0 then proceed: a host answer would be read as the proceed answer.
    const result = await runCli(home, ["init"], "1\n0\ny\n");
    assert.equal(result.code, 0, result.stderr);
    assert.deepEqual((await readJson(CLAUDE_FILE(home))).mcpServers.formulon, {
      command: "npx",
      args: ["-y", "@libraz/formulon-mcp"],
    });
    assert.match(result.stdout, /Excel profile: engine default \(win-365-en_US\)/);
    assert.doesNotMatch(result.stdout, /Excel host:/);
  });
});

const PROFILE_FLAG_FORMS = [
  [["init", "--profile", "win-365-ja_JP"], "win-365-ja_JP"],
  [["init", "--profile=mac-365-ja_JP"], "mac-365-ja_JP"],
  [["init", "--profile=win-365-de_DE"], "win-365-de_DE"],
  [["init", "--profile", "mac-365-th_TH"], "mac-365-th_TH"],
];

for (const [args, profile] of PROFILE_FLAG_FORMS) {
  test(`init ${args.slice(1).join(" ")} skips the locale and host menus`, async () => {
    await withHome(async (home) => {
      // Only the target and the confirmation are asked.
      const result = await runCli(home, args, "1,3\ny\n");
      assert.equal(result.code, 0, result.stderr);
      assert.doesNotMatch(result.stdout, /Excel locale new sessions evaluate with/);
      assert.doesNotMatch(result.stdout, /Excel host:/);
      assert.match(result.stdout, new RegExp(`Excel profile: ${profile}`));
      assert.equal(
        (await readJson(CLAUDE_FILE(home))).mcpServers.formulon.env.FORMULON_MCP_PROFILE,
        profile,
      );
      assert.match(
        await readFile(CODEX_FILE(home), "utf8"),
        new RegExp(`env = \\{ FORMULON_MCP_PROFILE = "${profile}" \\}\\n$`),
      );
    });
  });
}

const BAD_PROFILE_FLAGS = [
  [["init", "--profile"], /^\[formulon-mcp\] --profile needs a profile id/],
  [["init", "--profile="], /^\[formulon-mcp\] Unknown Excel profile: \s*\(expected/],
  [["init", "--profile", "win-365-xx_XX"], /Unknown Excel profile: win-365-xx_XX/],
  [["init", "--profile=win-365-xx_XX"], /Unknown Excel profile: win-365-xx_XX/],
  [["init", "--profile", "ja_JP"], /Unknown Excel profile: ja_JP/],
  [["init", "--profile=WIN-365-ja_JP"], /Unknown Excel profile: WIN-365-ja_JP/],
];

for (const [args, pattern] of BAD_PROFILE_FLAGS) {
  test(`init ${args.slice(1).join(" ")} fails before asking anything`, async () => {
    await withHome(async (home) => {
      const result = await runCli(home, args, "1\ny\n");
      assert.equal(result.code, 1);
      assert.equal(result.signal, null);
      assert.match(result.stderr, pattern);
      assert.equal(result.stdout, "", "no menu is shown for a rejected profile");
      assert.deepEqual(await readdir(home), []);
    });
  });
}

const TRUNCATED_INPUTS = [
  ["", /Input ended before an answer to: Choice \[1\]:/],
  ["1\n", /Input ended before an answer to: Choice \[1\]:/],
  ["1\n2\n", new RegExp(`Input ended before an answer to: Choice \\[${DEFAULT_HOST_KEY}\\]:`)],
  ["1\n2\n1\n", /Input ended before an answer to: Proceed\? \[Y\/n\]/],
  ["1\n0\n", /Input ended before an answer to: Proceed\? \[Y\/n\]/],
];

for (const [input, pattern] of TRUNCATED_INPUTS) {
  test(`init exits 1 and names the unanswered question for input ${JSON.stringify(input)}`, async () => {
    await withHome(async (home) => {
      const result = await runCli(home, ["init"], input);
      assert.equal(result.code, 1);
      assert.match(result.stderr, /^\[formulon-mcp\] Input ended before an answer to:/);
      assert.match(result.stderr, pattern);
      assert.deepEqual(await readdir(home), [], "nothing is written");
    });
  });
}

const INVALID_ANSWERS = [
  ["9\n", /Invalid choice: 9/],
  [" , \n", /No targets selected/],
  ["1,x\n", /Invalid choice: x/],
  ["1\n99\n", /Invalid choice: 99/],
  ["1\nja\n", /Invalid choice: ja/],
  ["1\n2\n3\n", /Invalid choice: 3/],
];

for (const [input, pattern] of INVALID_ANSWERS) {
  test(`init rejects the answers ${JSON.stringify(input)}`, async () => {
    await withHome(async (home) => {
      const result = await runCli(home, ["init"], `${input}y\ny\n`);
      assert.equal(result.code, 1);
      assert.match(result.stderr, pattern);
      assert.deepEqual(await readdir(home), []);
    });
  });
}

const PROCEED_ANSWERS = [
  ["y", true],
  ["Y", true],
  ["yes", true],
  ["YES", true],
  [" y ", true],
  ["", true],
  ["n", false],
  ["N", false],
  ["no", false],
  ["maybe", false],
  ["yy", false],
];

for (const [answer, proceeds] of PROCEED_ANSWERS) {
  test(`init treats ${JSON.stringify(answer)} at Proceed as ${proceeds ? "yes" : "no"}`, async () => {
    await withHome(async (home) => {
      const result = await runCli(home, ["init", "--profile=win-365-ja_JP"], `1\n${answer}\n`);
      assert.equal(result.code, 0, result.stderr);
      if (proceeds) {
        assert.equal(
          (await readJson(CLAUDE_FILE(home))).mcpServers.formulon.env.FORMULON_MCP_PROFILE,
          "win-365-ja_JP",
        );
      } else {
        assert.match(result.stdout, /Aborted\./);
        assert.doesNotMatch(result.stdout, /Wrote /);
        assert.deepEqual(await readdir(home), []);
      }
    });
  });
}

/** The locale menu default each LANG should produce, as the init menu shows it. */
const LANG_DEFAULTS = {
  "ja_JP.UTF-8": "2",
  "en_US.UTF-8": "1",
  "de_DE.UTF-8": "3",
  C: "1",
  "": "0",
};
const LOCALE_ANSWER = { blank: "", 2: "2", 0: "0", 14: "14" };
const HOST_ANSWER = { blank: "", 1: "1", 2: "2" };
const PROCEED_ANSWER = { y: "y", blank: "", n: "n" };
const TARGET_ANSWER = { 1: "1", blank: "", "1,3": "1,3", 3: "3" };
const LANG_KEYS = {
  ja_JP: "ja_JP.UTF-8",
  en_US: "en_US.UTF-8",
  de_DE: "de_DE.UTF-8",
  C: "C",
  empty: "",
};

/** Pairwise rows of (target, locale, host, LANG, proceed); 25 rows covering every value pair. */
const INIT_FLOW_CASES = [
  ["blank", "0", "2", "empty", "blank"],
  ["1", "14", "blank", "de_DE", "n"],
  ["3", "blank", "1", "en_US", "blank"],
  ["1,3", "14", "1", "ja_JP", "blank"],
  ["1,3", "2", "blank", "C", "blank"],
  ["3", "2", "2", "empty", "n"],
  ["1", "blank", "1", "empty", "y"],
  ["blank", "14", "2", "en_US", "y"],
  ["blank", "blank", "2", "de_DE", "n"],
  ["1", "0", "blank", "en_US", "blank"],
  ["1,3", "2", "1", "de_DE", "y"],
  ["1,3", "0", "2", "de_DE", "blank"],
  ["3", "0", "2", "ja_JP", "y"],
  ["blank", "2", "1", "de_DE", "blank"],
  ["3", "14", "blank", "empty", "blank"],
  ["1,3", "blank", "1", "en_US", "n"],
  ["3", "0", "1", "C", "n"],
  ["1", "blank", "2", "C", "y"],
  ["blank", "14", "blank", "C", "y"],
  ["3", "blank", "2", "de_DE", "y"],
  ["blank", "blank", "blank", "ja_JP", "n"],
  ["1,3", "blank", "2", "empty", "n"],
  ["3", "2", "1", "en_US", "blank"],
  ["1", "0", "2", "ja_JP", "n"],
  ["1", "2", "1", "ja_JP", "n"],
];

for (const [targetKey, localeKey, hostKey, langKey, proceedKey] of INIT_FLOW_CASES) {
  test(`init flow: target ${targetKey}, locale ${localeKey}, host ${hostKey}, LANG ${langKey}, proceed ${proceedKey}`, async () => {
    await withHome(async (home) => {
      const lang = LANG_KEYS[langKey];
      const defaultLocale = LANG_DEFAULTS[lang];
      const effectiveLocale = LOCALE_ANSWER[localeKey] || defaultLocale;
      const effectiveHost = HOST_ANSWER[hostKey] || DEFAULT_HOST_KEY;
      const profile =
        effectiveLocale === "0"
          ? undefined
          : `${effectiveHost === "1" ? "win" : "mac"}-365-${LOCALE_CODES[Number(effectiveLocale) - 1]}`;
      const proceeds = proceedKey !== "n";
      const targets = TARGET_ANSWER[targetKey] || "1";

      const answers = [
        TARGET_ANSWER[targetKey],
        LOCALE_ANSWER[localeKey],
        ...(effectiveLocale === "0" ? [] : [HOST_ANSWER[hostKey]]),
        PROCEED_ANSWER[proceedKey],
      ];
      const result = await runCli(home, ["init"], `${answers.join("\n")}\n`, { LANG: lang });
      assert.equal(result.code, 0, result.stderr);
      // The prompt advertises the default that a bare Enter takes.
      assert.ok(result.stdout.includes(`Choice [${defaultLocale}]: `), result.stdout);
      if (effectiveLocale !== "0") {
        assert.ok(result.stdout.includes(`Choice [${DEFAULT_HOST_KEY}]: `), result.stdout);
      }
      assert.ok(
        result.stdout.includes(`Excel profile: ${profile ?? "engine default (win-365-en_US)"}`),
      );

      if (!proceeds) {
        assert.match(result.stdout, /Aborted\./);
        assert.deepEqual(await readdir(home), []);
        return;
      }
      if (targets.split(",").includes("1")) {
        const entry = (await readJson(CLAUDE_FILE(home))).mcpServers.formulon;
        assert.deepEqual(entry.env, profile ? { FORMULON_MCP_PROFILE: profile } : undefined);
      } else {
        await assert.rejects(readFile(CLAUDE_FILE(home)), { code: "ENOENT" });
      }
      if (targets.split(",").includes("3")) {
        const toml = await readFile(CODEX_FILE(home), "utf8");
        assert.equal(toml.match(/^\[mcp_servers\.formulon\]$/gm)?.length, 1);
        if (profile) {
          assert.ok(toml.includes(`env = { FORMULON_MCP_PROFILE = "${profile}" }`));
        } else {
          assert.doesNotMatch(toml, /FORMULON_MCP_PROFILE/);
        }
      } else {
        await assert.rejects(readFile(CODEX_FILE(home)), { code: "ENOENT" });
      }
    });
  });
}

test("the system locale decides what a bare Enter picks at the locale prompt", async () => {
  const outcomes = {};
  for (const lang of ["ja_JP.UTF-8", "en_US.UTF-8", "de_DE.UTF-8", "fr_FR.UTF-8", "ko_KR.UTF-8"]) {
    await withHome(async (home) => {
      const result = await runCli(home, ["init"], "1\n\n1\ny\n", { LANG: lang });
      assert.equal(result.code, 0, result.stderr);
      outcomes[lang] = (
        await readJson(CLAUDE_FILE(home))
      ).mcpServers.formulon.env.FORMULON_MCP_PROFILE;
    });
  }
  assert.deepEqual(outcomes, {
    "ja_JP.UTF-8": "win-365-ja_JP",
    "en_US.UTF-8": "win-365-en_US",
    "de_DE.UTF-8": "win-365-de_DE",
    "fr_FR.UTF-8": "win-365-fr_FR",
    "ko_KR.UTF-8": "win-365-ko_KR",
  });
});

test("init preserves an existing config and replaces only the formulon entry", async () => {
  await withHome(async (home) => {
    await writeFile(
      CLAUDE_FILE(home),
      JSON.stringify({ theme: "dark", mcpServers: { other: { command: "node", args: ["x"] } } }),
    );
    const result = await runCli(home, ["init", "--profile=win-365-de_DE"], "\ny\n");
    assert.equal(result.code, 0, result.stderr);
    assert.match(result.stdout, /\(merge\)/);
    const config = await readJson(CLAUDE_FILE(home));
    assert.equal(config.theme, "dark");
    assert.deepEqual(config.mcpServers.other, { command: "node", args: ["x"] });
    assert.equal(config.mcpServers.formulon.env.FORMULON_MCP_PROFILE, "win-365-de_DE");

    const again = await runCli(home, ["init"], "1\n0\ny\n");
    assert.equal(again.code, 0, again.stderr);
    assert.match(again.stdout, /\(replace formulon\)/);
    assert.equal((await readJson(CLAUDE_FILE(home))).mcpServers.formulon.env, undefined);
  });
});

test("uninstall removes the formulon entry and its env from the chosen configs", async () => {
  await withHome(async (home) => {
    await writeFile(
      CLAUDE_FILE(home),
      JSON.stringify({ mcpServers: { other: { command: "node" } } }),
    );
    const install = await runCli(home, ["init", "--profile=win-365-ja_JP"], "1,3\ny\n");
    assert.equal(install.code, 0, install.stderr);
    await writeFile(
      CODEX_FILE(home),
      `[keep]\nv = 1\n\n${(await readFile(CODEX_FILE(home), "utf8")).trimEnd()}\n`,
    );

    const result = await runCli(home, ["uninstall"], "1,3\ny\n");
    assert.equal(result.code, 0, result.stderr);
    assert.match(result.stdout, /Removed formulon from ~\/\.claude\.json/);
    assert.match(result.stdout, /Removed formulon from ~\/\.codex\/config\.toml/);
    const config = await readJson(CLAUDE_FILE(home));
    assert.deepEqual(config.mcpServers, { other: { command: "node" } });
    const toml = await readFile(CODEX_FILE(home), "utf8");
    assert.doesNotMatch(toml, /formulon|FORMULON_MCP_PROFILE/);
    assert.match(toml, /\[keep\]\nv = 1\n/);

    // A second run finds nothing to remove.
    const again = await runCli(home, ["uninstall"], "1,3\ny\n");
    assert.equal(again.code, 0, again.stderr);
    assert.match(again.stdout, /No formulon in ~\/\.claude\.json; skipped\./);
  });
});

test("uninstall with the default selection skips files that do not exist", async () => {
  await withHome(async (home) => {
    const install = await runCli(home, ["init", "--profile=win-365-ja_JP"], "1\ny\n");
    assert.equal(install.code, 0, install.stderr);
    const result = await runCli(home, ["uninstall"], "\ny\n");
    assert.equal(result.code, 0, result.stderr);
    assert.match(result.stdout, /Removed formulon from ~\/\.claude\.json/);
    assert.equal(result.stdout.match(/does not exist; skipped\./g)?.length, 3);
    assert.deepEqual((await readJson(CLAUDE_FILE(home))).mcpServers, {});
  });
});

test("uninstall declined or truncated changes nothing", async () => {
  await withHome(async (home) => {
    await runCli(home, ["init", "--profile=win-365-ja_JP"], "1\ny\n");
    const before = await readFile(CLAUDE_FILE(home), "utf8");

    const declined = await runCli(home, ["uninstall"], "1\nn\n");
    assert.equal(declined.code, 0, declined.stderr);
    assert.match(declined.stdout, /Aborted\./);
    assert.equal(await readFile(CLAUDE_FILE(home), "utf8"), before);

    const truncated = await runCli(home, ["uninstall"], "1\n");
    assert.equal(truncated.code, 1);
    assert.match(truncated.stderr, /Input ended before an answer to: Proceed\? \[Y\/n\]/);
    assert.equal(await readFile(CLAUDE_FILE(home), "utf8"), before);

    const invalid = await runCli(home, ["uninstall"], "7\ny\n");
    assert.equal(invalid.code, 1);
    assert.match(invalid.stderr, /Invalid choice: 7/);
    assert.equal(await readFile(CLAUDE_FILE(home), "utf8"), before);
  });
});

test("help and version do not touch any config", async () => {
  await withHome(async (home) => {
    const help = await runCli(home, ["--help"], "");
    assert.equal(help.code, 0);
    assert.match(help.stdout, /formulon-mcp init \[--profile <id>\]/);
    const version = await runCli(home, ["--version"], "");
    assert.equal(version.code, 0);
    assert.match(version.stdout, /^\d+\.\d+\.\d+\n$/);
    assert.deepEqual(await readdir(home), []);
  });
});
