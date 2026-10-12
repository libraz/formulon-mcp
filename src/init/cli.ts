import { homedir, platform } from "node:os";
import { join } from "node:path";
import { stdin, stdout } from "node:process";
import { isKnownProfile } from "../formulon.js";
import {
  claudeDesktopConfigPath,
  displayPath,
  type InstallOptions,
  previewRemoveImpact,
  previewWriteImpact,
  type RemoveOutcome,
  removeFromCodexConfig,
  removeFromJsonConfig,
  SERVER_NAME,
  writeCodexConfig,
  writeJsonConfig,
} from "./config.js";
import { createPrompter, type Prompter } from "./prompt.js";

const promptYesNo = async (
  prompter: Prompter,
  question: string,
  defaultYes: boolean,
): Promise<boolean> => {
  const def = defaultYes ? "Y/n" : "y/N";
  const ans = (await prompter.ask(`${question} [${def}] `)).trim().toLowerCase();
  if (ans === "") {
    return defaultYes;
  }
  return ans === "y" || ans === "yes";
};

export type TargetKind = "claude-user" | "claude-project" | "codex" | "claude-desktop";

type TargetOption = {
  key: string;
  kind: TargetKind;
  label: string;
  resolvePath: () => string;
};

const TARGET_OPTIONS: TargetOption[] = [
  {
    key: "1",
    kind: "claude-user",
    label: "Claude Code — user",
    resolvePath: () => join(homedir(), ".claude.json"),
  },
  {
    key: "2",
    kind: "claude-project",
    label: "Claude Code — project",
    resolvePath: () => join(process.cwd(), ".mcp.json"),
  },
  {
    key: "3",
    kind: "codex",
    label: "Codex CLI",
    resolvePath: () => join(homedir(), ".codex", "config.toml"),
  },
  {
    key: "4",
    kind: "claude-desktop",
    label: "Claude Desktop",
    resolvePath: claudeDesktopConfigPath,
  },
];

export const parseTargetChoice = (raw: string, defaultRaw: string): TargetKind[] => {
  const input = raw.trim() === "" ? defaultRaw : raw;
  const parts = input
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  if (parts.length === 0) {
    throw new Error("No targets selected.");
  }
  const selected = new Set<TargetKind>();
  for (const p of parts) {
    const opt = TARGET_OPTIONS.find((o) => o.key === p);
    if (!opt) {
      throw new Error(`Invalid choice: ${p}`);
    }
    selected.add(opt.kind);
  }
  return [...selected];
};

const writeForKind = (
  kind: TargetKind,
  path: string,
  options: InstallOptions,
): (() => Promise<void>) => {
  if (kind === "codex") {
    return () => writeCodexConfig(path, options);
  }
  return () => writeJsonConfig(path, options);
};

const removeForKind = (kind: TargetKind, path: string): (() => Promise<RemoveOutcome>) => {
  if (kind === "codex") {
    return () => removeFromCodexConfig(path);
  }
  return () => removeFromJsonConfig(path);
};

type Target = { label: string; path: string; write: () => Promise<void> };

const pickTargets = (kinds: TargetKind[], options: InstallOptions): Target[] => {
  const targets: Target[] = [];
  for (const opt of TARGET_OPTIONS) {
    if (!kinds.includes(opt.kind)) {
      continue;
    }
    const path = opt.resolvePath();
    targets.push({ label: opt.label, path, write: writeForKind(opt.kind, path, options) });
  }
  return targets;
};

const renderTargetMenu = (): string =>
  [
    `  1) Claude Code              ${displayPath(join(homedir(), ".claude.json"))}`,
    `  2) Claude Code (project)    ${displayPath(join(process.cwd(), ".mcp.json"))}`,
    `  3) Codex CLI                ${displayPath(join(homedir(), ".codex", "config.toml"))}`,
    `  4) Claude Desktop           ${displayPath(claudeDesktopConfigPath())}`,
  ].join("\n");

/** Locales the engine has Excel profiles for, in menu order. */
export const PROFILE_LOCALES = [
  { code: "en_US", label: "English (United States)" },
  { code: "ja_JP", label: "Japanese" },
  { code: "de_DE", label: "German" },
  { code: "fr_FR", label: "French" },
  { code: "es_ES", label: "Spanish (Spain)" },
  { code: "es_MX", label: "Spanish (Mexico)" },
  { code: "pt_BR", label: "Portuguese (Brazil)" },
  { code: "it_IT", label: "Italian" },
  { code: "nl_NL", label: "Dutch" },
  { code: "ru_RU", label: "Russian" },
  { code: "zh_CN", label: "Chinese (Simplified)" },
  { code: "zh_TW", label: "Chinese (Traditional)" },
  { code: "ko_KR", label: "Korean" },
  { code: "th_TH", label: "Thai" },
] as const;

const PROFILE_HOSTS = [
  { key: "1", host: "win", label: "Windows" },
  { key: "2", host: "mac", label: "Mac" },
] as const;

/** Rejects a profile id the engine does not know. */
export const assertKnownProfile = async (profile: string): Promise<void> => {
  if (!(await isKnownProfile(profile))) {
    throw new Error(
      `Unknown Excel profile: ${profile} (expected {mac|win}-365-<locale>, for example win-365-ja_JP)`,
    );
  }
};

/**
 * Menu key of the locale matching a BCP 47 tag such as `ja-JP`, or "0" for
 * none. Without an exact region match, a locale of the same language and
 * script wins (`zh-HK` is Traditional, so zh_TW), then the same language.
 */
export const defaultLocaleKey = (tag: string): string => {
  let wanted: Intl.Locale;
  try {
    // POSIX spellings (`ja_JP.UTF-8@euro`) drop the codeset and modifier first.
    wanted = new Intl.Locale(tag.split(/[.@]/)[0].replaceAll("_", "-"));
  } catch {
    return "0";
  }
  // An undetermined language guesses nothing, though likely subtags would say
  // English; V8 reports `und` as an absent language.
  if (!wanted.language || wanted.language === "und") {
    return "0";
  }
  wanted = wanted.maximize();
  const offered = PROFILE_LOCALES.map((entry) =>
    new Intl.Locale(entry.code.replace("_", "-")).maximize(),
  );
  const pick = (match: (locale: Intl.Locale) => boolean) => offered.findIndex(match);
  const sameLanguage = (locale: Intl.Locale) => locale.language === wanted.language;
  const candidates = [
    pick((locale) => sameLanguage(locale) && locale.region === wanted.region),
    pick((locale) => sameLanguage(locale) && locale.script === wanted.script),
    pick(sameLanguage),
  ];
  const index = candidates.find((candidate) => candidate >= 0);
  return index === undefined ? "0" : String(index + 1);
};

/**
 * Maps the locale and host menu answers to a profile id; locale "0" keeps the
 * engine default and returns undefined.
 */
export const parseProfileChoice = (
  localeRaw: string,
  hostRaw: string,
  defaults: { locale: string; host: string },
): string | undefined => {
  const localeKey = localeRaw.trim() === "" ? defaults.locale : localeRaw.trim();
  if (localeKey === "0") {
    return undefined;
  }
  const locale = PROFILE_LOCALES[Number(localeKey) - 1];
  if (!(locale && /^[0-9]+$/.test(localeKey))) {
    throw new Error(`Invalid choice: ${localeKey}`);
  }
  const hostKey = hostRaw.trim() === "" ? defaults.host : hostRaw.trim();
  const host = PROFILE_HOSTS.find((option) => option.key === hostKey);
  if (!host) {
    throw new Error(`Invalid choice: ${hostKey}`);
  }
  return `${host.host}-365-${locale.code}`;
};

const promptProfile = async (prompter: Prompter): Promise<string | undefined> => {
  const defaults = {
    locale: defaultLocaleKey(Intl.DateTimeFormat().resolvedOptions().locale),
    host: platform() === "darwin" ? "2" : "1",
  };
  stdout.write("\nExcel locale new sessions evaluate with:\n");
  stdout.write("   0) Engine default (English, Windows)\n");
  for (const [index, locale] of PROFILE_LOCALES.entries()) {
    stdout.write(`  ${String(index + 1).padStart(2)}) ${locale.label.padEnd(24)} ${locale.code}\n`);
  }
  const localeRaw = (await prompter.ask(`Choice [${defaults.locale}]: `)).trim();
  if ((localeRaw === "" ? defaults.locale : localeRaw) === "0") {
    return parseProfileChoice(localeRaw, "", defaults);
  }
  stdout.write("\nExcel host:\n");
  for (const option of PROFILE_HOSTS) {
    stdout.write(`  ${option.key}) ${option.label}\n`);
  }
  const hostRaw = (await prompter.ask(`Choice [${defaults.host}]: `)).trim();
  return parseProfileChoice(localeRaw, hostRaw, defaults);
};

/**
 * Interactive setup. A `profile` given on the command line is validated and
 * used without asking; otherwise the profile is prompted for.
 */
export const runInit = async (options: InstallOptions = {}): Promise<void> => {
  if (options.profile !== undefined) {
    await assertKnownProfile(options.profile);
  }
  const prompter = createPrompter(stdin, stdout);
  try {
    stdout.write("formulon-mcp setup\n\n");
    stdout.write("Where to install? (pick one or more, comma-separated)\n");
    stdout.write(`${renderTargetMenu()}\n`);
    const choice = (await prompter.ask("Choice [1]: ")).trim();
    const kinds = parseTargetChoice(choice, "1");
    const profile = options.profile ?? (await promptProfile(prompter));
    const targets = pickTargets(kinds, { profile });

    stdout.write("\nWill update:\n");
    for (const t of targets) {
      const summary = await previewWriteImpact(t.path);
      stdout.write(`  - ${displayPath(t.path)} ${summary}\n`);
    }
    stdout.write(`Excel profile: ${profile ?? "engine default (win-365-en_US)"}\n\n`);

    const confirmed = await promptYesNo(prompter, "Proceed?", true);
    if (!confirmed) {
      stdout.write("Aborted.\n");
      return;
    }

    for (const t of targets) {
      await t.write();
      stdout.write(`Wrote ${displayPath(t.path)}\n`);
    }

    stdout.write("\nDone. Restart your MCP client to pick up the new server.\n");
  } finally {
    prompter.close();
  }
};

type RemoveTarget = {
  label: string;
  path: string;
  remove: () => Promise<RemoveOutcome>;
};

const pickRemoveTargets = (kinds: TargetKind[]): RemoveTarget[] => {
  const targets: RemoveTarget[] = [];
  for (const opt of TARGET_OPTIONS) {
    if (!kinds.includes(opt.kind)) {
      continue;
    }
    const path = opt.resolvePath();
    targets.push({ label: opt.label, path, remove: removeForKind(opt.kind, path) });
  }
  return targets;
};

export const runUninstall = async (): Promise<void> => {
  const prompter = createPrompter(stdin, stdout);
  try {
    stdout.write("formulon-mcp uninstall\n\n");
    stdout.write("Where to remove from? (pick one or more, comma-separated)\n");
    stdout.write(`${renderTargetMenu()}\n`);
    const choice = (await prompter.ask("Choice [1,2,3,4]: ")).trim();
    const kinds = parseTargetChoice(choice, "1,2,3,4");
    const targets = pickRemoveTargets(kinds);

    stdout.write("\nWill update:\n");
    for (const t of targets) {
      const summary = await previewRemoveImpact(t.path);
      stdout.write(`  - ${displayPath(t.path)} ${summary}\n`);
    }
    stdout.write("\n");

    const confirmed = await promptYesNo(prompter, "Proceed?", true);
    if (!confirmed) {
      stdout.write("Aborted.\n");
      return;
    }

    for (const t of targets) {
      const outcome = await t.remove();
      const p = displayPath(t.path);
      if (outcome === "removed") {
        stdout.write(`Removed ${SERVER_NAME} from ${p}\n`);
      } else if (outcome === "absent") {
        stdout.write(`No ${SERVER_NAME} in ${p}; skipped.\n`);
      } else {
        stdout.write(`${p} does not exist; skipped.\n`);
      }
    }

    stdout.write("\nDone. Restart your MCP client for the change to take effect.\n");
  } finally {
    prompter.close();
  }
};
