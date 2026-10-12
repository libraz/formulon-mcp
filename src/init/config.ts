import { existsSync } from "node:fs";
import { mkdir, readFile, rename, unlink, writeFile } from "node:fs/promises";
import { homedir, platform } from "node:os";
import { dirname, join } from "node:path";
import { PROFILE_ENV } from "../formulon.js";

const PACKAGE_SPEC = "@libraz/formulon-mcp";
export const SERVER_NAME = "formulon";

type McpServerEntry = {
  command: string;
  args: string[];
  env?: Record<string, string>;
};

/** What `init` writes beyond the launch command. */
export type InstallOptions = {
  /** Excel profile new sessions default to, written as `FORMULON_MCP_PROFILE`. */
  profile?: string;
};

type ClaudeConfig = {
  mcpServers?: Record<string, McpServerEntry>;
  [key: string]: unknown;
};

const buildEntry = (options: InstallOptions): McpServerEntry => ({
  command: "npx",
  args: ["-y", PACKAGE_SPEC],
  ...(options.profile ? { env: { [PROFILE_ENV]: options.profile } } : {}),
});

/**
 * Writes `data` to `path` through a sibling temp file and an atomic rename.
 *
 * These are other tools' config files, and `~/.claude.json` in particular is
 * large and read by a client that may be running right now, so a partial
 * in-place write would leave it truncated and unparseable.
 */
const writeFileAtomic = async (path: string, data: string): Promise<void> => {
  const tmp = `${path}.formulon-mcp.tmp`;
  try {
    await writeFile(tmp, data);
    await rename(tmp, path);
  } catch (error) {
    await unlink(tmp).catch(() => {
      // Best effort: there is nothing to clean up if the temp file never existed.
    });
    throw error;
  }
};

export const writeJsonConfig = async (
  path: string,
  options: InstallOptions = {},
): Promise<void> => {
  let data: ClaudeConfig = {};
  if (existsSync(path)) {
    const raw = await readFile(path, "utf8");
    if (raw.trim().length > 0) {
      try {
        data = JSON.parse(raw) as ClaudeConfig;
      } catch {
        throw new Error(
          `Failed to parse existing JSON at ${path}. Fix or move it before re-running init.`,
        );
      }
    }
  }
  data.mcpServers = data.mcpServers ?? {};
  data.mcpServers[SERVER_NAME] = buildEntry(options);
  await mkdir(dirname(path), { recursive: true });
  await writeFileAtomic(path, `${JSON.stringify(data, null, 2)}\n`);
};

export const writeClaudeConfig = (path: string, options?: InstallOptions): Promise<void> =>
  writeJsonConfig(path, options);
export const writeClaudeDesktopConfig = (path: string, options?: InstallOptions): Promise<void> =>
  writeJsonConfig(path, options);

/**
 * Matches the server's own Codex TOML header and any sub-table under it, so a
 * hand-added `[mcp_servers.formulon.env]` is rewritten with its parent rather
 * than left behind as an orphan section.
 */
const CODEX_SECTION_HEADER = new RegExp(
  `^[ \\t]*\\[mcp_servers\\.${SERVER_NAME}(?:\\.[^\\]]+)?\\][ \\t]*(?:#.*)?$`,
);

/** The same matcher applied line-by-line across a whole document. */
const CODEX_SECTION_HEADER_ANYWHERE = new RegExp(CODEX_SECTION_HEADER.source, "m");

/** Any TOML table or array-of-tables header, with an optional trailing comment. */
const TOML_TABLE_HEADER = /^[ \t]*\[\[?[^\]]+\]\]?[ \t]*(?:#.*)?$/;

const stripCodexFormulonSection = (content: string): string => {
  const lines = content.split("\n");
  const out: string[] = [];
  let inBlock = false;
  for (const line of lines) {
    const trimmed = line.trim();
    if (CODEX_SECTION_HEADER.test(trimmed)) {
      inBlock = true;
      continue;
    }
    if (inBlock) {
      if (TOML_TABLE_HEADER.test(trimmed)) {
        inBlock = false;
        out.push(line);
      }
      continue;
    }
    out.push(line);
  }
  return out.join("\n");
};

export const writeCodexConfig = async (
  path: string,
  options: InstallOptions = {},
): Promise<void> => {
  let existing = "";
  if (existsSync(path)) {
    existing = await readFile(path, "utf8");
  }
  const stripped = stripCodexFormulonSection(existing).replace(/\n*$/, "");
  const prefix = stripped.length > 0 ? `${stripped}\n\n` : "";
  const block = [
    `[mcp_servers.${SERVER_NAME}]`,
    'command = "npx"',
    `args = ["-y", "${PACKAGE_SPEC}"]`,
    ...(options.profile ? [`env = { ${PROFILE_ENV} = "${options.profile}" }`] : []),
  ].join("\n");
  await mkdir(dirname(path), { recursive: true });
  await writeFileAtomic(path, `${prefix}${block}\n`);
};

export const displayPath = (path: string): string => {
  const home = homedir();
  if (path === home) {
    return "~";
  }
  if (path.startsWith(`${home}/`)) {
    return `~/${path.slice(home.length + 1)}`;
  }
  return path;
};

const hasFormulonEntry = (
  data: { mcpServers?: Record<string, unknown> } | null | undefined,
): boolean => Boolean(data?.mcpServers && SERVER_NAME in data.mcpServers);

export const previewWriteImpact = async (path: string): Promise<string> => {
  if (!existsSync(path)) {
    return "(new)";
  }
  const raw = await readFile(path, "utf8").catch(() => "");
  if (raw.trim().length === 0) {
    return "(new)";
  }
  try {
    const data = JSON.parse(raw) as { mcpServers?: Record<string, unknown> };
    return hasFormulonEntry(data) ? `(replace ${SERVER_NAME})` : "(merge)";
  } catch {
    return CODEX_SECTION_HEADER_ANYWHERE.test(raw) ? `(replace ${SERVER_NAME})` : "(merge)";
  }
};

export const previewRemoveImpact = async (path: string): Promise<string> => {
  if (!existsSync(path)) {
    return "(no file; skip)";
  }
  const raw = await readFile(path, "utf8").catch(() => "");
  if (raw.trim().length === 0) {
    return `(no ${SERVER_NAME}; skip)`;
  }
  try {
    const data = JSON.parse(raw) as { mcpServers?: Record<string, unknown> };
    return hasFormulonEntry(data) ? `(remove ${SERVER_NAME})` : `(no ${SERVER_NAME}; skip)`;
  } catch {
    return CODEX_SECTION_HEADER_ANYWHERE.test(raw)
      ? `(remove ${SERVER_NAME})`
      : `(no ${SERVER_NAME}; skip)`;
  }
};

export type RemoveOutcome = "removed" | "absent" | "no-file";

export const removeFromJsonConfig = async (path: string): Promise<RemoveOutcome> => {
  if (!existsSync(path)) {
    return "no-file";
  }
  const raw = await readFile(path, "utf8");
  if (raw.trim().length === 0) {
    return "absent";
  }
  let data: ClaudeConfig;
  try {
    data = JSON.parse(raw) as ClaudeConfig;
  } catch {
    throw new Error(
      `Failed to parse existing JSON at ${path}. Fix or move it before re-running uninstall.`,
    );
  }
  if (!data.mcpServers?.[SERVER_NAME]) {
    return "absent";
  }
  delete data.mcpServers[SERVER_NAME];
  await writeFileAtomic(path, `${JSON.stringify(data, null, 2)}\n`);
  return "removed";
};

export const removeFromClaudeConfig = (path: string): Promise<RemoveOutcome> =>
  removeFromJsonConfig(path);
export const removeFromClaudeDesktopConfig = (path: string): Promise<RemoveOutcome> =>
  removeFromJsonConfig(path);

export const removeFromCodexConfig = async (path: string): Promise<RemoveOutcome> => {
  if (!existsSync(path)) {
    return "no-file";
  }
  const existing = await readFile(path, "utf8");
  if (!CODEX_SECTION_HEADER_ANYWHERE.test(existing)) {
    return "absent";
  }
  const stripped = stripCodexFormulonSection(existing).replace(/\n*$/, "");
  await writeFileAtomic(path, stripped.length > 0 ? `${stripped}\n` : "");
  return "removed";
};

export const claudeDesktopConfigPath = (): string => {
  const plat = platform();
  if (plat === "darwin") {
    return join(
      homedir(),
      "Library",
      "Application Support",
      "Claude",
      "claude_desktop_config.json",
    );
  }
  if (plat === "win32") {
    const appData = process.env.APPDATA ?? join(homedir(), "AppData", "Roaming");
    return join(appData, "Claude", "claude_desktop_config.json");
  }
  return join(homedir(), ".config", "Claude", "claude_desktop_config.json");
};
