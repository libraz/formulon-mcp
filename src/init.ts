import { existsSync } from "node:fs";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { homedir, platform } from "node:os";
import { dirname, join } from "node:path";
import { stdin, stdout } from "node:process";
import { createInterface, type Interface as ReadlineInterface } from "node:readline/promises";

const PACKAGE_SPEC = "@libraz/formulon-mcp";
const SERVER_NAME = "formulon";

type McpServerEntry = {
  command: string;
  args: string[];
};

type ClaudeConfig = {
  mcpServers?: Record<string, McpServerEntry>;
  [key: string]: unknown;
};

const buildEntry = (): McpServerEntry => ({
  command: "npx",
  args: ["-y", PACKAGE_SPEC],
});

const writeJsonConfig = async (path: string): Promise<void> => {
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
  data.mcpServers[SERVER_NAME] = buildEntry();
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, `${JSON.stringify(data, null, 2)}\n`);
};

export const writeClaudeConfig = (path: string): Promise<void> => writeJsonConfig(path);
export const writeClaudeDesktopConfig = (path: string): Promise<void> => writeJsonConfig(path);

const stripCodexFormulonSection = (content: string): string => {
  const lines = content.split("\n");
  const out: string[] = [];
  let inBlock = false;
  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed === `[mcp_servers.${SERVER_NAME}]`) {
      inBlock = true;
      continue;
    }
    if (inBlock) {
      if (trimmed.startsWith("[") && trimmed.endsWith("]")) {
        inBlock = false;
        out.push(line);
      }
      continue;
    }
    out.push(line);
  }
  return out.join("\n");
};

export const writeCodexConfig = async (path: string): Promise<void> => {
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
  ].join("\n");
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, `${prefix}${block}\n`);
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
    return new RegExp(`^\\[mcp_servers\\.${SERVER_NAME}\\]\\s*$`, "m").test(raw)
      ? `(replace ${SERVER_NAME})`
      : "(merge)";
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
    return new RegExp(`^\\[mcp_servers\\.${SERVER_NAME}\\]\\s*$`, "m").test(raw)
      ? `(remove ${SERVER_NAME})`
      : `(no ${SERVER_NAME}; skip)`;
  }
};

export type RemoveOutcome = "removed" | "absent" | "no-file";

const removeFromJsonConfig = async (path: string): Promise<RemoveOutcome> => {
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
  await writeFile(path, `${JSON.stringify(data, null, 2)}\n`);
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
  if (!new RegExp(`^\\[mcp_servers\\.${SERVER_NAME}\\]\\s*$`, "m").test(existing)) {
    return "absent";
  }
  const stripped = stripCodexFormulonSection(existing).replace(/\n*$/, "");
  await writeFile(path, stripped.length > 0 ? `${stripped}\n` : "");
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

const promptYesNo = async (
  rl: ReadlineInterface,
  question: string,
  defaultYes: boolean,
): Promise<boolean> => {
  const def = defaultYes ? "Y/n" : "y/N";
  const ans = (await rl.question(`${question} [${def}] `)).trim().toLowerCase();
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

const writeForKind = (kind: TargetKind, path: string): (() => Promise<void>) => {
  if (kind === "codex") {
    return () => writeCodexConfig(path);
  }
  return () => writeJsonConfig(path);
};

const removeForKind = (kind: TargetKind, path: string): (() => Promise<RemoveOutcome>) => {
  if (kind === "codex") {
    return () => removeFromCodexConfig(path);
  }
  return () => removeFromJsonConfig(path);
};

type Target = { label: string; path: string; write: () => Promise<void> };

const pickTargets = (kinds: TargetKind[]): Target[] => {
  const targets: Target[] = [];
  for (const opt of TARGET_OPTIONS) {
    if (!kinds.includes(opt.kind)) {
      continue;
    }
    const path = opt.resolvePath();
    targets.push({ label: opt.label, path, write: writeForKind(opt.kind, path) });
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

export const runInit = async (): Promise<void> => {
  const rl = createInterface({ input: stdin, output: stdout });
  try {
    stdout.write("formulon-mcp setup\n\n");
    stdout.write("Where to install? (pick one or more, comma-separated)\n");
    stdout.write(`${renderTargetMenu()}\n`);
    const choice = (await rl.question("Choice [1]: ")).trim();
    const kinds = parseTargetChoice(choice, "1");
    const targets = pickTargets(kinds);

    stdout.write("\nWill update:\n");
    for (const t of targets) {
      const summary = await previewWriteImpact(t.path);
      stdout.write(`  - ${displayPath(t.path)} ${summary}\n`);
    }
    stdout.write("\n");

    const confirmed = await promptYesNo(rl, "Proceed?", true);
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
    rl.close();
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
  const rl = createInterface({ input: stdin, output: stdout });
  try {
    stdout.write("formulon-mcp uninstall\n\n");
    stdout.write("Where to remove from? (pick one or more, comma-separated)\n");
    stdout.write(`${renderTargetMenu()}\n`);
    const choice = (await rl.question("Choice [1,2,3,4]: ")).trim();
    const kinds = parseTargetChoice(choice, "1,2,3,4");
    const targets = pickRemoveTargets(kinds);

    stdout.write("\nWill update:\n");
    for (const t of targets) {
      const summary = await previewRemoveImpact(t.path);
      stdout.write(`  - ${displayPath(t.path)} ${summary}\n`);
    }
    stdout.write("\n");

    const confirmed = await promptYesNo(rl, "Proceed?", true);
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
    rl.close();
  }
};
