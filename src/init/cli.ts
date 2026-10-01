import { homedir } from "node:os";
import { join } from "node:path";
import { stdin, stdout } from "node:process";
import { createInterface, type Interface as ReadlineInterface } from "node:readline/promises";
import {
  claudeDesktopConfigPath,
  displayPath,
  previewRemoveImpact,
  previewWriteImpact,
  type RemoveOutcome,
  removeFromCodexConfig,
  removeFromJsonConfig,
  SERVER_NAME,
  writeCodexConfig,
  writeJsonConfig,
} from "./config.js";

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
