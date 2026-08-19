import assert from "node:assert/strict";
import { mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import {
  parseTargetChoice,
  previewRemoveImpact,
  previewWriteImpact,
  removeFromClaudeConfig,
  removeFromCodexConfig,
  writeClaudeConfig,
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
