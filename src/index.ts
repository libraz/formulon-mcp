#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { runInit, runUninstall } from "./init.js";
import { registerAllTools } from "./server/register.js";

/** Single source of truth for the server version: package.json at runtime. */
const PACKAGE_VERSION = ((): string => {
  try {
    const pkgUrl = new URL("../package.json", import.meta.url);
    const pkg = JSON.parse(readFileSync(pkgUrl, "utf8")) as { version?: string };
    return pkg.version ?? "0.0.0";
  } catch {
    return "0.0.0";
  }
})();

const server = new McpServer({
  name: "formulon-mcp",
  version: PACKAGE_VERSION,
});

registerAllTools(server, PACKAGE_VERSION);

const HELP = `formulon-mcp ${PACKAGE_VERSION}
MCP server for Formulon Excel-compatible formula and workbook evaluation.

Usage:
  formulon-mcp            Start the stdio MCP server.
  formulon-mcp init       Interactive setup: write the formulon MCP entry into
                          Claude Code, Codex CLI, and/or Claude Desktop configs.
  formulon-mcp uninstall  Interactive removal: drop the formulon MCP entry from
                          those config files.
  formulon-mcp --help     Show this help.
  formulon-mcp --version  Show version.

Docs: https://github.com/libraz/formulon-mcp
`;

const main = async (): Promise<void> => {
  const argv = process.argv.slice(2);
  if (argv.includes("--help") || argv.includes("-h")) {
    process.stdout.write(HELP);
    return;
  }
  if (argv.includes("--version") || argv.includes("-v")) {
    process.stdout.write(`${PACKAGE_VERSION}\n`);
    return;
  }
  if (argv[0] === "init") {
    await runInit();
    return;
  }
  if (argv[0] === "uninstall") {
    await runUninstall();
    return;
  }
  const transport = new StdioServerTransport();
  await server.connect(transport);
};

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  process.stderr.write(`[formulon-mcp] ${message}\n`);
  process.exit(1);
});
