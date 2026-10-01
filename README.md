# formulon-mcp

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-mcp/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-mcp/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-mcp)](https://www.npmjs.com/package/@libraz/formulon-mcp)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-mcp/blob/main/LICENSE)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net/mcp/)

A local stdio MCP server for [Formulon](https://github.com/libraz/formulon). Agents can inspect and edit Excel workbooks, recalculate formulas, author printable documents, preview ranges, audit copied formulas, and create native Tables and PivotTables.

**[Full documentation](https://formulon.libraz.net/mcp/)** lives on the Formulon website. It covers tool inputs, responses, limits, and end-to-end examples.

## Quick start

Requires Node.js 22 or newer; see [`engines`](https://github.com/libraz/formulon-mcp/blob/main/package.json). Register the server with the bundled interactive installer:

```sh
npx -y @libraz/formulon-mcp init
```

Choose Claude Code, Codex CLI, or Claude Desktop, review the configuration changes, and restart the client. Ask the agent to evaluate `=SUM(10,20,30)` using Formulon; the result is `60`.

See [installation](https://formulon.libraz.net/mcp/install) for manual client configuration, local checkouts, and troubleshooting. Remove registration with `npx -y @libraz/formulon-mcp uninstall`.

## Guides

| Task | Documentation |
| --- | --- |
| Edit a workbook template and save a copy | [Workbook workflow](https://formulon.libraz.net/mcp/workflow) |
| Create an invoice with formulas and print settings | [Document authoring](https://formulon.libraz.net/mcp/authoring) |
| Reproduce a form and inspect its appearance | [Layout and previews](https://formulon.libraz.net/mcp/layout-preview) |
| Find overwritten or missing formulas | [Formula audit](https://formulon.libraz.net/mcp/formula-audit) |
| Summarize sales data | [Tables and pivots](https://formulon.libraz.net/mcp/tables-pivots) |
| Inspect tool inputs or discover a low-level method | [Tools](https://formulon.libraz.net/mcp/tools) · [Advanced API](https://formulon.libraz.net/mcp/advanced) |

Sessions live in memory and disappear when the server stops. Use explicit output paths and review load/save diagnostics. Preview rendering is approximate; formula audits report review candidates. See the [security model](https://formulon.libraz.net/mcp/security).

## Development

```sh
yarn install
yarn run check
yarn test
```

Tests build and exercise `dist/`. The calculation engine is maintained in the separate [Formulon repository](https://github.com/libraz/formulon).

## License

Apache-2.0. See [LICENSE](https://github.com/libraz/formulon-mcp/blob/main/LICENSE).
