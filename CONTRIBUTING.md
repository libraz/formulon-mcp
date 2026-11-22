# Contributing to formulon-mcp

Contributions are welcome, including small and unfinished ones. A maintainer reviews every pull request, finishes what is missing, and merges it, so you do not need to get everything in this document right, or even read all of it.

## What a pull request needs

- A description of what the change does and why. If the intent is clear, the maintainer can verify the behaviour.
- If you can, a build and the tests for the part you changed.

That is all. A red check, a missing test, or a missing type declaration is fine: say what you know in the pull request, and the maintainer will finish it or tell you what is left. Target `main`; this repository has no other long-lived branch.

For a larger change, such as a new tool or a change to a tool's input shape, opening an issue first to describe the idea can save you work. It is not required.

Reporting a bug you do not intend to fix is a useful contribution too, and needs none of the setup below.

### What the maintainer can verify

The maintainer tests on macOS, and on Windows in Japanese only. If your change depends on Windows in another language (a path convention, how an MCP client launches the server over stdio), describe what you see there, because it cannot be reproduced on the maintainer's side.

## What you are contributing to

A thin MCP server over the [Formulon](https://github.com/libraz/formulon) engine. It registers tools that open, inspect, edit, recalculate and save workbooks, keeps sessions in memory, and ships an `init` / `uninstall` installer for MCP client configuration.

- `src/server` defines the tools: names, descriptions and input schemas (zod).
- `src/session` holds the in-memory workbook sessions the tools operate on.
- `src/init` is the installer for Claude Code, Codex CLI and Claude Desktop.
- `test` holds the tests, which run against the built `dist/`.

A wrong calculation result, or a file that loads or saves wrongly, belongs to the engine; report it in the [formulon repository](https://github.com/libraz/formulon/issues). Tool shape, tool descriptions and session handling belong here.

The rest of this document has two parts. [Checklist](#checklist) describes what a finished change looks like. [Reference](#reference) explains how the checks behave.

## Development setup

Node 22 and Yarn 4 (the versions are pinned in `mise.toml` and `package.json`). Use Yarn; npm, pnpm and bun are not supported in this repository.

```bash
yarn install
yarn check        # Biome check and tsc --noEmit
yarn test         # builds dist/, then runs vitest
yarn format       # Biome, writes fixes
```

## What CI requires

A red check on your pull request is not a reason to hold it back; the maintainer can look at it with you.

| Workflow      | Runs on                                                                                                                                                              | What it checks                                       |
| ------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------- |
| `ci.yml`      | pushes to `main` and pull requests to `main` that change `src/`, `test/`, `package.json`, `yarn.lock`, `tsconfig.json`, `biome.json` or the workflow itself; Node 22 | `yarn check`, `yarn build`, `yarn test`              |
| `publish.yml` | version tags                                                                                                                                                         | publishes the package to npm and creates the release |

## Checklist

This is what a change looks like by the time it is merged. The maintainer checks it and fills in what is missing, so you are not expected to know all of it.

- Name a tool `formulon_<verb>_<noun>`, register it with a zod input schema, and describe every input; the description is what the agent reads.
- A failure goes back to the agent as a tool error through `fail`, not as a thrown exception.
- Fix an engine-level defect in the engine repository and pull it in with a version bump rather than working around it here.
- Add a test under `test` for the code you changed. Tests exercise `dist/`, so rebuild before running a single spec.
- Where a change fixes a defect, include a test that fails without the fix.
- Write comments in English and give exported symbols a doc comment.

### Commit messages

Conventional Commits, in English: `type(scope): summary`, lowercase after the colon, no trailing period. A body, when there is one, is a blank line followed by `- ` bullets. Commit messages can be rewritten at merge, so do not let them hold up a pull request.

## Reference

Nothing in this section asks anything of you; read it when a checklist item or a check puzzles you.

- `yarn test` and `yarn test:coverage` build first. Running `vitest` directly uses whatever `dist/` already holds, so a source edit that is not rebuilt is not under test.
- `ci.yml` does not run for a change that touches only documentation, such as `README.md` or this file; the maintainer reviews those by reading.
- Sessions live in memory and disappear when the server stops; a test that needs a workbook opens its own session.

## Reporting issues

Search the existing issues first, then include what you did, what happened, what you expected, the server version, your MCP client and its version, the Node version, and your OS and its display language. Include the tool call and its arguments. For a problem with a loaded file, the file matters more than anything else in the report.

Security vulnerabilities go through the process in [SECURITY.md](SECURITY.md), not the public issue tracker.

## Licence

By contributing, you agree that your contributions are licensed under the Apache License 2.0.
