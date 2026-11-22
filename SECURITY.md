# Security policy

## Reporting a vulnerability

Report privately by email to `libraz@libraz.net`, not through the public issue tracker.

Include the tool call and its arguments or the workbook that triggers it, what happened, the server version, your MCP client and its version, the Node version, and a minimal reproduction if you have one. Expect an acknowledgement within a few days.

## Supported versions

Fixes land on the latest release only; older versions are not patched.

## What is in scope

formulon-mcp is a local stdio process that reads and writes workbook files on behalf of an agent, with the filesystem permissions of the process that launched it. It has no directory allowlist or configured root: its tools accept any path the process can reach, so a path the caller named is not itself a finding. In scope:

- Reading or writing a path other than the one the caller passed to a tool, for example through a workbook that redirects a save, a preview, or an image or theme file operation.
- Workbook content that steers the server into file access or network access the caller did not request. The server exposes no HTTP transport and runs no VBA or caller-supplied code.
- A tool that executes code, or reaches an engine method outside the documented allowlist, through crafted input.
- Session handling that lets one session read or change another session's workbook without its `sessionId`.
- A crash, hang or unbounded memory use from a crafted workbook, when the cause is in this server. Engine-level cases belong to the [Formulon security policy](https://github.com/libraz/formulon/blob/main/SECURITY.md).

## What is not in scope

- What the agent itself decides to do with tool results, including overwriting a file it was allowed to write.
- The MCP client, its permission prompts, and how it launches or sandboxes the server.
- Wrong formula results, which are ordinary issues for the [formulon](https://github.com/libraz/formulon/issues) engine.
- Sheet and workbook protection as stored in the file; it is not a security boundary.
