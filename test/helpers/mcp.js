import assert from "node:assert/strict";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

export function textPayload(result) {
  assert.notEqual(result.isError, true, result.content[0]?.text);
  assert.equal(result.content[0].type, "text");
  return JSON.parse(result.content[0].text);
}

export function errorPayload(result) {
  assert.equal(result.isError, true);
  assert.equal(result.content[0].type, "text");
  return result.content[0].text;
}

export function callTool(client, name, arguments_ = {}) {
  return client.callTool({ name, arguments: arguments_ });
}

/**
 * Starts the server over stdio and runs `fn` with a connected client. `options.env`
 * is merged over the transport's default environment, which otherwise leaves out
 * variables such as FORMULON_MCP_PROFILE.
 */
export async function withClient(fn, options = {}) {
  const client = new Client({ name: "formulon-mcp-test", version: "0.1.0" });
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: ["./dist/index.js"],
    cwd: process.cwd(),
    stderr: "pipe",
    ...(options.env ? { env: options.env } : {}),
  });
  await client.connect(transport);
  try {
    return await fn(client);
  } finally {
    await client.close();
  }
}
