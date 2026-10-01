import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerAdvancedTools } from "./advanced-tools.js";
import { registerApiTools } from "./api-tools.js";
import { registerAuditTools } from "./audit-tools.js";
import { registerCoreTools } from "./core-tools.js";
import { registerFileTools } from "./file-tools.js";
import { registerLayoutTools } from "./layout-tools.js";
import { registerMutationTools } from "./mutation-tools.js";
import { registerObjectTools } from "./object-tools.js";
import { registerPivotTools } from "./pivot-tools.js";
import { registerPresentationTools } from "./presentation-tools.js";
import { registerPreviewTools } from "./preview-tools.js";
import { registerReadTools } from "./read-tools.js";
import { registerSessionTools } from "./session-tools.js";
import { registerTableTools } from "./table-tools.js";

export function registerAllTools(server: McpServer, version: string): void {
  registerCoreTools(server, version);
  registerSessionTools(server);
  registerMutationTools(server);
  registerReadTools(server);
  registerObjectTools(server);
  registerPresentationTools(server);
  registerAdvancedTools(server);
  registerFileTools(server);
  registerApiTools(server);
  registerTableTools(server);
  registerPivotTools(server);
  registerPreviewTools(server);
  registerLayoutTools(server);
  registerAuditTools(server);
}
