export type { TargetKind } from "./init/cli.js";
export { parseTargetChoice, runInit, runUninstall } from "./init/cli.js";
export type { RemoveOutcome } from "./init/config.js";
export {
  claudeDesktopConfigPath,
  displayPath,
  previewRemoveImpact,
  previewWriteImpact,
  removeFromClaudeConfig,
  removeFromClaudeDesktopConfig,
  removeFromCodexConfig,
  writeClaudeConfig,
  writeClaudeDesktopConfig,
  writeCodexConfig,
} from "./init/config.js";
