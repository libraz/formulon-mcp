export type { TargetKind } from "./init/cli.js";
export {
  assertKnownProfile,
  defaultLocaleKey,
  PROFILE_LOCALES,
  parseProfileChoice,
  parseTargetChoice,
  runInit,
  runUninstall,
} from "./init/cli.js";
export type { InstallOptions, RemoveOutcome } from "./init/config.js";
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
