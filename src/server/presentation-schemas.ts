/**
 * Compatibility facade for presentation schemas.
 *
 * The schemas live in the application-independent layout module so session
 * batch operations can validate the same shapes without importing an MCP
 * server module.
 */
export {
  borderSideSchema,
  colorSchema,
  fontStyleSchema,
  nonEmptyPrintSettingsSchema,
  printSettingsSchema,
  styleSchema,
} from "../layout/schemas.js";
