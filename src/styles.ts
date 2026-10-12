export type {
  AlignStyleInput,
  BorderSideInput,
  BorderStyleInput,
  FillStyleInput,
  FontStyleInput,
  StyleBase,
  StyleInput,
  StyleRect,
} from "./styles/input.js";
export {
  formatColor,
  parseColor,
  STYLE_VOCABULARY,
} from "./styles/input.js";
export type { AppliedStyle, StyledRegion } from "./styles/resolve.js";
export {
  applyDefaultFont,
  applyRangeStyle,
  describeFont,
  materializedStyleCells,
  readDefaultFont,
} from "./styles/resolve.js";
