/** OOXML `ST_UnderlineValues`, in the order the engine's ordinals follow. */
export const UNDERLINE = [
  "none",
  "single",
  "double",
  "singleAccounting",
  "doubleAccounting",
] as const;

/** OOXML `ST_VerticalAlignRun`. */
export const VERT_ALIGN_RUN = ["baseline", "superscript", "subscript"] as const;

/** OOXML `ST_PatternType`. */
export const FILL_PATTERN = [
  "none",
  "solid",
  "mediumGray",
  "darkGray",
  "lightGray",
  "darkHorizontal",
  "darkVertical",
  "darkDown",
  "darkUp",
  "darkGrid",
  "darkTrellis",
  "lightHorizontal",
  "lightVertical",
  "lightDown",
  "lightUp",
  "lightGrid",
  "lightTrellis",
  "gray125",
  "gray0625",
] as const;

/** OOXML `ST_BorderStyle`. */
export const BORDER_STYLE = [
  "none",
  "thin",
  "medium",
  "dashed",
  "dotted",
  "thick",
  "double",
  "hair",
  "mediumDashed",
  "dashDot",
  "mediumDashDot",
  "dashDotDot",
  "mediumDashDotDot",
  "slantDashDot",
] as const;

/** OOXML `ST_HorizontalAlignment`. */
export const HORIZONTAL_ALIGN = [
  "general",
  "left",
  "center",
  "right",
  "fill",
  "justify",
  "centerContinuous",
  "distributed",
] as const;

/** OOXML `ST_VerticalAlignment`. */
export const VERTICAL_ALIGN = ["top", "center", "bottom", "justify", "distributed"] as const;

/** OOXML `<scheme>` theme link, in the order the engine's ordinals follow. */
export const FONT_SCHEME = ["none", "major", "minor"] as const;

export const STYLE_VOCABULARY = Object.freeze({
  underline: UNDERLINE,
  vertAlign: VERT_ALIGN_RUN,
  fillPattern: FILL_PATTERN,
  borderStyle: BORDER_STYLE,
  horizontalAlign: HORIZONTAL_ALIGN,
  verticalAlign: VERTICAL_ALIGN,
});

export function ordinal(table: readonly string[], name: string, what: string): number {
  const index = table.indexOf(name);
  if (index < 0) {
    throw new Error(`unknown ${what}: ${name} (expected one of ${table.join(", ")})`);
  }
  return index;
}

/**
 * Parses `#RRGGBB` / `#AARRGGBB` (the leading `#` optional) into the AARRGGBB
 * integer the engine stores. A six-digit colour is fully opaque, which is what
 * Excel writes for a colour picked in its own dialog.
 */
export function parseColor(spec: string): number {
  const hex = spec.trim().replace(/^#/, "");
  if (!(/^[0-9a-fA-F]{6}$/.test(hex) || /^[0-9a-fA-F]{8}$/.test(hex))) {
    throw new Error(`invalid color: ${spec} (expected #RRGGBB or #AARRGGBB)`);
  }
  const value = Number.parseInt(hex, 16);
  return hex.length === 6 ? (value + 0xff000000) >>> 0 : value >>> 0;
}

/**
 * Formats an AARRGGBB integer the way `parseColor` accepts it back: `#RRGGBB`
 * for a fully opaque colour, `#AARRGGBB` when the alpha channel carries
 * anything else.
 */
export function formatColor(argb: number): string {
  const hex = (argb >>> 0).toString(16).padStart(8, "0").toUpperCase();
  return hex.startsWith("FF") ? `#${hex.slice(2)}` : `#${hex}`;
}

export type FontStyleInput = {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  underline?: string;
  vertAlign?: string;
  color?: string;
};

export type FillStyleInput = {
  /** Foreground colour. With the default `solid` pattern this is the fill colour. */
  color?: string;
  bgColor?: string;
  pattern?: string;
};

export type BorderSideInput = string | { style?: string; color?: string };

export type BorderStyleInput = {
  /** Applied to all four sides of every cell in the range — a ruled grid. */
  all?: BorderSideInput;
  /** Applied to the outer edge of the range only — a box around the block. */
  outline?: BorderSideInput;
  left?: BorderSideInput;
  right?: BorderSideInput;
  top?: BorderSideInput;
  bottom?: BorderSideInput;
};

export type AlignStyleInput = {
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
  indent?: number;
  textRotation?: number;
  shrinkToFit?: boolean;
};

export type StyleInput = {
  font?: FontStyleInput;
  fill?: FillStyleInput;
  border?: BorderStyleInput;
  /** Excel format code, for example `#,##0` or `yyyy/mm/dd`. Empty means General. */
  numberFormat?: string;
  align?: AlignStyleInput;
};

export type StyleRect = {
  firstRow: number;
  firstCol: number;
  lastRow: number;
  lastCol: number;
};

/** Where a `style_range` call starts from before applying the requested deltas. */
export type StyleBase = "existing" | "default";
