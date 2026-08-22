import {
  assertStatus,
  type BorderRecord,
  type BorderSide,
  type CellXf,
  type FillRecord,
  type FontRecord,
  type Workbook,
} from "./formulon.js";

/**
 * Declarative styling on top of Formulon's style tables.
 *
 * The engine stores styles the way OOXML does: a cell carries one `xf` index,
 * and that record points at rows of the font, fill, border and number-format
 * tables, each of which is a complete record of ordinals. Authoring a styled
 * document through that surface means assembling four full records and knowing
 * five ordinal tables by heart, so this module accepts names and hex colours
 * and does the assembly, reusing the engine's own deduplication.
 */

/** OOXML `ST_UnderlineValues`, in the order the engine's ordinals follow. */
const UNDERLINE = ["none", "single", "double", "singleAccounting", "doubleAccounting"] as const;

/** OOXML `ST_VerticalAlignRun`. */
const VERT_ALIGN_RUN = ["baseline", "superscript", "subscript"] as const;

/** OOXML `ST_PatternType`. */
const FILL_PATTERN = [
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
const BORDER_STYLE = [
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
const HORIZONTAL_ALIGN = [
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
const VERTICAL_ALIGN = ["top", "center", "bottom", "justify", "distributed"] as const;

/** OOXML `<scheme>` theme link, in the order the engine's ordinals follow. */
const FONT_SCHEME = ["none", "major", "minor"] as const;

export const STYLE_VOCABULARY = Object.freeze({
  underline: UNDERLINE,
  vertAlign: VERT_ALIGN_RUN,
  fillPattern: FILL_PATTERN,
  borderStyle: BORDER_STYLE,
  horizontalAlign: HORIZONTAL_ALIGN,
  verticalAlign: VERTICAL_ALIGN,
});

function ordinal(table: readonly string[], name: string, what: string): number {
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

/**
 * Builds the explicit RGB `ColorSpec` for a caller-supplied colour. A record
 * read back from a file may carry a theme or indexed selector, which stays
 * authoritative over the sibling ARGB field, so an explicit colour has to
 * replace the selector rather than only the fallback.
 */
function rgbColorSpec(argb: number) {
  return { kind: 1, rgb: argb, theme: 0, tint: 0, indexed: 0 };
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

function applyFont(base: FontRecord, input: FontStyleInput): FontRecord {
  const font: FontRecord = { ...base, color: { ...base.color } };
  if (input.name !== undefined) {
    font.name = input.name;
    // A `<scheme>` link tells Excel to take the typeface from the theme, which
    // would re-resolve the name just stated as soon as the theme changes. The
    // base font of a ja-JP workbook carries `minor`, so the link has to be cut
    // rather than inherited — Excel drops the element the same way when a font
    // is picked by name.
    font.scheme = 0;
  }
  if (input.size !== undefined) {
    if (!(input.size > 0)) {
      throw new Error(`invalid font size: ${input.size}`);
    }
    font.size = input.size;
  }
  // The `has*` flags say whether the element is written at all; an explicit
  // false has to be stated, or the writer omits it and Excel inherits instead.
  if (input.bold !== undefined) {
    font.bold = input.bold;
    font.hasBold = true;
  }
  if (input.italic !== undefined) {
    font.italic = input.italic;
    font.hasItalic = true;
  }
  if (input.strike !== undefined) {
    font.strike = input.strike;
    font.hasStrike = true;
  }
  if (input.underline !== undefined) {
    font.underline = ordinal(UNDERLINE, input.underline, "underline style");
  }
  if (input.vertAlign !== undefined) {
    font.vertAlign = ordinal(VERT_ALIGN_RUN, input.vertAlign, "font vertical alignment");
  }
  if (input.color !== undefined) {
    const argb = parseColor(input.color);
    font.colorArgb = argb;
    font.color = rgbColorSpec(argb);
  }
  return font;
}

/** Reports a font record in the vocabulary `FontStyleInput` accepts back. */
export function describeFont(font: FontRecord) {
  return {
    name: font.name,
    size: font.size,
    bold: font.bold,
    italic: font.italic,
    strike: font.strike,
    underline: UNDERLINE[font.underline] ?? font.underline,
    vertAlign: VERT_ALIGN_RUN[font.vertAlign] ?? font.vertAlign,
    color: formatColor(font.colorArgb),
    // Reported because it decides whether `name` survives a theme change: a
    // font linked to the theme is re-resolved from it rather than kept.
    themeLink: FONT_SCHEME[font.scheme] ?? font.scheme,
  };
}

/**
 * The workbook default font: font 0, the record every unstyled cell resolves
 * to. The style table always owns the slot, so it is set in place rather than
 * appended — `addFont` can only add a font beside it.
 */
const DEFAULT_FONT_INDEX = 0;

/** Reads the workbook default font. */
export function readDefaultFont(wb: Workbook) {
  const font = wb.getFont(DEFAULT_FONT_INDEX);
  assertStatus(font.status, "read default font");
  return describeFont(record(font));
}

/**
 * Redeclares the workbook default font from the deltas in `input`, leaving
 * every property the call does not state as it was. Every cell that was never
 * styled follows, which is what a workbook needs when its text is Japanese and
 * the seeded default is Excel's Calibri.
 */
export function applyDefaultFont(wb: Workbook, input: FontStyleInput) {
  const base = wb.getFont(DEFAULT_FONT_INDEX);
  assertStatus(base.status, "read default font");
  const font = applyFont(record(base), input);
  assertStatus(wb.setDefaultFont(font), "set default font");
  return describeFont(font);
}

function applyFill(base: FillRecord, input: FillStyleInput): FillRecord {
  const fill: FillRecord = { ...base, fg: { ...base.fg }, bg: { ...base.bg } };
  if (input.color !== undefined) {
    const argb = parseColor(input.color);
    fill.fgArgb = argb;
    fill.fg = rgbColorSpec(argb);
    // A colour with no stated pattern means "paint the cell", which OOXML
    // spells as a solid pattern whose foreground carries the colour.
    if (input.pattern === undefined && fill.pattern === 0) {
      fill.pattern = 1;
    }
  }
  if (input.bgColor !== undefined) {
    const argb = parseColor(input.bgColor);
    fill.bgArgb = argb;
    fill.bg = rgbColorSpec(argb);
  }
  if (input.pattern !== undefined) {
    fill.pattern = ordinal(FILL_PATTERN, input.pattern, "fill pattern");
  }
  return fill;
}

function normalizeSide(input: BorderSideInput): { style: number; argb?: number } {
  const spec = typeof input === "string" ? { style: input } : input;
  const style = ordinal(BORDER_STYLE, spec.style ?? "thin", "border style");
  return { style, argb: spec.color === undefined ? undefined : parseColor(spec.color) };
}

function withSide(base: BorderSide, side: { style: number; argb?: number }): BorderSide {
  const next: BorderSide = { ...base, color: { ...base.color }, style: side.style };
  if (side.argb !== undefined) {
    next.colorArgb = side.argb;
    next.color = rgbColorSpec(side.argb);
  }
  return next;
}

/** One band of a range split: an edge row/column, or the span between edges. */
type Band = { first: number; last: number; atStart: boolean; atEnd: boolean };

/**
 * Splits an axis into leading edge, interior and trailing edge so an outline
 * border can be drawn without styling each cell individually. A one-cell axis
 * yields a single band that is both edges; a two-cell axis has no interior.
 */
function bands(first: number, last: number): Band[] {
  if (first === last) {
    return [{ first, last, atStart: true, atEnd: true }];
  }
  const edges: Band[] = [
    { first, last: first, atStart: true, atEnd: false },
    { first: last, last, atStart: false, atEnd: true },
  ];
  if (last === first + 1) {
    return edges;
  }
  return [edges[0], { first: first + 1, last: last - 1, atStart: false, atEnd: false }, edges[1]];
}

function buildBorder(
  base: BorderRecord,
  input: BorderStyleInput,
  rowBand: Band,
  colBand: Band,
): BorderRecord {
  let border: BorderRecord = {
    ...base,
    left: { ...base.left, color: { ...base.left.color } },
    right: { ...base.right, color: { ...base.right.color } },
    top: { ...base.top, color: { ...base.top.color } },
    bottom: { ...base.bottom, color: { ...base.bottom.color } },
    diagonal: { ...base.diagonal, color: { ...base.diagonal.color } },
  };
  if (input.all !== undefined) {
    const side = normalizeSide(input.all);
    border = {
      ...border,
      left: withSide(border.left, side),
      right: withSide(border.right, side),
      top: withSide(border.top, side),
      bottom: withSide(border.bottom, side),
    };
  }
  for (const key of ["left", "right", "top", "bottom"] as const) {
    const spec = input[key];
    if (spec !== undefined) {
      border[key] = withSide(border[key], normalizeSide(spec));
    }
  }
  if (input.outline !== undefined) {
    const side = normalizeSide(input.outline);
    if (rowBand.atStart) {
      border.top = withSide(border.top, side);
    }
    if (rowBand.atEnd) {
      border.bottom = withSide(border.bottom, side);
    }
    if (colBand.atStart) {
      border.left = withSide(border.left, side);
    }
    if (colBand.atEnd) {
      border.right = withSide(border.right, side);
    }
  }
  return border;
}

function applyAlignment(xf: CellXf, input: AlignStyleInput): CellXf {
  const next: CellXf = { ...xf };
  let touched = false;
  if (input.horizontal !== undefined) {
    next.horizontalAlign = ordinal(HORIZONTAL_ALIGN, input.horizontal, "horizontal alignment");
    next.hasHorizontalAlign = true;
    touched = true;
  }
  if (input.vertical !== undefined) {
    next.verticalAlign = ordinal(VERTICAL_ALIGN, input.vertical, "vertical alignment");
    next.hasVerticalAlign = true;
    touched = true;
  }
  if (input.wrapText !== undefined) {
    next.wrapText = input.wrapText;
    next.hasWrapText = true;
    touched = true;
  }
  if (input.indent !== undefined) {
    next.indent = input.indent;
    touched = true;
  }
  if (input.textRotation !== undefined) {
    next.textRotation = input.textRotation;
    touched = true;
  }
  if (input.shrinkToFit !== undefined) {
    next.shrinkToFit = input.shrinkToFit;
    touched = true;
  }
  if (touched) {
    next.hasAlignment = true;
  }
  return next;
}

/** Strips the `status` field engine getters carry so a record round-trips into a setter. */
function record<T>(value: T & { status?: unknown }): T {
  const { status: _status, ...rest } = value;
  return rest as T;
}

function addStyle(
  wb: Workbook,
  kind: "font" | "fill" | "border",
  entry: FontRecord | FillRecord | BorderRecord,
): number {
  const result =
    kind === "font"
      ? wb.addFont(entry as FontRecord)
      : kind === "fill"
        ? wb.addFill(entry as FillRecord)
        : wb.addBorder(entry as BorderRecord);
  assertStatus(result.status, `add ${kind}`);
  return result.index;
}

/** The style properties a base `xf` contributes once the deltas are applied. */
type DerivedStyle = {
  xf: CellXf;
  fontIndex: number;
  fillIndex: number;
  numFmtId: number;
  border: BorderRecord;
};

/**
 * Resolves what one existing style becomes under `style`. Memoized per base
 * index because a range typically holds only a handful of distinct styles, and
 * every property except the border is the same for all cells sharing one.
 */
function deriveStyle(
  wb: Workbook,
  baseXfIndex: number,
  style: StyleInput,
  cache: Map<number, DerivedStyle>,
): DerivedStyle {
  const cached = cache.get(baseXfIndex);
  if (cached) {
    return cached;
  }
  const baseXf = wb.getCellXf(baseXfIndex);
  assertStatus(baseXf.status, `read style ${baseXfIndex}`);

  let fontIndex = baseXf.fontIndex;
  if (style.font) {
    const baseFont = wb.getFont(baseXf.fontIndex);
    assertStatus(baseFont.status, `read font ${baseXf.fontIndex}`);
    fontIndex = addStyle(wb, "font", applyFont(record(baseFont), style.font));
  }

  let fillIndex = baseXf.fillIndex;
  if (style.fill) {
    const baseFill = wb.getFill(baseXf.fillIndex);
    assertStatus(baseFill.status, `read fill ${baseXf.fillIndex}`);
    fillIndex = addStyle(wb, "fill", applyFill(record(baseFill), style.fill));
  }

  let numFmtId = baseXf.numFmtId;
  if (style.numberFormat !== undefined) {
    if (style.numberFormat === "") {
      numFmtId = 0;
    } else {
      const added = wb.addNumFmt(style.numberFormat);
      assertStatus(added.status, "add number format");
      numFmtId = added.numFmtId;
    }
  }

  const baseBorder = wb.getBorder(baseXf.borderIndex);
  assertStatus(baseBorder.status, `read border ${baseXf.borderIndex}`);

  const derived: DerivedStyle = {
    xf: record(baseXf),
    fontIndex,
    fillIndex,
    numFmtId,
    border: record(baseBorder),
  };
  cache.set(baseXfIndex, derived);
  return derived;
}

type CellAddress = { row: number; col: number };

/**
 * Groups a band's cells by the style they already carry, so a restyle keeps
 * what it was not asked to change. Excel formats this way — making a range bold
 * does not reset the number format under it — and flattening a band to its
 * top-left cell's style would silently undo an earlier pass.
 */
function groupByBaseStyle(
  wb: Workbook,
  sheet: number,
  rowBand: Band,
  colBand: Band,
): Map<number, CellAddress[]> {
  const groups = new Map<number, CellAddress[]>();
  for (let row = rowBand.first; row <= rowBand.last; row += 1) {
    for (let col = colBand.first; col <= colBand.last; col += 1) {
      const cell = wb.getCellXfIndex(sheet, row, col);
      assertStatus(cell.status, `read style at ${row},${col}`);
      const group = groups.get(cell.xfIndex);
      if (group) {
        group.push({ row, col });
      } else {
        groups.set(cell.xfIndex, [{ row, col }]);
      }
    }
  }
  return groups;
}

export type StyledRegion = {
  firstRow: number;
  firstCol: number;
  lastRow: number;
  lastCol: number;
  baseXfIndex: number;
  xfIndex: number;
  cellCount: number;
};

export type AppliedStyle = {
  regions: StyledRegion[];
};

/**
 * Applies a declarative style across a rectangle.
 *
 * The rectangle is split two ways: by band, when an outline border means the
 * edges differ from the interior, and by the style each cell already carries,
 * so the requested change is a delta rather than a replacement. A band whose
 * cells all share one style is written with a single range call, which
 * materializes blank cells so an empty ruled box renders.
 */
export function applyRangeStyle(
  wb: Workbook,
  sheet: number,
  rect: StyleRect,
  style: StyleInput,
  base: StyleBase,
): AppliedStyle {
  const borderInput = style.border;
  // Without an outline the border is the same everywhere, so the rectangle
  // stays one band on each axis.
  const rowBands =
    borderInput?.outline === undefined
      ? [{ first: rect.firstRow, last: rect.lastRow, atStart: true, atEnd: true }]
      : bands(rect.firstRow, rect.lastRow);
  const colBands =
    borderInput?.outline === undefined
      ? [{ first: rect.firstCol, last: rect.lastCol, atStart: true, atEnd: true }]
      : bands(rect.firstCol, rect.lastCol);

  const cache = new Map<number, DerivedStyle>();
  const regions: StyledRegion[] = [];
  for (const rowBand of rowBands) {
    for (const colBand of colBands) {
      const bandCells = (rowBand.last - rowBand.first + 1) * (colBand.last - colBand.first + 1);
      // "default" discards whatever is there, so the band needs no scan.
      const groups =
        base === "default"
          ? new Map<number, CellAddress[] | null>([[0, null]])
          : (groupByBaseStyle(wb, sheet, rowBand, colBand) as Map<number, CellAddress[] | null>);

      for (const [baseXfIndex, cells] of groups) {
        const derived = deriveStyle(wb, baseXfIndex, style, cache);
        const borderIndex = borderInput
          ? addStyle(wb, "border", buildBorder(derived.border, borderInput, rowBand, colBand))
          : derived.xf.borderIndex;
        const added = wb.addXf(
          applyAlignment(
            {
              ...derived.xf,
              fontIndex: derived.fontIndex,
              fillIndex: derived.fillIndex,
              borderIndex,
              numFmtId: derived.numFmtId,
            },
            style.align ?? {},
          ),
        );
        assertStatus(added.status, "add cell format");

        if (cells === null || cells.length === bandCells) {
          assertStatus(
            wb.setRangeXfIndex(
              sheet,
              rowBand.first,
              colBand.first,
              rowBand.last,
              colBand.last,
              added.index,
            ),
            "apply cell format to range",
          );
          regions.push({
            firstRow: rowBand.first,
            firstCol: colBand.first,
            lastRow: rowBand.last,
            lastCol: colBand.last,
            baseXfIndex,
            xfIndex: added.index,
            cellCount: bandCells,
          });
          continue;
        }

        for (const cell of cells) {
          // One-cell range rather than setCellXfIndex, so a blank cell in a
          // mixed band is materialized the same way a whole-band write does.
          assertStatus(
            wb.setRangeXfIndex(sheet, cell.row, cell.col, cell.row, cell.col, added.index),
            "apply cell format",
          );
        }
        regions.push({
          firstRow: Math.min(...cells.map((cell) => cell.row)),
          firstCol: Math.min(...cells.map((cell) => cell.col)),
          lastRow: Math.max(...cells.map((cell) => cell.row)),
          lastCol: Math.max(...cells.map((cell) => cell.col)),
          baseXfIndex,
          xfIndex: added.index,
          cellCount: cells.length,
        });
      }
    }
  }

  return { regions };
}
