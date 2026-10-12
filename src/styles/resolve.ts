import { MAX_COL, MAX_ROW } from "../a1.js";
import {
  assertStatus,
  type BorderRecord,
  type BorderSide,
  type CellXf,
  type FillRecord,
  type FontRecord,
  type Workbook,
} from "../formulon.js";
import type {
  AlignStyleInput,
  BorderSideInput,
  BorderStyleInput,
  FillStyleInput,
  FontStyleInput,
  StyleBase,
  StyleInput,
  StyleRect,
} from "./input.js";
import {
  BORDER_STYLE,
  FILL_PATTERN,
  FONT_SCHEME,
  formatColor,
  HORIZONTAL_ALIGN,
  ordinal,
  parseColor,
  UNDERLINE,
  VERT_ALIGN_RUN,
  VERTICAL_ALIGN,
} from "./input.js";

/**
 * Builds the explicit RGB `ColorSpec` for a caller-supplied colour. A record
 * read back from a file may carry a theme or indexed selector, which stays
 * authoritative over the sibling ARGB field, so an explicit colour has to
 * replace the selector rather than only the fallback.
 */
function rgbColorSpec(argb: number) {
  return { kind: 1, rgb: argb, theme: 0, tint: 0, indexed: 0 };
}

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
  } else if (side.style !== 0 && base.color.kind === 0 && base.colorArgb === 0) {
    next.colorArgb = 0xff000000;
    next.color = rgbColorSpec(0xff000000);
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

/**
 * Counts the cells styling `rect` materializes. Whole rows or columns are
 * styled as rows or columns and materialize none; an outline border needs
 * edges, so it is refused on them.
 */
export function materializedStyleCells(rect: StyleRect, style: StyleInput, ref: string): number {
  const wholeLines =
    (rect.firstRow === 0 && rect.lastRow === MAX_ROW) ||
    (rect.firstCol === 0 && rect.lastCol === MAX_COL);
  if (!wholeLines) {
    return (rect.lastRow - rect.firstRow + 1) * (rect.lastCol - rect.firstCol + 1);
  }
  if (style.border?.outline !== undefined) {
    throw new Error(`an outline border needs a bounded range, not whole rows or columns: ${ref}`);
  }
  return 0;
}

export type StyledRegion = {
  firstRow: number;
  firstCol: number;
  lastRow: number;
  lastCol: number;
  baseXfIndex: number;
  xfIndex: number;
};

export type AppliedStyle = {
  regions: StyledRegion[];
};

/**
 * Applies a declarative style across a rectangle, as Excel's attribute-only
 * formatting does.
 *
 * The rectangle is split by band when an outline border means the edges
 * differ from the interior. Within a band the engine lists the styles the
 * selection resolves to and remaps each to its patched copy, so the change is
 * a delta rather than a replacement: blank cells of a bounded band are
 * materialized so an empty ruled box renders, while a whole column or row
 * restyles the column or row itself.
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
      const resolved = wb.rangeStyleBases(
        sheet,
        rowBand.first,
        colBand.first,
        rowBand.last,
        colBand.last,
      );
      assertStatus(resolved.status, "read range styles");
      const fromXfs = resolved.xfIndices;
      const patched = new Map<number, number>();
      const toXfs = fromXfs.map((xfIndex) => {
        // "default" discards whatever is there, so every base patches xf 0.
        const baseXfIndex = base === "default" ? 0 : xfIndex;
        const known = patched.get(baseXfIndex);
        if (known !== undefined) {
          return known;
        }
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
        patched.set(baseXfIndex, added.index);
        regions.push({
          firstRow: rowBand.first,
          firstCol: colBand.first,
          lastRow: rowBand.last,
          lastCol: colBand.last,
          baseXfIndex,
          xfIndex: added.index,
        });
        return added.index;
      });
      assertStatus(
        wb.remapRangeStyle(
          sheet,
          rowBand.first,
          colBand.first,
          rowBand.last,
          colBand.last,
          fromXfs,
          toXfs,
        ),
        "apply cell format to range",
      );
    }
  }

  return { regions };
}
