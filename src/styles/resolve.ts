import { ColorKind } from "@libraz/formulon";
import { MAX_COL, MAX_ROW } from "../a1.js";
import {
  type Border,
  type BorderSide,
  type CellXf,
  ColorContext,
  type Fill,
  type Font,
  sheetAt,
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

type FontPatch = Parameters<Workbook["ooxml"]["addFont"]>[0];
type FillPatch = Parameters<Workbook["ooxml"]["addFill"]>[0];
type BorderPatch = Parameters<Workbook["ooxml"]["addBorder"]>[0];
type CellXfPatch = Parameters<Workbook["ooxml"]["addCellXf"]>[0];

/** The colour string the engine takes for a caller-supplied colour. */
function colorText(spec: string): string {
  return formatColor(parseColor(spec));
}

function applyFont(base: Font, input: FontStyleInput): FontPatch {
  const font: { -readonly [K in keyof FontPatch]: FontPatch[K] } = { ...base };
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
    font.sizePt = input.size;
  }
  // An explicit false has to be stated: an absent element makes Excel inherit instead.
  if (input.bold !== undefined) {
    font.bold = input.bold;
  }
  if (input.italic !== undefined) {
    font.italic = input.italic;
  }
  if (input.strike !== undefined) {
    font.strike = input.strike;
  }
  if (input.underline !== undefined) {
    font.underline = ordinal(UNDERLINE, input.underline, "underline style");
  }
  if (input.vertAlign !== undefined) {
    font.vertAlign = ordinal(VERT_ALIGN_RUN, input.vertAlign, "font vertical alignment");
  }
  if (input.color !== undefined) {
    font.color = colorText(input.color);
  }
  return font;
}

/** Reports a font record in the vocabulary `FontStyleInput` accepts back. */
export function describeFont(wb: Workbook, font: Font) {
  const color =
    font.color.kind === ColorKind.Rgb
      ? font.color.argb
      : wb.resolveColor(font.color, ColorContext.Font).argb;
  return {
    name: font.name,
    size: font.sizePt,
    bold: font.bold ?? false,
    italic: font.italic ?? false,
    strike: font.strike ?? false,
    underline: UNDERLINE[font.underline] ?? font.underline,
    vertAlign: VERT_ALIGN_RUN[font.vertAlign] ?? font.vertAlign,
    color: formatColor(color),
    // Reported because it decides whether `name` survives a theme change: a
    // font linked to the theme is re-resolved from it rather than kept.
    themeLink: FONT_SCHEME[font.scheme] ?? font.scheme,
  };
}

/** Reads the workbook default font. */
export function readDefaultFont(wb: Workbook) {
  return describeFont(wb, wb.getDefaultFont());
}

/**
 * Redeclares the workbook default font from the deltas in `input`, leaving
 * every property the call does not state as it was. Every cell that was never
 * styled follows, which is what a workbook needs when its text is Japanese and
 * the seeded default is Excel's Calibri.
 */
export function applyDefaultFont(wb: Workbook, input: FontStyleInput) {
  wb.setDefaultFont(applyFont(wb.getDefaultFont(), input));
  return describeFont(wb, wb.getDefaultFont());
}

function applyFill(base: Fill, input: FillStyleInput): FillPatch {
  const fill: { -readonly [K in keyof FillPatch]: FillPatch[K] } = { ...base };
  if (input.color !== undefined) {
    fill.foreground = colorText(input.color);
    // A colour with no stated pattern means "paint the cell", which OOXML
    // spells as a solid pattern whose foreground carries the colour.
    if (input.pattern === undefined && base.pattern === 0) {
      fill.pattern = 1;
    }
  }
  if (input.bgColor !== undefined) {
    fill.background = colorText(input.bgColor);
  }
  if (input.pattern !== undefined) {
    fill.pattern = ordinal(FILL_PATTERN, input.pattern, "fill pattern");
  }
  return fill;
}

type EngineSide = NonNullable<BorderPatch["left"]>;

const BLACK = "#000000";

type NormalizedSide = { style: BorderSide["style"]; color?: string };

function normalizeSide(input: BorderSideInput): NormalizedSide {
  const spec = typeof input === "string" ? { style: input, color: undefined } : input;
  const style = ordinal<BorderSide["style"]>(BORDER_STYLE, spec.style ?? "thin", "border style");
  return { style, color: spec.color === undefined ? undefined : colorText(spec.color) };
}

function withSide(base: BorderSide, side: NormalizedSide): EngineSide {
  if (side.color !== undefined) {
    return { style: side.style, color: side.color };
  }
  if (side.style !== 0 && base.color.kind === ColorKind.None) {
    return { style: side.style, color: BLACK };
  }
  return { ...base, style: side.style };
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
  base: Border,
  input: BorderStyleInput,
  rowBand: Band,
  colBand: Band,
): BorderPatch {
  const border: { -readonly [K in keyof Border]: BorderPatch[K] } = { ...base };
  const sides = ["left", "right", "top", "bottom"] as const;
  if (input.all !== undefined) {
    const side = normalizeSide(input.all);
    for (const key of sides) {
      border[key] = withSide(base[key], side);
    }
  }
  for (const key of sides) {
    const spec = input[key];
    if (spec !== undefined) {
      border[key] = withSide(base[key], normalizeSide(spec));
    }
  }
  if (input.outline !== undefined) {
    const side = normalizeSide(input.outline);
    if (rowBand.atStart) {
      border.top = withSide(base.top, side);
    }
    if (rowBand.atEnd) {
      border.bottom = withSide(base.bottom, side);
    }
    if (colBand.atStart) {
      border.left = withSide(base.left, side);
    }
    if (colBand.atEnd) {
      border.right = withSide(base.right, side);
    }
  }
  return border;
}

function applyAlignment(xf: CellXfPatch, input: AlignStyleInput): CellXfPatch {
  const next: { -readonly [K in keyof CellXfPatch]: CellXfPatch[K] } = { ...xf };
  let touched = false;
  if (input.horizontal !== undefined) {
    next.horizontal = ordinal(HORIZONTAL_ALIGN, input.horizontal, "horizontal alignment");
    touched = true;
  }
  if (input.vertical !== undefined) {
    next.vertical = ordinal(VERTICAL_ALIGN, input.vertical, "vertical alignment");
    touched = true;
  }
  if (input.wrapText !== undefined) {
    next.wrapText = input.wrapText;
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
    next.applyAlignment = true;
  }
  return next;
}

/** The style properties a base `xf` contributes once the deltas are applied. */
type DerivedStyle = {
  xf: CellXf;
  fontIndex: number;
  fillIndex: number;
  numFmtId: number;
  border: Border;
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
  const { ooxml } = wb;
  const baseXf = ooxml.getCellXf(baseXfIndex);

  const fontIndex = style.font
    ? ooxml.addFont(applyFont(ooxml.getFont(baseXf.fontIndex), style.font))
    : baseXf.fontIndex;
  const fillIndex = style.fill
    ? ooxml.addFill(applyFill(ooxml.getFill(baseXf.fillIndex), style.fill))
    : baseXf.fillIndex;

  let numFmtId = baseXf.numFmtId;
  if (style.numberFormat !== undefined) {
    numFmtId = style.numberFormat === "" ? 0 : ooxml.addNumberFormat(style.numberFormat);
  }

  const derived: DerivedStyle = {
    xf: baseXf,
    fontIndex,
    fillIndex,
    numFmtId,
    border: ooxml.getBorder(baseXf.borderIndex),
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
  const { ooxml } = wb;
  const sheetOoxml = sheetAt(wb, sheet).ooxml;
  for (const rowBand of rowBands) {
    for (const colBand of colBands) {
      const band = {
        firstRow: rowBand.first,
        firstCol: colBand.first,
        lastRow: rowBand.last,
        lastCol: colBand.last,
      };
      const patched = new Map<number, number>();
      const mapping = new Map<number, number>();
      for (const xfIndex of sheetOoxml.getRangeXfBases(band)) {
        // "default" discards whatever is there, so every base patches xf 0.
        const baseXfIndex = base === "default" ? 0 : xfIndex;
        let patchedIndex = patched.get(baseXfIndex);
        if (patchedIndex === undefined) {
          const derived = deriveStyle(wb, baseXfIndex, style, cache);
          const borderIndex = borderInput
            ? ooxml.addBorder(buildBorder(derived.border, borderInput, rowBand, colBand))
            : derived.xf.borderIndex;
          patchedIndex = ooxml.addCellXf(
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
          patched.set(baseXfIndex, patchedIndex);
          regions.push({ ...band, baseXfIndex, xfIndex: patchedIndex });
        }
        mapping.set(xfIndex, patchedIndex);
      }
      sheetOoxml.remapRangeXfs(band, mapping);
    }
  }

  return { regions };
}
