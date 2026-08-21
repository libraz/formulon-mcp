import { cellToA1 } from "./a1.js";
import { assertStatus, normalizeFormula, type Workbook } from "./formulon.js";
import { classifyNumberFormat, isoToExcelSerial } from "./numfmt.js";
import { PRINT_PRESETS, type PrintSettingsInput, writePrintSettings } from "./print.js";
import { applyRangeStyle, type StyleInput } from "./styles.js";

/**
 * Block-composed document authoring.
 *
 * The primitives (`set_range`, `style_range`, `print_settings`) each do one
 * thing well, but composing them into a document leaves the caller owning all
 * the row arithmetic: where the total row lands, which range `SUM` covers,
 * which cells the borders go around. Get one offset wrong and the formula is
 * silently off by a row, and nothing surfaces it until someone reads the file.
 *
 * This module takes the stack of blocks a document actually is — a title, some
 * header fields, a line-item table, a summary — and resolves positions itself.
 * References between blocks are written by name (`{table.Amount}`,
 * `{Subtotal}`) and bound to A1 only once the layout is known, so a range
 * cannot drift from what it was meant to cover.
 *
 * It is deliberately blind to document semantics: nothing here knows what an
 * invoice is, what tax is, or how a total should be computed. Those are the
 * caller's labels and the caller's formulas.
 */

/** Short names for the number formats a document reaches for constantly. */
const FORMAT_ALIASES: Record<string, string> = Object.freeze({
  date: "yyyy/mm/dd",
  datetime: "yyyy/mm/dd hh:mm",
  time: "hh:mm",
  number: "#,##0",
  decimal: "#,##0.00",
  percent: "0.0%",
});

/** A literal a block can write into a cell. */
export type CellLiteral = string | number | boolean | null;

export type Alignment = "left" | "center" | "right";

type BlockBase = {
  /** Qualifies the names this block registers, so two blocks can share labels. */
  name?: string;
  /** Start at the same row as the previous block instead of below it. */
  sameRow?: boolean;
};

export type TitleBlock = BlockBase & {
  type: "title";
  text: string;
  align?: Alignment;
  size?: number;
  bold?: boolean;
};

export type TextBlock = BlockBase & {
  type: "text";
  text: string;
  align?: Alignment;
  size?: number;
  bold?: boolean;
  /** Merge across this many columns; defaults to the document width. */
  span?: number;
  wrap?: boolean;
  /** Rows the text occupies; useful with `wrap` for a notes paragraph. */
  rows?: number;
};

export type FieldItem = {
  label: string;
  value?: CellLiteral;
  formula?: string;
  format?: string;
  name?: string;
};

export type FieldsBlock = BlockBase & {
  type: "fields";
  items: FieldItem[];
  align?: "left" | "right";
  labelSpan?: number;
  valueSpan?: number;
  /** Rule the value cell underneath, the way a paper form does. */
  rule?: boolean;
};

export type TableColumn = {
  header: string;
  /** Name this column is referenced by in another column's formula. Defaults to the header. */
  key?: string;
  /** Per-row formula; `{key}` binds to that column's cell in the same row. */
  formula?: string;
  width?: number;
  format?: string;
  align?: Alignment;
};

export type TableRow = CellLiteral[] | Record<string, CellLiteral>;

export type TableBlock = BlockBase & {
  type: "table";
  columns: TableColumn[];
  rows?: TableRow[];
  /** Blank ruled rows to lay out when `rows` is omitted — a form to fill in. */
  rowCount?: number;
  /** Fill applied to every other body row. */
  bandColor?: string;
};

export type SummaryItem = {
  label: string;
  value?: CellLiteral;
  formula?: string;
  format?: string;
  /** Bold the row and box it, for the line that carries the final figure. */
  emphasis?: boolean;
  name?: string;
};

export type SummaryBlock = BlockBase & {
  type: "summary";
  items: SummaryItem[];
  align?: "left" | "right";
  labelSpan?: number;
  valueSpan?: number;
  /** Rule each label/value pair. Defaults to true. */
  border?: boolean;
};

export type SpacerBlock = BlockBase & {
  type: "spacer";
  rows?: number;
};

export type DocumentBlock =
  | TitleBlock
  | TextBlock
  | FieldsBlock
  | TableBlock
  | SummaryBlock
  | SpacerBlock;

export type DocumentTheme = {
  /** Base font applied to every cell the document writes. */
  font?: string;
  size?: number;
  /** Header-band fill. Header text switches to `headerText` when this is set. */
  accent?: string;
  headerText?: string;
  /** Border style used for a table grid and a summary rule. Defaults to `thin`. */
  border?: string;
  /** Border style around a table and around an emphasized summary row. Defaults to `medium`. */
  outline?: string;
};

export type DocumentSpec = {
  blocks: DocumentBlock[];
  /** Top-left anchor. Defaults to `B2`, leaving a margin row and column. */
  start?: string;
  /** Document width in columns. Defaults to the widest table's column count. */
  width?: number;
  theme?: DocumentTheme;
  /** Named page setup; also sets the print area to the document's extent. */
  print?: string;
  /** Repeat the first table's header row on every printed page. Defaults to true. */
  repeatTableHeader?: boolean;
};

type Rect = { firstRow: number; firstCol: number; lastRow: number; lastCol: number };

type ResolvedTheme = DocumentTheme & { border: string; outline: string };

/** Writing context threaded through every block writer. */
type Layout = {
  wb: Workbook;
  sheet: number;
  firstCol: number;
  lastCol: number;
  theme: ResolvedTheme;
  names: NameTable;
  /** Set by the first table block, so print settings can repeat its header. */
  headerRow?: number;
};

/** What a block occupied: its height in rows and the columns it spans. */
type Placed = { height: number; range: string; firstCol: number; lastCol: number };

/** Resolves a format alias to an Excel format code, passing a literal code through. */
function formatCode(format: string): string {
  return FORMAT_ALIASES[format] ?? format;
}

/**
 * Converts a value for a cell that carries `format`. An ISO date string under a
 * date format becomes the serial Excel stores, so the cell holds a real date
 * rather than text that merely looks like one.
 */
function coerceValue(value: CellLiteral, format?: string): CellLiteral {
  if (typeof value !== "string" || format === undefined) {
    return value;
  }
  const kind = classifyNumberFormat(164, formatCode(format));
  if (kind !== "date" && kind !== "datetime" && kind !== "time") {
    return value;
  }
  return isoToExcelSerial(value) ?? value;
}

/** Tracks where every named cell or range landed, for formulas and for the caller. */
class NameTable {
  private readonly entries = new Map<string, string>();

  set(name: string, reference: string): void {
    this.entries.set(name, reference);
  }

  get(name: string): string | undefined {
    return this.entries.get(name);
  }

  known(): string[] {
    return Array.from(this.entries.keys());
  }

  toJson(): Record<string, string> {
    return Object.fromEntries(this.entries);
  }
}

/**
 * Binds `{name}` references in a formula to the A1 ranges the layout produced.
 *
 * An unknown reference is an error rather than a silently broken formula: a
 * `SUM` over the wrong range is exactly the failure this module exists to
 * prevent. Braces holding a comma, semicolon or quote are an Excel array
 * constant such as `{1,2;3,4}` and pass through untouched.
 */
function bindFormula(formula: string, names: NameTable, rowScope?: Map<string, string>): string {
  return formula.replace(/\{([^{}]*)\}/g, (whole, token: string) => {
    const key = token.trim();
    const bound = rowScope?.get(key) ?? names.get(key);
    if (bound !== undefined) {
      return bound;
    }
    if (/[,;"]/.test(key)) {
      return whole;
    }
    const known = [...(rowScope ? Array.from(rowScope.keys()) : []), ...names.known()];
    throw new Error(
      `unknown reference {${key}} in formula "${formula}"; known names: ${known.join(", ") || "(none)"}`,
    );
  });
}

function ref(row: number, col: number): string {
  return cellToA1(row, col);
}

function rectRef(rect: Rect): string {
  return `${ref(rect.firstRow, rect.firstCol)}:${ref(rect.lastRow, rect.lastCol)}`;
}

function row(rowIndex: number, firstCol: number, lastCol: number): Rect {
  return { firstRow: rowIndex, firstCol, lastRow: rowIndex, lastCol };
}

function writeValue(layout: Layout, rowIndex: number, col: number, value: CellLiteral): void {
  if (value === null) {
    return;
  }
  const { wb, sheet } = layout;
  const where = `write ${ref(rowIndex, col)}`;
  if (typeof value === "number") {
    assertStatus(wb.setNumber(sheet, rowIndex, col, value), where);
  } else if (typeof value === "boolean") {
    assertStatus(wb.setBool(sheet, rowIndex, col, value), where);
  } else {
    assertStatus(wb.setText(sheet, rowIndex, col, value), where);
  }
}

function writeFormula(layout: Layout, rowIndex: number, col: number, formula: string): void {
  assertStatus(
    layout.wb.setFormula(layout.sheet, rowIndex, col, normalizeFormula(formula)),
    `write formula at ${ref(rowIndex, col)}`,
  );
}

function mergeRect(layout: Layout, rect: Rect): void {
  if (rect.lastCol <= rect.firstCol && rect.lastRow <= rect.firstRow) {
    return;
  }
  assertStatus(layout.wb.addMerge(layout.sheet, rect), `merge ${rectRef(rect)}`);
}

function style(layout: Layout, rect: Rect, input: StyleInput): void {
  applyRangeStyle(layout.wb, layout.sheet, rect, input, "existing");
}

function baseFont(layout: Layout): StyleInput["font"] {
  const { font, size } = layout.theme;
  return {
    ...(font === undefined ? {} : { name: font }),
    ...(size === undefined ? {} : { size }),
  };
}

function placeTitle(layout: Layout, block: TitleBlock, top: number): Placed {
  const rect = row(top, layout.firstCol, layout.lastCol);
  writeValue(layout, top, layout.firstCol, block.text);
  mergeRect(layout, rect);
  style(layout, rect, {
    font: { ...baseFont(layout), size: block.size ?? 18, bold: block.bold ?? true },
    align: { horizontal: block.align ?? "center", vertical: "center" },
  });
  layout.names.set(block.name ?? "title", rectRef(rect));
  return { height: 1, range: rectRef(rect), firstCol: rect.firstCol, lastCol: rect.lastCol };
}

function placeText(layout: Layout, block: TextBlock, top: number): Placed {
  const height = Math.max(1, block.rows ?? 1);
  const lastCol =
    block.span === undefined
      ? layout.lastCol
      : Math.min(layout.lastCol, layout.firstCol + block.span - 1);
  const rect: Rect = {
    firstRow: top,
    firstCol: layout.firstCol,
    lastRow: top + height - 1,
    lastCol,
  };
  writeValue(layout, top, layout.firstCol, block.text);
  mergeRect(layout, rect);
  style(layout, rect, {
    font: {
      ...baseFont(layout),
      ...(block.size === undefined ? {} : { size: block.size }),
      bold: block.bold ?? false,
    },
    align: {
      horizontal: block.align ?? "left",
      vertical: block.wrap ? "top" : "center",
      ...(block.wrap ? { wrapText: true } : {}),
    },
  });
  if (block.name) {
    layout.names.set(block.name, rectRef(rect));
  }
  return { height, range: rectRef(rect), firstCol: rect.firstCol, lastCol };
}

/** Resolves the label and value column spans a fields or summary block occupies. */
function pairColumns(
  layout: Layout,
  align: "left" | "right",
  labelSpan: number,
  valueSpan: number,
): { labelFirst: number; labelLast: number; valueFirst: number; valueLast: number } {
  if (align === "right") {
    const valueLast = layout.lastCol;
    const valueFirst = valueLast - valueSpan + 1;
    const labelLast = valueFirst - 1;
    return { labelFirst: labelLast - labelSpan + 1, labelLast, valueFirst, valueLast };
  }
  const labelFirst = layout.firstCol;
  const labelLast = labelFirst + labelSpan - 1;
  const valueFirst = labelLast + 1;
  return { labelFirst, labelLast, valueFirst, valueLast: valueFirst + valueSpan - 1 };
}

function assertFits(layout: Layout, cols: ReturnType<typeof pairColumns>, what: string): void {
  if (cols.labelFirst < layout.firstCol || cols.valueLast > layout.lastCol) {
    throw new Error(
      `${what} needs ${cols.valueLast - cols.labelFirst + 1} columns but the document is ${
        layout.lastCol - layout.firstCol + 1
      } wide`,
    );
  }
}

function placeFields(layout: Layout, block: FieldsBlock, top: number): Placed {
  const cols = pairColumns(
    layout,
    block.align ?? "left",
    block.labelSpan ?? 1,
    block.valueSpan ?? 1,
  );
  assertFits(layout, cols, "fields block");
  block.items.forEach((item, index) => {
    const rowIndex = top + index;
    writeValue(layout, rowIndex, cols.labelFirst, item.label);
    mergeRect(layout, row(rowIndex, cols.labelFirst, cols.labelLast));
    mergeRect(layout, row(rowIndex, cols.valueFirst, cols.valueLast));
    if (item.formula !== undefined) {
      writeFormula(layout, rowIndex, cols.valueFirst, bindFormula(item.formula, layout.names));
    } else if (item.value !== undefined) {
      writeValue(layout, rowIndex, cols.valueFirst, coerceValue(item.value, item.format));
    }
    style(layout, row(rowIndex, cols.labelFirst, cols.labelLast), { font: baseFont(layout) });
    style(layout, row(rowIndex, cols.valueFirst, cols.valueLast), {
      font: baseFont(layout),
      ...(item.format === undefined ? {} : { numberFormat: formatCode(item.format) }),
      ...(block.rule ? { border: { bottom: layout.theme.border } } : {}),
    });
    const key = item.name ?? item.label;
    const valueRef = ref(rowIndex, cols.valueFirst);
    layout.names.set(key, valueRef);
    if (block.name) {
      layout.names.set(`${block.name}.${key}`, valueRef);
    }
  });
  const rect: Rect = {
    firstRow: top,
    firstCol: cols.labelFirst,
    lastRow: top + block.items.length - 1,
    lastCol: cols.valueLast,
  };
  if (block.name) {
    layout.names.set(block.name, rectRef(rect));
  }
  return {
    height: block.items.length,
    range: rectRef(rect),
    firstCol: cols.labelFirst,
    lastCol: cols.valueLast,
  };
}

function tableRowValue(source: TableRow, column: TableColumn, index: number): CellLiteral {
  if (Array.isArray(source)) {
    return source[index] ?? null;
  }
  return source[column.key ?? column.header] ?? source[column.header] ?? null;
}

function placeTable(layout: Layout, block: TableBlock, top: number): Placed {
  const width = layout.lastCol - layout.firstCol + 1;
  if (block.columns.length === 0) {
    throw new Error("table block needs at least one column");
  }
  if (block.columns.length > width) {
    throw new Error(
      `table needs ${block.columns.length} columns but the document is ${width} wide`,
    );
  }
  const bodyRows = block.rows?.length ?? block.rowCount ?? 0;
  const headerRow = top;
  const bodyTop = top + 1;
  const bodyBottom = bodyTop + bodyRows - 1;
  const lastCol = layout.firstCol + block.columns.length - 1;
  const prefix = block.name ?? "table";

  block.columns.forEach((column, index) => {
    writeValue(layout, headerRow, layout.firstCol + index, column.header);
  });

  for (let offset = 0; offset < bodyRows; offset += 1) {
    const rowIndex = bodyTop + offset;
    // Every column's address is known before any formula is bound, so a
    // per-row formula can name a column written later in the same row.
    const rowScope = new Map<string, string>();
    block.columns.forEach((column, index) => {
      const address = ref(rowIndex, layout.firstCol + index);
      rowScope.set(column.header, address);
      if (column.key) {
        rowScope.set(column.key, address);
      }
    });
    block.columns.forEach((column, index) => {
      const source = block.rows?.[offset];
      const value = source === undefined ? null : tableRowValue(source, column, index);
      if (column.formula !== undefined && value === null) {
        writeFormula(
          layout,
          rowIndex,
          layout.firstCol + index,
          bindFormula(column.formula, layout.names, rowScope),
        );
      } else {
        writeValue(layout, rowIndex, layout.firstCol + index, coerceValue(value, column.format));
      }
    });
  }

  const rect: Rect = {
    firstRow: headerRow,
    firstCol: layout.firstCol,
    lastRow: bodyRows > 0 ? bodyBottom : headerRow,
    lastCol,
  };
  const headerRect = row(headerRow, layout.firstCol, lastCol);

  // Rule the whole block first, then the header band, then the per-column
  // formats: each pass is a delta, so the later ones keep what the first laid
  // down rather than replacing it.
  style(layout, rect, {
    font: baseFont(layout),
    border: { all: layout.theme.border, outline: layout.theme.outline },
  });
  style(layout, headerRect, {
    font: {
      ...baseFont(layout),
      bold: true,
      ...(layout.theme.accent === undefined ? {} : { color: layout.theme.headerText ?? "#FFFFFF" }),
    },
    ...(layout.theme.accent === undefined ? {} : { fill: { color: layout.theme.accent } }),
    align: { horizontal: "center", vertical: "center" },
    border: { all: layout.theme.border, outline: layout.theme.outline },
  });

  block.columns.forEach((column, index) => {
    const col = layout.firstCol + index;
    if (column.width !== undefined) {
      assertStatus(
        layout.wb.setColumnWidth(layout.sheet, col, col, column.width),
        `set width of column ${column.header}`,
      );
    }
    if (bodyRows === 0) {
      return;
    }
    const columnRect: Rect = {
      firstRow: bodyTop,
      firstCol: col,
      lastRow: bodyBottom,
      lastCol: col,
    };
    if (column.format !== undefined || column.align !== undefined) {
      style(layout, columnRect, {
        ...(column.format === undefined ? {} : { numberFormat: formatCode(column.format) }),
        ...(column.align === undefined ? {} : { align: { horizontal: column.align } }),
      });
    }
    layout.names.set(`${prefix}.${column.header}`, rectRef(columnRect));
    if (column.key) {
      layout.names.set(`${prefix}.${column.key}`, rectRef(columnRect));
    }
  });

  if (block.bandColor !== undefined) {
    for (let offset = 1; offset < bodyRows; offset += 2) {
      style(layout, row(bodyTop + offset, layout.firstCol, lastCol), {
        fill: { color: block.bandColor },
      });
    }
  }

  layout.names.set(prefix, rectRef(rect));
  layout.names.set(`${prefix}.header`, rectRef(headerRect));
  if (bodyRows > 0) {
    layout.names.set(
      `${prefix}.body`,
      rectRef({ firstRow: bodyTop, firstCol: layout.firstCol, lastRow: bodyBottom, lastCol }),
    );
  }
  layout.headerRow ??= headerRow;
  return { height: 1 + bodyRows, range: rectRef(rect), firstCol: layout.firstCol, lastCol };
}

function placeSummary(layout: Layout, block: SummaryBlock, top: number): Placed {
  const width = layout.lastCol - layout.firstCol + 1;
  const cols = pairColumns(
    layout,
    block.align ?? "right",
    block.labelSpan ?? Math.max(1, Math.min(2, width - 1)),
    block.valueSpan ?? 1,
  );
  assertFits(layout, cols, "summary block");
  const ruled = block.border ?? true;
  block.items.forEach((item, index) => {
    const rowIndex = top + index;
    writeValue(layout, rowIndex, cols.labelFirst, item.label);
    mergeRect(layout, row(rowIndex, cols.labelFirst, cols.labelLast));
    mergeRect(layout, row(rowIndex, cols.valueFirst, cols.valueLast));
    if (item.formula !== undefined) {
      writeFormula(layout, rowIndex, cols.valueFirst, bindFormula(item.formula, layout.names));
    } else if (item.value !== undefined) {
      writeValue(layout, rowIndex, cols.valueFirst, coerceValue(item.value, item.format));
    }
    style(layout, row(rowIndex, cols.labelFirst, cols.valueLast), {
      font: { ...baseFont(layout), bold: item.emphasis ?? false },
      ...(ruled
        ? {
            border: {
              all: layout.theme.border,
              ...(item.emphasis ? { outline: layout.theme.outline } : {}),
            },
          }
        : {}),
    });
    style(layout, row(rowIndex, cols.labelFirst, cols.labelLast), {
      align: { horizontal: "right" },
    });
    if (item.format !== undefined) {
      style(layout, row(rowIndex, cols.valueFirst, cols.valueLast), {
        numberFormat: formatCode(item.format),
      });
    }
    // Registered after the row is written, so a later item can reference it and
    // an earlier one cannot — the same direction Excel's dependency chain runs.
    const key = item.name ?? item.label;
    const valueRef = ref(rowIndex, cols.valueFirst);
    layout.names.set(key, valueRef);
    if (block.name) {
      layout.names.set(`${block.name}.${key}`, valueRef);
    }
  });
  const rect: Rect = {
    firstRow: top,
    firstCol: cols.labelFirst,
    lastRow: top + block.items.length - 1,
    lastCol: cols.valueLast,
  };
  if (block.name) {
    layout.names.set(block.name, rectRef(rect));
  }
  return {
    height: block.items.length,
    range: rectRef(rect),
    firstCol: cols.labelFirst,
    lastCol: cols.valueLast,
  };
}

/** Derives the document's column count from the widest table it contains. */
function documentWidth(spec: DocumentSpec): number {
  if (spec.width !== undefined) {
    if (spec.width < 1) {
      throw new Error("width must be at least 1 column");
    }
    return spec.width;
  }
  const widest = spec.blocks.reduce(
    (max, block) => (block.type === "table" ? Math.max(max, block.columns.length) : max),
    0,
  );
  return widest > 0 ? widest : 2;
}

export type PlacedBlock = {
  index: number;
  type: DocumentBlock["type"];
  name?: string;
  range: string;
};

export type BuiltDocument = {
  start: string;
  range: string;
  width: number;
  blocks: PlacedBlock[];
  names: Record<string, string>;
  pageCount?: number;
};

/**
 * Lays out and writes a document, returning where every block and named cell
 * landed so the caller can refine any of it with the primitive tools.
 */
export function buildDocument(
  wb: Workbook,
  sheet: number,
  anchor: { row: number; col: number },
  spec: DocumentSpec,
): BuiltDocument {
  if (spec.blocks.length === 0) {
    throw new Error("a document needs at least one block");
  }
  const width = documentWidth(spec);
  const layout: Layout = {
    wb,
    sheet,
    firstCol: anchor.col,
    lastCol: anchor.col + width - 1,
    theme: { border: "thin", outline: "medium", ...spec.theme },
    names: new NameTable(),
  };

  const placed: PlacedBlock[] = [];
  let top = anchor.row;
  let groupHeight = 0;
  // Columns already taken by blocks sharing the current row group. Two blocks
  // on one row that overlap would merge over each other, which the writer
  // cannot express and Excel would offer to repair.
  let groupSpans: { index: number; firstCol: number; lastCol: number }[] = [];
  spec.blocks.forEach((block, index) => {
    if (!block.sameRow) {
      top += groupHeight;
      groupHeight = 0;
      groupSpans = [];
    }
    if (block.type === "spacer") {
      const height = Math.max(0, block.rows ?? 1);
      placed.push({
        index,
        type: "spacer",
        range:
          height > 0
            ? rectRef({
                firstRow: top,
                firstCol: layout.firstCol,
                lastRow: top + height - 1,
                lastCol: layout.lastCol,
              })
            : "",
      });
      groupHeight = Math.max(groupHeight, height);
      return;
    }
    const result =
      block.type === "title"
        ? placeTitle(layout, block, top)
        : block.type === "text"
          ? placeText(layout, block, top)
          : block.type === "fields"
            ? placeFields(layout, block, top)
            : block.type === "table"
              ? placeTable(layout, block, top)
              : placeSummary(layout, block, top);
    const clash = groupSpans.find(
      (span) => result.firstCol <= span.lastCol && result.lastCol >= span.firstCol,
    );
    if (clash) {
      throw new Error(
        `block ${index} (${block.type}) shares a row with block ${clash.index} and overlaps its columns; give one of them a narrower span or drop sameRow`,
      );
    }
    groupSpans.push({ index, firstCol: result.firstCol, lastCol: result.lastCol });
    placed.push({ index, type: block.type, name: block.name, range: result.range });
    groupHeight = Math.max(groupHeight, result.height);
  });

  const lastRow = Math.max(anchor.row, top + groupHeight - 1);
  const range = rectRef({
    firstRow: anchor.row,
    firstCol: layout.firstCol,
    lastRow,
    lastCol: layout.lastCol,
  });

  let pageCount: number | undefined;
  if (spec.print !== undefined) {
    const preset = PRINT_PRESETS[spec.print];
    if (!preset) {
      throw new Error(
        `unknown print preset: ${spec.print} (expected one of ${Object.keys(PRINT_PRESETS).join(", ")})`,
      );
    }
    const settings: PrintSettingsInput = { ...preset, printArea: range };
    if (layout.headerRow !== undefined && spec.repeatTableHeader !== false) {
      const line = layout.headerRow + 1;
      settings.printTitles = { repeatRows: `${line}:${line}` };
    }
    writePrintSettings(wb, sheet, settings);
    const pagination = wb.paginate(sheet);
    assertStatus(pagination.status, "paginate document");
    pageCount = pagination.pageCount;
  }

  assertStatus(wb.recalc(), "recalc document");
  return {
    start: ref(anchor.row, anchor.col),
    range,
    width,
    blocks: placed,
    names: layout.names.toJson(),
    pageCount,
  };
}
