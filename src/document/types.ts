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
