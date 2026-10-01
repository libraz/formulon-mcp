import type { PrintSettingsInput } from "../print.js";
import type { StyleBase, StyleInput } from "../styles.js";

export type ColumnLayoutOperation = {
  type: "column";
  first: number;
  last?: number;
  size: number;
  unit?: "chars" | "pt" | "px" | "mm";
};

export type RowLayoutOperation = {
  type: "row";
  first: number;
  last?: number;
  size: number;
  unit?: "pt" | "px" | "mm";
};

export type MergeLayoutOperation = {
  type: "merge";
  operation: "add" | "remove";
  range: string;
};

export type StyleLayoutOperation = {
  type: "style";
  range: string;
  style: StyleInput;
  base?: StyleBase;
};

export type PrintLayoutOperation = {
  type: "print";
  settings: PrintSettingsInput;
};

export type LayoutOperation =
  | ColumnLayoutOperation
  | RowLayoutOperation
  | MergeLayoutOperation
  | StyleLayoutOperation
  | PrintLayoutOperation;

export type LayoutRect = {
  firstRow: number;
  firstCol: number;
  lastRow: number;
  lastCol: number;
};

export type ResolvedLayoutRange = {
  ref: string;
  rect: LayoutRect;
};

export type PreparedLayoutOperation =
  | (ColumnLayoutOperation & { first: number; last: number; storedSize: number })
  | (RowLayoutOperation & { first: number; last: number; storedSize: number })
  | (MergeLayoutOperation & { resolved: ResolvedLayoutRange })
  | (StyleLayoutOperation & { resolved: ResolvedLayoutRange; base: StyleBase })
  | PrintLayoutOperation;

export const MAX_LAYOUT_OPERATIONS = 200;
export const MAX_LAYOUT_COLUMN_WIDTH_CHARS = 255;
export const MAX_LAYOUT_ROW_HEIGHT_PT = 409;
export const MAX_EXPANDED_ROW_WRITES = 10_000;
export const MAX_STYLED_CELLS = 100_000;
