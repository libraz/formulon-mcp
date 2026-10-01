import type { ColumnMetrics } from "./geometry.js";

export type PreviewOptions = {
  sheet?: number | string;
  range?: string;
  scale?: number;
  showGridLines?: boolean;
  showPageBreaks?: boolean;
  recalc?: boolean;
  outputPath?: string;
};

export type PreviewBorder = {
  style: number;
  color: string;
};

export type PreviewStyle = {
  fill: string | undefined;
  fontFamily: string;
  fontSize: number;
  fontColor: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  horizontal: number;
  vertical: number;
  wrapText: boolean;
  indent: number;
  borders: {
    left: PreviewBorder;
    right: PreviewBorder;
    top: PreviewBorder;
    bottom: PreviewBorder;
  };
};

export type PreviewCell = {
  row: number;
  col: number;
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  formula: string;
  numeric: boolean;
  style: PreviewStyle;
  merge?: { firstRow: number; firstCol: number; lastRow: number; lastCol: number };
};

export type PreviewCellMetadata = {
  row: number;
  col: number;
  a1: string;
  text: string;
  formula?: string;
  x: number;
  y: number;
  width: number;
  height: number;
};

export type PreviewModel = {
  firstRow: number;
  firstCol: number;
  lastRow: number;
  lastCol: number;
  logicalWidth: number;
  logicalHeight: number;
  logicalUnit: "pt";
  columns: number[];
  rows: number[];
  columnMetrics: ColumnMetrics;
  cells: PreviewCell[];
  rowBreaks: number[];
  colBreaks: number[];
  showGridLines: boolean;
  showPageBreaks: boolean;
};

export type PreviewMetadata = {
  sheet: number;
  sheetName: string;
  range: string;
  logicalWidth: number;
  logicalHeight: number;
  logicalUnit: "pt";
  scaledWidth: number;
  scaledHeight: number;
  scale: number;
  fontList: string[];
  pageCount: number;
  rowBreaks: number[];
  colBreaks: number[];
  warnings: string[];
  cells: PreviewCellMetadata[];
  cellGeometryUnit: "pt";
  metadataTruncated: boolean;
  omittedCellCount: number;
  warningsTruncated: boolean;
};

export type PreviewResult = {
  svg: string;
  png: Buffer;
  metadata: PreviewMetadata;
  outputPath?: string;
};
