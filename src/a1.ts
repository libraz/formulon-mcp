export type CellAddress = {
  sheetName?: string;
  row: number;
  col: number;
};

export type RangeAddress = {
  sheetName?: string;
  start: CellAddress;
  end: CellAddress;
};

const CELL_PATTERN =
  /^(?:(?<sheet>'(?:[^']|'')+'|[^!]+)!)?\$?(?<col>[A-Za-z]+)\$?(?<row>[1-9][0-9]*)$/;

function normalizeSheetName(sheet: string | undefined): string | undefined {
  if (!sheet) {
    return undefined;
  }
  if (sheet.startsWith("'") && sheet.endsWith("'")) {
    return sheet.slice(1, -1).replaceAll("''", "'");
  }
  return sheet;
}

/** Converts an Excel column label such as `A` or `AA` to a zero-based index. */
export function colToIndex(col: string): number {
  if (col.length === 0) {
    throw new Error("invalid column: empty");
  }
  let index = 0;
  for (const char of col.toUpperCase()) {
    const code = char.charCodeAt(0);
    if (code < 65 || code > 90) {
      throw new Error(`invalid column: ${col}`);
    }
    index = index * 26 + (code - 64);
  }
  return index - 1;
}

/** Converts a zero-based column index to an Excel column label. */
export function indexToCol(index: number): string {
  if (!Number.isInteger(index) || index < 0) {
    throw new Error(`invalid zero-based column index: ${index}`);
  }
  let n = index + 1;
  let col = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

/** Formats a zero-based row and column pair as an A1 reference. */
export function cellToA1(row: number, col: number): string {
  if (!Number.isInteger(row) || row < 0) {
    throw new Error(`invalid zero-based row index: ${row}`);
  }
  return `${indexToCol(col)}${row + 1}`;
}

/** Parses an A1 cell reference with an optional sheet name. */
export function parseCellRef(ref: string): CellAddress {
  const match = CELL_PATTERN.exec(ref.trim());
  if (!match?.groups) {
    throw new Error(`invalid A1 cell reference: ${ref}`);
  }
  return {
    sheetName: normalizeSheetName(match.groups.sheet),
    row: Number.parseInt(match.groups.row, 10) - 1,
    col: colToIndex(match.groups.col),
  };
}

/** Parses an A1 range reference with an optional sheet name. */
export function parseRangeRef(ref: string): RangeAddress {
  const parts = ref.split(":");
  if (parts.length > 2) {
    throw new Error(`invalid A1 range reference: ${ref}`);
  }
  const [left, right] = parts;
  if (!(left && right)) {
    const cell = parseCellRef(ref);
    return { sheetName: cell.sheetName, start: cell, end: cell };
  }
  const start = parseCellRef(left);
  const end = parseCellRef(
    right.includes("!") || !start.sheetName ? right : `${start.sheetName}!${right}`,
  );
  const sheetName = start.sheetName ?? end.sheetName;
  if (start.sheetName && end.sheetName && start.sheetName !== end.sheetName) {
    throw new Error(`cross-sheet ranges are not supported: ${ref}`);
  }
  return {
    sheetName,
    start: {
      ...start,
      sheetName,
    },
    end: {
      ...end,
      sheetName,
    },
  };
}

/** Normalizes a range so `start` is the top-left cell and `end` is the bottom-right cell. */
export function normalizeRange(range: RangeAddress): RangeAddress {
  return {
    sheetName: range.sheetName,
    start: {
      sheetName: range.sheetName,
      row: Math.min(range.start.row, range.end.row),
      col: Math.min(range.start.col, range.end.col),
    },
    end: {
      sheetName: range.sheetName,
      row: Math.max(range.start.row, range.end.row),
      col: Math.max(range.start.col, range.end.col),
    },
  };
}
