/**
 * Formula fingerprints used by the formula audit.
 *
 * This is intentionally a small lexical scanner rather than a formula parser.
 * It only rewrites unambiguous A1 references, preserving literals, quoted
 * sheet names, names, and structured references byte-for-byte. A fingerprint
 * therefore remains useful when the engine accepts syntax this package does
 * not understand yet.
 */

const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

type ScanResult = {
  fingerprint: string;
  supported: boolean;
};

type CellReference = {
  end: number;
  text: string;
  absoluteRow: boolean;
  absoluteCol: boolean;
  row: number;
  col: number;
};

function isIdentifierChar(char: string | undefined): boolean {
  return char !== undefined && /[\p{L}\p{N}\p{M}_.]/u.test(char);
}

function columnIndex(column: string): number {
  let value = 0;
  for (const char of column.toUpperCase()) {
    value = value * 26 + char.charCodeAt(0) - 64;
  }
  return value - 1;
}

function parseCellReference(text: string, start: number): CellReference | undefined {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)/.exec(text.slice(start));
  if (!match) {
    return undefined;
  }
  const full = match[0];
  const end = start + full.length;
  if (isIdentifierChar(text[end])) {
    return undefined;
  }
  const colText = match[1];
  const row = Number.parseInt(match[2], 10);
  const col = columnIndex(colText);
  if (!Number.isInteger(row) || row < 1 || row > MAX_ROW || col < 0 || col >= MAX_COL) {
    return undefined;
  }
  return {
    end,
    text: full,
    absoluteRow: full.includes(`$${match[2]}`),
    absoluteCol: full.startsWith("$"),
    row: row - 1,
    col,
  };
}

function parseWholeColumn(
  text: string,
  start: number,
):
  | {
      end: number;
      left: { absolute: boolean; col: number };
      right: { absolute: boolean; col: number };
    }
  | undefined {
  const match = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_.])/.exec(
    text.slice(start),
  );
  if (!match || isIdentifierChar(text[start - 1])) {
    return undefined;
  }
  const left = columnIndex(match[2]);
  const right = columnIndex(match[4]);
  if (left < 0 || left >= MAX_COL || right < 0 || right >= MAX_COL) {
    return undefined;
  }
  return {
    end: start + match[0].length,
    left: { absolute: match[1] === "$", col: left },
    right: { absolute: match[3] === "$", col: right },
  };
}

function parseWholeRow(
  text: string,
  start: number,
):
  | {
      end: number;
      left: { absolute: boolean; row: number };
      right: { absolute: boolean; row: number };
    }
  | undefined {
  const match = /^(\$?)(\d+):(\$?)(\d+)(?![A-Za-z0-9_.])/.exec(text.slice(start));
  if (!match || isIdentifierChar(text[start - 1])) {
    return undefined;
  }
  const left = Number.parseInt(match[2], 10) - 1;
  const right = Number.parseInt(match[4], 10) - 1;
  if (left < 0 || left >= MAX_ROW || right < 0 || right >= MAX_ROW) {
    return undefined;
  }
  return {
    end: start + match[0].length,
    left: { absolute: match[1] === "$", row: left },
    right: { absolute: match[3] === "$", row: right },
  };
}

function relativeAxis(value: number, origin: number, absolute: boolean, prefix: "R" | "C"): string {
  return absolute ? `${prefix}${value + 1}` : `${prefix}[${value - origin}]`;
}

function normalizeCell(reference: CellReference, originRow: number, originCol: number): string {
  return `${relativeAxis(reference.row, originRow, reference.absoluteRow, "R")}${relativeAxis(
    reference.col,
    originCol,
    reference.absoluteCol,
    "C",
  )}`;
}

function normalizeColumn(axis: { absolute: boolean; col: number }, originCol: number): string {
  return relativeAxis(axis.col, originCol, axis.absolute, "C");
}

function normalizeRow(axis: { absolute: boolean; row: number }, originRow: number): string {
  return relativeAxis(axis.row, originRow, axis.absolute, "R");
}

function quotedStringEnd(text: string, start: number, quote: '"' | "'"): number {
  for (let index = start + 1; index < text.length; index += 1) {
    if (text[index] !== quote) {
      continue;
    }
    if (text[index + 1] === quote) {
      index += 1;
      continue;
    }
    return index + 1;
  }
  return text.length;
}

function balancedBracketEnd(text: string, start: number): number | undefined {
  let depth = 0;
  for (let index = start; index < text.length; index += 1) {
    if (text[index] === "[") {
      depth += 1;
    } else if (text[index] === "]") {
      depth -= 1;
      if (depth === 0) {
        return index + 1;
      }
    }
  }
  return undefined;
}

function externalQualifierEnd(text: string, start: number): number | undefined {
  const bookEnd = text.indexOf("]", start + 1);
  if (bookEnd < 0) {
    return undefined;
  }
  const bang = text.indexOf("!", bookEnd + 1);
  if (bang < 0) {
    return undefined;
  }
  const qualifier = text.slice(bookEnd + 1, bang);
  if (!(qualifier && /^[\p{L}\p{N}_.]+$/u.test(qualifier))) {
    return undefined;
  }
  return bang + 1;
}

function threeDQualifierEnd(text: string, start: number): number | undefined {
  const match = /^[\p{L}\p{N}_.]+:[\p{L}\p{N}_.]+!/u.exec(text.slice(start));
  return match ? start + match[0].length : undefined;
}

function scanFormula(formula: string, originRow: number, originCol: number): ScanResult {
  const text = formula.trim();
  let fingerprint = "";
  let supported = true;
  let index = 0;

  while (index < text.length) {
    const char = text[index];

    // Excel string literals use doubled quotes for an embedded quote. Copy the
    // complete literal so A1-looking text can never become a reference.
    if (char === '"') {
      const end = quotedStringEnd(text, index, '"');
      fingerprint += text.slice(index, end);
      index = end;
      continue;
    }

    // Quoted sheet names use doubled apostrophes. Their spelling and case are
    // significant to the caller, and their contents are not formula tokens.
    if (char === "'") {
      const end = quotedStringEnd(text, index, "'");
      fingerprint += text.slice(index, end);
      index = end;
      continue;
    }

    // External workbook qualifiers are handled before bracketed structured
    // references. This keeps `[Book]Sheet!A1` from being treated as a table.
    if (char === "[") {
      const externalEnd = externalQualifierEnd(text, index);
      if (externalEnd !== undefined) {
        fingerprint += text.slice(index, externalEnd).toUpperCase();
        index = externalEnd;
        continue;
      }
      const end = balancedBracketEnd(text, index);
      if (end === undefined) {
        fingerprint += text.slice(index);
        supported = false;
        break;
      }
      const bracketed = text.slice(index, end);
      if (!/^\[(?:[@#]|\[)/.test(bracketed)) {
        supported = false;
      }
      fingerprint += bracketed;
      index = end;
      continue;
    }

    const threeDQualifier = threeDQualifierEnd(text, index);
    if (threeDQualifier !== undefined) {
      fingerprint += text.slice(index, threeDQualifier).toUpperCase();
      index = threeDQualifier;
      continue;
    }

    // A table name immediately followed by a balanced bracket expression is a
    // structured reference. Preserve the whole construct instead of rewriting
    // a column-looking prefix such as `Table1`.
    if (/[A-Za-z_]/.test(char)) {
      const identifier = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(text.slice(index));
      if (identifier) {
        const identifierEnd = index + identifier[0].length;
        if (text[identifierEnd] === "[") {
          const end = balancedBracketEnd(text, identifierEnd);
          if (end !== undefined) {
            fingerprint += text.slice(index, end);
            index = end;
            continue;
          }
        }
      }
    }

    const wholeColumn = parseWholeColumn(text, index);
    if (wholeColumn) {
      fingerprint += `${normalizeColumn(wholeColumn.left, originCol)}:${normalizeColumn(
        wholeColumn.right,
        originCol,
      )}`;
      index = wholeColumn.end;
      continue;
    }
    const wholeRow = parseWholeRow(text, index);
    if (wholeRow) {
      fingerprint += `${normalizeRow(wholeRow.left, originRow)}:${normalizeRow(
        wholeRow.right,
        originRow,
      )}`;
      index = wholeRow.end;
      continue;
    }

    const reference = parseCellReference(text, index);
    if (reference && !isIdentifierChar(text[index - 1])) {
      // LOG10(…) and similar function names are not A1 references. A token
      // immediately before ! is a sheet qualifier and is copied as such.
      let lookahead = reference.end;
      while (/\s/.test(text[lookahead] ?? "")) {
        lookahead += 1;
      }
      const functionName = /^[A-Za-z]+\d+$/.test(reference.text);
      if ((functionName && text[lookahead] === "(") || text[lookahead] === "!") {
        fingerprint += reference.text.toUpperCase();
        index = reference.end;
        continue;
      }
      fingerprint += normalizeCell(reference, originRow, originCol);
      index = reference.end;
      continue;
    }

    // Keep identifiers containing an underscore/dot intact. This is especially
    // important for names such as A1_NAME, which must never be split as a cell.
    if (/[A-Za-z_]/.test(char)) {
      const identifier = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(text.slice(index));
      if (identifier) {
        const token = identifier[0];
        fingerprint += token.toUpperCase();
        index += token.length;
        continue;
      }
    }

    fingerprint += char.toUpperCase();
    index += 1;
  }

  return { fingerprint, supported };
}

/** Returns a relative R1C1-like formula fingerprint for an origin cell. */
export function formulaFingerprint(formula: string, originRow: number, originCol: number): string {
  return scanFormula(formula, originRow, originCol).fingerprint;
}

/** Internal audit metadata; exported for focused lexer tests and diagnostics. */
export function formulaFingerprintDetails(
  formula: string,
  originRow: number,
  originCol: number,
): { fingerprint: string; supported: boolean } {
  return scanFormula(formula, originRow, originCol);
}
