import { cellToA1 } from "../a1.js";
import { assertStatus, requireCell, requireNumberResult, statusToJson } from "../formulon.js";
import { cellDisplayText, searchSheets, sheetName, textValue } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

export type SearchTarget = "texts" | "formulas" | "both";

export type SearchOptions = {
  sheet?: number | string;
  target: SearchTarget;
  matchCase: boolean;
  wholeCell: boolean;
  regex: boolean;
  maxResults: number;
};

export type ReplaceOptions = SearchOptions & {
  replacement: string;
  maxReplacements: number;
  recalc: boolean;
};
type CellSearchResult = {
  sheet: number;
  sheetName: string;
  row: number;
  col: number;
  a1: string;
  ref: string;
  target: "text" | "formula";
  text: string;
};

type CellReplaceResult = CellSearchResult & {
  before: string;
  after: string;
  status: unknown;
};
function makeMatcher(
  query: string,
  options: Pick<SearchOptions, "matchCase" | "wholeCell" | "regex">,
) {
  if (options.regex) {
    const flags = options.matchCase ? "g" : "gi";
    const pattern = new RegExp(query, flags);
    return (text: string) => {
      pattern.lastIndex = 0;
      const matched = pattern.test(text);
      pattern.lastIndex = 0;
      return matched;
    };
  }

  const needle = options.matchCase ? query : query.toLocaleLowerCase();
  return (text: string) => {
    const haystack = options.matchCase ? text : text.toLocaleLowerCase();
    return options.wholeCell ? haystack === needle : haystack.includes(needle);
  };
}

function replaceText(
  text: string,
  query: string,
  replacement: string,
  options: Pick<SearchOptions, "matchCase" | "wholeCell" | "regex">,
): string {
  if (options.regex) {
    return text.replace(new RegExp(query, options.matchCase ? "g" : "gi"), replacement);
  }
  if (options.wholeCell) {
    return options.matchCase
      ? text === query
        ? replacement
        : text
      : text.toLocaleLowerCase() === query.toLocaleLowerCase()
        ? replacement
        : text;
  }
  const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return text.replace(new RegExp(escaped, options.matchCase ? "g" : "gi"), replacement);
}
/** Searches text cells and/or formula text across an open workbook session. */
export function findSessionCells(id: string, query: string, options: SearchOptions) {
  const session = getSession(id);
  if (query.length === 0) {
    throw new Error("query must not be empty");
  }
  const matcher = makeMatcher(query, options);
  const results: CellSearchResult[] = [];
  let truncated = false;

  for (const sheet of searchSheets(session, options.sheet)) {
    const name = sheetName(session, sheet);
    const count = requireNumberResult(
      session.workbook.cellCount(sheet),
      `read sheet ${sheet} cell count`,
    );
    for (let index = 0; index < count; index += 1) {
      const cell = requireCell(
        session.workbook.cellAt(sheet, index),
        `read sheet ${sheet} cell ${index}`,
      );
      const a1 = cellToA1(cell.row, cell.col);
      const addResult = (target: "text" | "formula", text: string) => {
        if (!matcher(text)) {
          return;
        }
        if (results.length >= options.maxResults) {
          truncated = true;
          return;
        }
        results.push({
          sheet,
          sheetName: name,
          row: cell.row,
          col: cell.col,
          a1,
          ref: `${name}!${a1}`,
          target,
          text,
        });
      };

      if (cell.formula) {
        if (options.target === "formulas" || options.target === "both") {
          addResult("formula", cell.formula);
        }
      } else {
        const text = cellDisplayText(cell.value);
        if ((options.target === "texts" || options.target === "both") && text !== undefined) {
          addResult("text", text);
        }
      }
      if (truncated) {
        break;
      }
    }
    if (truncated) {
      break;
    }
  }

  touch(session, false);
  return {
    session: publicInfo(session),
    query,
    options,
    results,
    count: results.length,
    truncated,
  };
}

/** Replaces matching text cell values and/or formula text across an open workbook session. */
export function replaceSessionCells(id: string, query: string, options: ReplaceOptions) {
  const session = getSession(id);
  if (query.length === 0) {
    throw new Error("query must not be empty");
  }
  const matcher = makeMatcher(query, options);
  const replacements: CellReplaceResult[] = [];
  let truncated = false;

  for (const sheet of searchSheets(session, options.sheet)) {
    const name = sheetName(session, sheet);
    const count = requireNumberResult(
      session.workbook.cellCount(sheet),
      `read sheet ${sheet} cell count`,
    );
    for (let index = 0; index < count; index += 1) {
      const cell = requireCell(
        session.workbook.cellAt(sheet, index),
        `read sheet ${sheet} cell ${index}`,
      );
      const a1 = cellToA1(cell.row, cell.col);
      const replaceOne = (target: "text" | "formula", text: string) => {
        if (!matcher(text)) {
          return;
        }
        if (replacements.length >= options.maxReplacements) {
          truncated = true;
          return;
        }
        const next = replaceText(text, query, options.replacement, options);
        // The engine may commit a partial change before reporting a failure;
        // retain dirty state from the moment a mutating call is attempted.
        touch(session, true);
        const status =
          target === "formula"
            ? session.workbook.setFormula(sheet, cell.row, cell.col, next)
            : session.workbook.setText(sheet, cell.row, cell.col, next);
        assertStatus(status, `replace ${target} ${name}!${a1}`);
        replacements.push({
          sheet,
          sheetName: name,
          row: cell.row,
          col: cell.col,
          a1,
          ref: `${name}!${a1}`,
          target,
          text,
          before: text,
          after: next,
          status: statusToJson(status),
        });
      };

      if (cell.formula) {
        if (options.target === "formulas" || options.target === "both") {
          replaceOne("formula", cell.formula);
        }
      } else {
        const text = textValue(cell.value);
        if ((options.target === "texts" || options.target === "both") && text !== undefined) {
          replaceOne("text", text);
        }
      }
      if (truncated) {
        break;
      }
    }
    if (truncated) {
      break;
    }
  }

  if (replacements.length > 0 && options.recalc) {
    assertStatus(session.workbook.recalc(), "recalc workbook");
  }
  if (replacements.length === 0) {
    touch(session, false);
  }
  return {
    session: publicInfo(session),
    query,
    replacement: options.replacement,
    options,
    replacements,
    count: replacements.length,
    truncated,
  };
}
