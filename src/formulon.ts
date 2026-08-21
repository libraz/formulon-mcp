import { randomUUID } from "node:crypto";
import { readFile, rename, unlink, writeFile } from "node:fs/promises";
import path from "node:path";
import createFormulon, {
  type BorderRecord,
  type BorderSide,
  type CellEntry,
  type CellXf,
  type FillRecord,
  type FontRecord,
  type FormulonModule,
  type HeaderFooterInput,
  type PageMarginsInput,
  type PageSetupInput,
  type PrintOptionsInput,
  type ReadDiagnosticsResult,
  type SaveDiagnosticsResult,
  SheetVisibility,
  type Status,
  type Value,
  type Workbook,
  WorkbookFormat,
} from "@libraz/formulon";

export type {
  BorderRecord,
  BorderSide,
  CellXf,
  FillRecord,
  FontRecord,
  HeaderFooterInput,
  PageMarginsInput,
  PageSetupInput,
  PrintOptionsInput,
  Status,
  Workbook,
};
export { SheetVisibility };

export type CellMutation =
  | { type: "number"; sheet: number; row: number; col: number; value: number }
  | { type: "bool"; sheet: number; row: number; col: number; value: boolean }
  | { type: "text"; sheet: number; row: number; col: number; value: string }
  | { type: "blank"; sheet: number; row: number; col: number }
  | { type: "formula"; sheet: number; row: number; col: number; formula: string };

export type SheetJson = {
  index: number;
  name: string;
  cellCount: number;
  cells?: unknown[];
  cellsTruncated?: boolean;
};

const VALUE_KIND = Object.freeze({
  0: "blank",
  1: "number",
  2: "bool",
  3: "text",
  4: "error",
  5: "array",
  6: "ref",
  7: "lambda",
} as const);

let modulePromise: Promise<FormulonModule> | undefined;
let loadedModule: FormulonModule | undefined;

/** Sends one engine output line to stderr, keeping fd 1 free for JSON-RPC. */
function toStderr(message: string): void {
  process.stderr.write(`${message}\n`);
}

/**
 * Returns the singleton Formulon WASM module instance.
 *
 * Both engine output streams are pinned to stderr. The MCP stdio transport owns
 * fd 1, and Emscripten's Node default writes there with `fs.writeSync(1, ...)`
 * rather than `console.log` — so a single line of engine output would land in
 * the middle of a JSON-RPC frame, and replacing `console.log` would not stop it.
 */
export function formulonModule(): Promise<FormulonModule> {
  modulePromise ??= (
    createFormulon({ print: toStderr, printErr: toStderr }) as Promise<FormulonModule>
  ).then((module) => {
    // Cache the resolved instance so synchronous helpers such as errorName()
    // can reach the engine without threading a promise through every caller.
    loadedModule = module;
    return module;
  });
  return modulePromise;
}

/**
 * Returns the Excel error literal for an `ErrorCode` ordinal.
 *
 * Ordinals are the engine's own enum order (not Excel's BIFF codes), so the
 * mapping is asked of the engine rather than restated here. Every read that
 * can surface an error code happens after a workbook is open, so the module is
 * loaded by then; the ordinal fallback only covers a call made before that.
 */
export function errorName(errorCode: number): string {
  return loadedModule?.errorDisplayName(errorCode) ?? `#ERR(${errorCode})`;
}

/**
 * Fields of a `cellAt` entry, which the engine leaves unset when the read
 * fails. Throws on a failed read so callers get the populated record.
 */
export type ResolvedCell = {
  row: number;
  col: number;
  formula: string;
  value: Value;
};

/** Asserts a `cellAt` read succeeded and returns its populated fields. */
export function requireCell(entry: CellEntry, action: string): ResolvedCell {
  assertStatus(entry.status, action);
  if (entry.row === undefined || entry.col === undefined || entry.value === undefined) {
    throw new Error(`${action} failed: incomplete cell record`);
  }
  return { row: entry.row, col: entry.col, formula: entry.formula ?? "", value: entry.value };
}

/** Returns a `cellAt` entry's fields, or null when the read failed. */
export function readCell(entry: CellEntry): ResolvedCell | null {
  if (
    !entry.status.ok ||
    entry.row === undefined ||
    entry.col === undefined ||
    entry.value === undefined
  ) {
    return null;
  }
  return { row: entry.row, col: entry.col, formula: entry.formula ?? "", value: entry.value };
}

/** Ensures a formula has a leading equals sign. */
export function normalizeFormula(formula: string): string {
  const trimmed = formula.trim();
  return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
}

/**
 * Serializes an MCP tool response payload as compact JSON.
 *
 * Compact output keeps token cost proportional to the data; set
 * `FORMULON_MCP_PRETTY=1` to pretty-print for human debugging.
 */
export function jsonText(value: unknown): string {
  return process.env.FORMULON_MCP_PRETTY === "1"
    ? JSON.stringify(value, null, 2)
    : JSON.stringify(value);
}

/** Converts a Formulon value envelope into a compact JSON shape for MCP responses. */
export function valueToJson(value: Value) {
  const kind = VALUE_KIND[value.kind as keyof typeof VALUE_KIND] ?? "unknown";
  switch (value.kind) {
    case 0:
      return { kind };
    case 1:
      return { kind, value: value.number };
    case 2:
      return { kind, value: Boolean(value.boolean) };
    case 3:
      return { kind, value: value.text };
    case 4:
      return { kind, errorCode: value.errorCode, errorName: errorName(value.errorCode) };
    default:
      return { kind, raw: value };
  }
}

/** Converts a Formulon status into plain JSON. */
export function statusToJson(status: Status) {
  return {
    ok: status.ok,
    status: status.status,
    message: status.message,
    context: status.context,
  };
}

function isStatusLike(value: unknown): value is Status {
  return (
    typeof value === "object" &&
    value !== null &&
    "ok" in value &&
    "status" in value &&
    typeof (value as { ok: unknown }).ok === "boolean"
  );
}

function isValueLike(value: unknown): value is Value {
  return (
    typeof value === "object" &&
    value !== null &&
    "kind" in value &&
    "number" in value &&
    "boolean" in value &&
    "text" in value &&
    "errorCode" in value
  );
}

/** Converts arbitrary Formulon return values into JSON-friendly data. */
export function resultToJson(value: unknown): unknown {
  if (isStatusLike(value)) {
    return statusToJson(value);
  }
  if (isValueLike(value)) {
    return valueToJson(value);
  }
  if (value instanceof Uint8Array) {
    return { byteLength: value.byteLength };
  }
  if (Array.isArray(value)) {
    return value.map((item) => resultToJson(item));
  }
  if (typeof value === "object" && value !== null) {
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [key, resultToJson(entry)]),
    );
  }
  return value;
}

/** Throws when a Formulon status is not ok. */
export function assertStatus(status: Status, action: string): void {
  if (status.ok) {
    return;
  }
  const detail = status.context ? `${status.message} (${status.context})` : status.message;
  throw new Error(`${action} failed: ${detail || `status ${status.status}`}`);
}

/** Resolves a user-supplied path relative to the MCP server process cwd. */
export function resolveUserPath(filePath: string): string {
  if (!filePath.trim()) {
    throw new Error("path must not be empty");
  }
  return path.resolve(process.cwd(), filePath);
}

/** Loads an xlsx workbook from disk and validates the native handle. */
export async function loadWorkbook(filePath: string): Promise<Workbook> {
  const Module = await formulonModule();
  const bytes = await readFile(resolveUserPath(filePath));
  const wb = Module.Workbook.loadBytes(bytes);
  if (!wb.isValid()) {
    const message = Module.lastErrorMessage();
    wb.delete();
    throw new Error(`load workbook failed: ${message || filePath}`);
  }
  return wb;
}

/** Loads an existing workbook or creates a default single-sheet workbook. */
export async function createOrLoadWorkbook(filePath?: string): Promise<Workbook> {
  if (filePath) {
    return loadWorkbook(filePath);
  }
  const Module = await formulonModule();
  const wb = Module.Workbook.createDefault();
  if (!wb.isValid()) {
    wb.delete();
    throw new Error("create workbook failed");
  }
  return wb;
}

/** Resolves either a zero-based sheet index or a sheet name to a zero-based sheet index. */
export function findSheetIndex(wb: Workbook, sheet: number | string | undefined): number {
  if (sheet === undefined) {
    return 0;
  }
  if (typeof sheet === "number") {
    if (!Number.isInteger(sheet) || sheet < 0 || sheet >= wb.sheetCount()) {
      throw new Error(`sheet index out of bounds: ${sheet}`);
    }
    return sheet;
  }
  for (let idx = 0; idx < wb.sheetCount(); idx += 1) {
    const name = wb.sheetName(idx);
    assertStatus(name.status, `read sheet ${idx} name`);
    if (name.value === sheet) {
      return idx;
    }
  }
  throw new Error(`sheet not found: ${sheet}`);
}

/**
 * `localSheetId` value denoting a workbook-scoped defined name, as opposed to
 * one scoped to a single sheet. Matches the engine's own sentinel.
 */
export const WORKBOOK_SCOPE = -1;

/** Builds a summary of workbook sheets, defined names, tables, and optionally sparse cells. */
export function workbookSummary(wb: Workbook, includeCells: boolean, maxCells: number) {
  const sheets: SheetJson[] = [];
  for (let sheet = 0; sheet < wb.sheetCount(); sheet += 1) {
    const name = wb.sheetName(sheet);
    assertStatus(name.status, `read sheet ${sheet} name`);
    const cellCount = wb.cellCount(sheet);
    const sheetJson: {
      index: number;
      name: string;
      cellCount: number;
      cells?: unknown[];
      cellsTruncated?: boolean;
    } = {
      index: sheet,
      name: name.value,
      cellCount,
    };
    if (includeCells) {
      const limit = Math.max(0, Math.min(maxCells, cellCount));
      sheetJson.cells = [];
      for (let idx = 0; idx < limit; idx += 1) {
        const cell = requireCell(wb.cellAt(sheet, idx), `read sheet ${sheet} cell ${idx}`);
        sheetJson.cells.push({
          row: cell.row,
          col: cell.col,
          formula: cell.formula,
          value: valueToJson(cell.value),
        });
      }
      sheetJson.cellsTruncated = limit < cellCount;
    }
    sheets.push(sheetJson);
  }

  const definedNames = [];
  for (let idx = 0; idx < wb.definedNameCount(); idx += 1) {
    const entry = wb.definedNameAt(idx);
    assertStatus(entry.status, `read defined name ${idx}`);
    // Scope is reported because a sheet-local name such as `_xlnm.Print_Area`
    // exists once per sheet: without it, several entries share one name and
    // differ only in the sheet their formula happens to mention.
    definedNames.push({
      name: entry.name,
      formula: entry.formula,
      localSheetId: entry.localSheetId ?? WORKBOOK_SCOPE,
    });
  }

  const tables = [];
  for (let idx = 0; idx < wb.tableCount(); idx += 1) {
    const entry = wb.tableAt(idx);
    assertStatus(entry.status, `read table ${idx}`);
    tables.push({
      name: entry.name,
      displayName: entry.displayName,
      ref: entry.ref,
      sheetIndex: entry.sheetIndex,
    });
  }

  return { sheets, definedNames, tables };
}

/** Applies a concrete zero-based cell mutation to a workbook. */
export function applyMutation(wb: Workbook, mutation: CellMutation): Status {
  switch (mutation.type) {
    case "number":
      return wb.setNumber(mutation.sheet, mutation.row, mutation.col, mutation.value);
    case "bool":
      return wb.setBool(mutation.sheet, mutation.row, mutation.col, mutation.value);
    case "text":
      return wb.setText(mutation.sheet, mutation.row, mutation.col, mutation.value);
    case "blank":
      return wb.setBlank(mutation.sheet, mutation.row, mutation.col);
    case "formula":
      return wb.setFormula(
        mutation.sheet,
        mutation.row,
        mutation.col,
        normalizeFormula(mutation.formula),
      );
  }
}

/** Counters describing what a save or load could not carry, keyed by event. */
export type LossCounters = Record<string, number>;

/**
 * Keeps only the non-zero counters of a diagnostics record, so a clean save or
 * load reports nothing and a lossy one reports exactly what it lost.
 */
function reportedLosses(counters: LossCounters): LossCounters | undefined {
  const lost = Object.entries(counters).filter(([, count]) => count > 0);
  return lost.length > 0 ? Object.fromEntries(lost) : undefined;
}

/** Reports what a load could not decode, or undefined for a clean load. */
export function readLosses(wb: Workbook): LossCounters | undefined {
  const diagnostics: ReadDiagnosticsResult = wb.readDiagnostics();
  if (!diagnostics.status.ok) {
    return undefined;
  }
  return reportedLosses({
    undecodedFormulas: diagnostics.undecodedFormulaCount,
    undecodedDefinedNames: diagnostics.undecodedDefinedNameCount,
    undecodedParts: diagnostics.undecodedPartCount,
    skippedFeatures: diagnostics.skippedFeatureCount,
    unknownContentTypes: diagnostics.unknownContentTypeCount,
  });
}

/** Selects the container format Formulon writes from the output extension. */
export function workbookFormatFor(outputPath: string): WorkbookFormat {
  return path.extname(outputPath).toLowerCase() === ".xlsb"
    ? WorkbookFormat.Xlsb
    : WorkbookFormat.Xlsx;
}

export type SavedWorkbook = {
  bytes: number;
  format: "xlsx" | "xlsb";
  /** Present only when the writer dropped or downgraded something. */
  losses?: LossCounters;
};

/**
 * Saves a workbook to disk, choosing the container from the output extension,
 * and reports what the writer could not carry.
 */
export async function saveWorkbook(wb: Workbook, outputPath: string): Promise<SavedWorkbook> {
  const format = workbookFormatFor(outputPath);
  const saved: SaveDiagnosticsResult = wb.saveWithDiagnostics(format);
  assertStatus(saved.status, "save workbook");
  if (!saved.bytes) {
    throw new Error("save workbook failed: no bytes returned");
  }
  const resolved = resolveUserPath(outputPath);
  // Write to a sibling temp file then atomically rename, so a crash or error
  // mid-write never truncates or corrupts an existing workbook at `resolved`.
  const tmp = `${resolved}.${randomUUID()}.tmp`;
  try {
    await writeFile(tmp, saved.bytes);
    await rename(tmp, resolved);
  } catch (error) {
    await unlink(tmp).catch(() => {});
    throw error;
  }
  return {
    bytes: saved.bytes.byteLength,
    format: format === WorkbookFormat.Xlsb ? "xlsb" : "xlsx",
    losses: reportedLosses({
      downgradedFormulas: saved.downgradedFormulaCount,
      deferredFeatures: saved.deferredFeatureCount,
      droppedParts: saved.droppedPartCount,
      droppedRelationships: saved.droppedRelationshipCount,
      renumberedParts: saved.renumberedPartCount,
    }),
  };
}
