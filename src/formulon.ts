import { randomUUID } from "node:crypto";
import { readFile, rename, unlink, writeFile } from "node:fs/promises";
import path from "node:path";
import createFormulon, {
  AnchorEditAs,
  AnchorKind,
  type Border,
  type BorderSide,
  CalcMode,
  type CellXf,
  ColorContext,
  ColorKind,
  ColorResolution,
  DisplayStatus,
  ErrorDisplayName,
  ExcelError,
  enumName,
  type Fill,
  type Font,
  type Formulon,
  FormulonError,
  GeometryMode,
  type HeaderFooterInput,
  ImageFormat,
  type NativeValue,
  type PageMarginsInput,
  type PageSetupInput,
  type PrintOptionsInput,
  type ProfileId,
  type ReadDiagnostics,
  type SaveDiagnostics,
  type Sheet,
  SheetVisibility,
  SparklineAxisType,
  SparklineType,
  StatusCode,
  type Value,
  ValueKind,
  type Workbook,
  WorkbookFormat,
} from "@libraz/formulon";
import { cellToA1 } from "./a1.js";

export type {
  Border,
  BorderSide,
  CellXf,
  Fill,
  Font,
  HeaderFooterInput,
  PageMarginsInput,
  PageSetupInput,
  PrintOptionsInput,
  ProfileId,
  Sheet,
  Workbook,
};
export {
  AnchorEditAs,
  AnchorKind,
  ColorContext,
  ColorKind,
  ColorResolution,
  DisplayStatus,
  ExcelError,
  FormulonError,
  GeometryMode,
  ImageFormat,
  SheetVisibility,
  SparklineAxisType,
  SparklineType,
  StatusCode,
};

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

let modulePromise: Promise<Formulon> | undefined;
let loadedModule: Formulon | undefined;

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
export function formulonModule(): Promise<Formulon> {
  modulePromise ??= createFormulon({ print: toStderr, printErr: toStderr }).then((module) => {
    loadedModule = module;
    return module;
  });
  return modulePromise;
}

/** The engine module, for synchronous callers; it is loaded before any workbook session exists. */
export function loadedFormulon(): Formulon {
  if (loadedModule === undefined) {
    throw statusError(StatusCode.PreconditionFailed, "the Formulon engine is not loaded yet");
  }
  return loadedModule;
}

/** Returns the Excel error literal for an `ErrorCode` ordinal. */
export function errorName(errorCode: number): string {
  return ErrorDisplayName[errorCode as keyof typeof ErrorDisplayName] ?? `#ERR(${errorCode})`;
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

/** Converts a tagged Formulon `Value` into a compact JSON shape for MCP responses. */
export function valueToJson(value: Value) {
  const kind = VALUE_KIND[value.kind as keyof typeof VALUE_KIND] ?? "unknown";
  switch (value.kind) {
    case ValueKind.Blank:
      return { kind };
    case ValueKind.Number:
      return { kind, value: value.number };
    case ValueKind.Bool:
      return { kind, value: value.boolean };
    case ValueKind.Text:
      return { kind, value: value.text };
    case ValueKind.Error:
      return { kind, errorCode: value.errorCode, errorName: errorName(value.errorCode) };
    default:
      return { kind, raw: value };
  }
}

/** Converts a native cell value (`null`, number, boolean, string, `ExcelError`) to the `valueToJson` shape. */
export function nativeToJson(value: NativeValue) {
  if (value === null) {
    return { kind: VALUE_KIND[0] };
  }
  if (typeof value === "number") {
    return { kind: VALUE_KIND[1], value };
  }
  if (typeof value === "boolean") {
    return { kind: VALUE_KIND[2], value };
  }
  if (typeof value === "string") {
    return { kind: VALUE_KIND[3], value };
  }
  if (value instanceof ExcelError) {
    return { kind: VALUE_KIND[4], errorCode: value.code, errorName: errorName(value.code) };
  }
  return { kind: VALUE_KIND[5], raw: value };
}

/** The per-step status JSON a tool result carries: `{ ok, status, message, context }`. */
export type StatusJson = { ok: boolean; status: number; message: string; context: string };

/** The status JSON of a step that completed. */
export function okStatus(): StatusJson {
  return { ok: true, status: 0, message: "", context: "" };
}

/**
 * The status JSON of a step that failed: a `FormulonError` keeps its code, message and context;
 * any other error reports `InternalError` with its message.
 */
export function errorStatus(error: unknown): StatusJson {
  if (error instanceof FormulonError) {
    return { ok: false, status: error.code, message: error.message, context: error.context };
  }
  return {
    ok: false,
    status: StatusCode.InternalError,
    message: error instanceof Error ? error.message : String(error),
    context: "",
  };
}

/** Runs one step and reports its outcome as status JSON instead of throwing. */
export function stepStatus(step: () => void): StatusJson {
  try {
    step();
    return okStatus();
  } catch (error) {
    return errorStatus(error);
  }
}

/**
 * The text of a tool error: a `FormulonError` leads with its status name, then the message,
 * the context and, for a formula syntax error, where the parse failed.
 */
export function errorText(error: unknown): string {
  if (!(error instanceof FormulonError)) {
    return error instanceof Error ? error.message : String(error);
  }
  const parts = [`${enumName(StatusCode, error.code) ?? error.code}: ${error.message}`];
  if (error.context) {
    parts.push(`(${error.context})`);
  }
  const parse = error.parseError;
  if (parse) {
    parts.push(
      `[parse ${parse.code} at ${parse.start}-${parse.end}, line ${parse.line} column ${parse.column}]`,
    );
  }
  return parts.join(" ");
}

/**
 * An error the server itself raises, in the engine's own form: tool results report it as
 * `<StatusName>: <message> (<context>)` with a JSON item, like an engine `FormulonError`.
 */
export function statusError(code: StatusCode, message: string, context?: string): FormulonError {
  return new FormulonError(code, message, context);
}

/** Converts arbitrary Formulon return values into JSON-friendly data. */
export function resultToJson(value: unknown): unknown {
  if (value instanceof ExcelError) {
    return nativeToJson(value);
  }
  if (isTaggedValue(value)) {
    return valueToJson(value);
  }
  if (value instanceof Uint8Array) {
    return { byteLength: value.byteLength };
  }
  // Other typed arrays carry numbers (collation class ids), not opaque bytes.
  if (ArrayBuffer.isView(value) && !(value instanceof DataView)) {
    return Array.from(value as unknown as ArrayLike<number>);
  }
  if (value instanceof Map) {
    return Object.fromEntries([...value].map(([key, entry]) => [String(key), resultToJson(entry)]));
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

function isTaggedValue(value: unknown): value is Value {
  if (typeof value !== "object" || value === null || !("kind" in value)) {
    return false;
  }
  const record = value as Record<string, unknown>;
  switch (record.kind) {
    case ValueKind.Blank:
      return Object.keys(record).length === 1;
    case ValueKind.Number:
      return typeof record.number === "number";
    case ValueKind.Bool:
      return typeof record.boolean === "boolean";
    case ValueKind.Text:
      return typeof record.text === "string";
    case ValueKind.Error:
      return typeof record.errorCode === "number";
    default:
      return false;
  }
}

/** Whether the workbook uses the 1904 date system (`<workbookPr date1904>`). */
export function isDate1904(wb: Workbook): boolean {
  return wb.getDate1904();
}

/** Resolves a user-supplied path relative to the MCP server process cwd. */
export function resolveUserPath(filePath: string): string {
  if (!filePath.trim()) {
    throw statusError(StatusCode.InvalidArgument, "path must not be empty");
  }
  return path.resolve(process.cwd(), filePath);
}

/**
 * Loads an xlsx or xlsb workbook from disk. With `cachedValues` the workbook is put in Manual
 * calculation, so reads return the values the file cached until an explicit recalc, as Excel
 * shows a file it opens in Manual mode; under Automatic a profile's open-time rule may compute
 * them first. The read-only one-shot tools load this way; a workbook that is saved back is not,
 * since the calculation mode is part of the file.
 */
export async function loadWorkbook(
  filePath: string,
  options: { cachedValues?: boolean; profile?: ProfileId } = {},
): Promise<Workbook> {
  const Module = await formulonModule();
  const bytes = await readFile(resolveUserPath(filePath));
  // The profile is given at load, so a manual-calculation file keeps the stale state it was saved with.
  const wb = Module.Workbook.load(
    bytes,
    options.profile === undefined ? undefined : { profileId: options.profile },
  );
  if (options.cachedValues) {
    wb.setCalcMode(CalcMode.Manual);
  }
  return wb;
}

/** Shape of an Excel profile id; the engine decides which ids exist. */
export const PROFILE_ID_PATTERN = /^(mac|win)-365-[a-z]{2}_[A-Z]{2}$/;

/** Environment variable naming the profile a workbook gets when none is requested. */
export const PROFILE_ENV = "FORMULON_MCP_PROFILE";

/** Reports whether the engine knows `profile`. */
export async function isKnownProfile(profile: string): Promise<boolean> {
  return (await formulonModule()).listProfileIds().includes(profile as ProfileId);
}

/**
 * Loads an existing workbook or creates a single-sheet one. A profile shapes a
 * new workbook as Excel creates it in that locale (fonts, style and sheet
 * names, margins) and sets the profile a loaded workbook evaluates with. An
 * unrequested profile falls back to `FORMULON_MCP_PROFILE`, then the engine
 * default.
 */
export async function createOrLoadWorkbook(
  filePath?: string,
  requested?: ProfileId,
  options: { cachedValues?: boolean } = {},
): Promise<Workbook> {
  const configured = process.env[PROFILE_ENV]?.trim() || undefined;
  const profile = requested ?? (configured as ProfileId | undefined);
  const source = requested === undefined ? ` (from ${PROFILE_ENV})` : "";
  if (profile !== undefined && !(await isKnownProfile(profile))) {
    throw statusError(StatusCode.InvalidArgument, `unknown excel profile: ${profile}${source}`);
  }
  if (filePath) {
    return loadWorkbook(filePath, { ...options, profile });
  }
  const Module = await formulonModule();
  return profile === undefined
    ? Module.Workbook.createDefault()
    : Module.Workbook.createForProfile(profile);
}

/** Resolves either a zero-based sheet index or a sheet name to a zero-based sheet index. */
export function findSheetIndex(wb: Workbook, sheet: number | string | undefined): number {
  const sheetCount = wb.sheetCount;
  if (sheet === undefined) {
    if (sheetCount === 0) {
      throw statusError(StatusCode.PreconditionFailed, "workbook has no sheets");
    }
    return 0;
  }
  if (typeof sheet === "number") {
    if (!Number.isInteger(sheet) || sheet < 0 || sheet >= sheetCount) {
      throw statusError(
        StatusCode.SheetIndexOutOfRange,
        "sheet index out of range",
        `sheet_index=${sheet} sheet_count=${sheetCount}`,
      );
    }
    return sheet;
  }
  const found = wb.listSheets().find((handle) => handle.name === sheet);
  if (found === undefined) {
    throw statusError(StatusCode.NotFound, "sheet not found", `sheet=${sheet}`);
  }
  return found.index;
}

/** The `Sheet` handle of a sheet index or name (the first sheet when omitted). */
export function sheetAt(wb: Workbook, sheet: number | string | undefined): Sheet {
  return wb.sheet(findSheetIndex(wb, sheet));
}

/**
 * `localSheetId` value denoting a workbook-scoped defined name, as opposed to
 * one scoped to a single sheet. Matches the engine's own sentinel.
 */
export const WORKBOOK_SCOPE = -1;

/** Builds a summary of workbook sheets, defined names, tables, and optionally sparse cells. */
export function workbookSummary(wb: Workbook, includeCells: boolean, maxCells: number) {
  const sheets: SheetJson[] = [];
  for (const handle of wb.listSheets()) {
    const cells = handle.listCells();
    const sheetJson: SheetJson = {
      index: handle.index,
      name: handle.name,
      cellCount: cells.length,
    };
    if (includeCells) {
      const limit = Math.max(0, Math.min(maxCells, cells.length));
      sheetJson.cells = cells.slice(0, limit).map((cell) => ({
        row: cell.row,
        col: cell.col,
        formula: cell.formula ?? "",
        value: nativeToJson(cell.value),
      }));
      sheetJson.cellsTruncated = limit < cells.length;
    }
    sheets.push(sheetJson);
  }

  // Scope is reported because a sheet-local name such as `_xlnm.Print_Area`
  // exists once per sheet: without it, several entries share one name and
  // differ only in the sheet their formula happens to mention.
  const definedNames = wb.listDefinedNames().map((entry) => ({
    name: entry.name,
    formula: entry.formula,
    localSheetId: entry.scope ?? WORKBOOK_SCOPE,
    ...(entry.comment ? { comment: entry.comment } : {}),
    ...(entry.hidden ? { hidden: true } : {}),
  }));

  const tables = [];
  for (const handle of wb.listSheets()) {
    for (const table of handle.listTables()) {
      const properties = table.getProperties();
      tables.push({
        index: tables.length,
        name: properties.name,
        displayName: properties.name,
        ref: rectA1(properties.range.rect),
        sheetIndex: properties.range.sheet,
      });
    }
  }

  return { sheets, definedNames, tables };
}

/** The A1 text of a rectangle (`A1` for one cell, `A1:C3` otherwise). */
export function rectA1(rect: {
  firstRow: number;
  firstCol: number;
  lastRow: number;
  lastCol: number;
}): string {
  const first = cellToA1(rect.firstRow, rect.firstCol);
  return rect.firstRow === rect.lastRow && rect.firstCol === rect.lastCol
    ? first
    : `${first}:${cellToA1(rect.lastRow, rect.lastCol)}`;
}

/** Applies a concrete zero-based cell mutation to a workbook; a refusal throws `FormulonError`. */
export function applyMutation(wb: Workbook, mutation: CellMutation): void {
  const sheet = wb.sheet(mutation.sheet);
  switch (mutation.type) {
    case "number":
      sheet.setNumber(mutation.row, mutation.col, mutation.value);
      return;
    case "bool":
      sheet.setBool(mutation.row, mutation.col, mutation.value);
      return;
    case "text":
      sheet.setText(mutation.row, mutation.col, mutation.value);
      return;
    case "blank":
      sheet.setBlank(mutation.row, mutation.col);
      return;
    case "formula":
      sheet.setFormula(mutation.row, mutation.col, normalizeFormula(mutation.formula));
      return;
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
  const diagnostics: ReadDiagnostics = wb.getDiagnostics();
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
  const extension = path.extname(outputPath).toLowerCase();
  if (extension === ".xlsx") {
    return WorkbookFormat.Xlsx;
  }
  if (extension === ".xlsb") {
    return WorkbookFormat.Xlsb;
  }
  throw statusError(StatusCode.InvalidArgument, "outputPath must end with .xlsx or .xlsb");
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
  const saved: SaveDiagnostics = wb.saveWithDiagnostics(format);
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
