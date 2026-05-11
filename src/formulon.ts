import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import createFormulon, {
  type FormulonModule,
  type Status,
  type Value,
  type Workbook,
} from "@libraz/formulon";

export type { Status, Workbook };

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

/** Returns the singleton Formulon WASM module instance. */
export function formulonModule(): Promise<FormulonModule> {
  modulePromise ??= createFormulon() as Promise<FormulonModule>;
  return modulePromise;
}

/** Ensures a formula has a leading equals sign. */
export function normalizeFormula(formula: string): string {
  const trimmed = formula.trim();
  return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
}

/** Serializes an MCP tool response payload as stable pretty JSON. */
export function jsonText(value: unknown): string {
  return JSON.stringify(value, null, 2);
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
      return { kind, errorCode: value.errorCode };
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

function isVectorLike(value: unknown): value is {
  size(): number;
  get(index: number): unknown;
  delete(): void;
} {
  return (
    typeof value === "object" &&
    value !== null &&
    "size" in value &&
    "get" in value &&
    "delete" in value &&
    typeof (value as { size: unknown }).size === "function" &&
    typeof (value as { get: unknown }).get === "function" &&
    typeof (value as { delete: unknown }).delete === "function"
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
  if (isVectorLike(value)) {
    try {
      return Array.from({ length: value.size() }, (_, index) => resultToJson(value.get(index)));
    } finally {
      value.delete();
    }
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
        const cell = wb.cellAt(sheet, idx);
        assertStatus(cell.status, `read sheet ${sheet} cell ${idx}`);
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
    definedNames.push({ name: entry.name, formula: entry.formula });
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

/** Saves a workbook to disk and returns the byte length written by Formulon. */
export async function saveWorkbook(wb: Workbook, outputPath: string): Promise<number> {
  const saved = wb.save();
  assertStatus(saved.status, "save workbook");
  if (!saved.bytes) {
    throw new Error("save workbook failed: no bytes returned");
  }
  const resolved = resolveUserPath(outputPath);
  await writeFile(resolved, saved.bytes);
  return saved.bytes.byteLength;
}
