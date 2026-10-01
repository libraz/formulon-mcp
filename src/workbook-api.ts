import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { READ_ONLY_METHODS, WORKBOOK_METHODS } from "./session/workbook-call.js";
import type { WorkbookApiDeclaration, WorkbookApiMethod } from "./workbook-api-parser.js";
import {
  type ParsedDeclaration,
  type ParsedMethod,
  parseExportedDeclarations,
  parseWorkbookMethods,
} from "./workbook-api-parser.js";

export type { WorkbookApiDeclaration, WorkbookApiMethod } from "./workbook-api-parser.js";

export const WORKBOOK_API_MAX_LIMIT = 100;

const MAX_DECLARATIONS = 64;
const MAX_DECLARATION_DEPTH = 5;
const MAX_DECLARATION_CHARS = 12_000;
const MAX_TOTAL_DECLARATION_CHARS = 48_000;
const EXCLUDED_WORKBOOK_METHODS = new Set([
  "isValid",
  "delete",
  "save",
  "saveAs",
  "saveWithDiagnostics",
  "setIterativeProgress",
]);

type PackageMetadata = {
  version?: unknown;
  types?: unknown;
};

export interface WorkbookApiCatalog {
  version: string;
  typesPath: string;
  methods: readonly WorkbookApiMethod[];
}

export interface WorkbookApiSearchResult {
  version: string;
  operation: "search";
  query?: string;
  total: number;
  truncated: boolean;
  count: number;
  methods: WorkbookApiMethod[];
}

export interface WorkbookApiDescribeResult {
  version: string;
  operation: "describe";
  total: number;
  truncated: boolean;
  count: number;
  methods: WorkbookApiMethod[];
  declarations: WorkbookApiDeclaration[];
  declarationsTruncated: boolean;
  truncation?: string;
}

function resolveEngineDeclarations(): { version: string; typesPath: string } {
  const require = createRequire(import.meta.url);
  let packagePath: string;
  try {
    packagePath = require.resolve("@libraz/formulon/package.json");
  } catch (error) {
    throw new Error(
      `cannot resolve installed @libraz/formulon/package.json: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  let packageMetadata: PackageMetadata;
  try {
    packageMetadata = JSON.parse(readFileSync(packagePath, "utf8")) as PackageMetadata;
  } catch (error) {
    throw new Error(
      `cannot read installed @libraz/formulon/package.json: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
  if (typeof packageMetadata.version !== "string" || packageMetadata.version.length === 0) {
    throw new Error("installed @libraz/formulon/package.json has no usable version");
  }
  if (typeof packageMetadata.types !== "string" || packageMetadata.types.length === 0) {
    throw new Error("installed @libraz/formulon/package.json has no usable types path");
  }

  const typesPath = resolve(dirname(packagePath), packageMetadata.types);
  try {
    readFileSync(typesPath, "utf8");
  } catch (error) {
    throw new Error(
      `cannot read installed @libraz/formulon declaration file ${typesPath}: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
  return { version: packageMetadata.version, typesPath };
}

type ParsedCatalog = {
  catalog: WorkbookApiCatalog;
  declarations: Map<string, ParsedDeclaration>;
  methods: Map<string, ParsedMethod>;
};

let cached: ParsedCatalog | undefined;

function getParsedCatalog(): ParsedCatalog {
  if (cached) {
    return cached;
  }
  const { version, typesPath } = resolveEngineDeclarations();
  const source = readFileSync(typesPath, "utf8");
  const declarations = parseExportedDeclarations(source);
  const parsedMethods = parseWorkbookMethods(source, READ_ONLY_METHODS, declarations);
  const methods: WorkbookApiMethod[] = [];
  const missing: string[] = [];
  for (const name of WORKBOOK_METHODS) {
    if (EXCLUDED_WORKBOOK_METHODS.has(name)) {
      continue;
    }
    const method = parsedMethods.get(name);
    if (!method) {
      missing.push(name);
      continue;
    }
    methods.push({
      name: method.name,
      signature: method.signature,
      description: method.description,
      access: method.access,
    });
  }
  if (missing.length > 0) {
    throw new Error(
      `installed @libraz/formulon Workbook declaration is missing allowlisted methods: ${missing.join(", ")}`,
    );
  }
  if (methods.length !== new Set(methods.map((method) => method.name)).size) {
    throw new Error("workbook API catalog contains duplicate allowlisted methods");
  }
  cached = {
    catalog: { version, typesPath, methods },
    declarations,
    methods: parsedMethods,
  };
  return cached;
}

export function getWorkbookApiCatalog(): WorkbookApiCatalog {
  return getParsedCatalog().catalog;
}

function boundedLimit(limit: number | undefined): number {
  if (limit === undefined) {
    return 20;
  }
  if (!Number.isInteger(limit) || limit < 1) {
    throw new Error(`limit must be an integer in 1..${WORKBOOK_API_MAX_LIMIT}`);
  }
  return Math.min(limit, WORKBOOK_API_MAX_LIMIT);
}

function boundedOffset(offset: number | undefined): number {
  if (offset === undefined) {
    return 0;
  }
  if (!Number.isInteger(offset) || offset < 0) {
    throw new Error("offset must be a non-negative integer");
  }
  return offset;
}

export function searchWorkbookApi(
  query?: string,
  limit?: number,
  offset?: number,
): WorkbookApiSearchResult {
  const { catalog } = getParsedCatalog();
  const bounded = boundedLimit(limit);
  const start = boundedOffset(offset);
  const normalizedQuery = query?.trim();
  const needle = normalizedQuery?.toLowerCase();
  const matching = needle
    ? catalog.methods.filter(
        (method) =>
          method.name.toLowerCase().includes(needle) ||
          method.signature.toLowerCase().includes(needle) ||
          method.description.toLowerCase().includes(needle),
      )
    : [...catalog.methods];
  const methods = matching.slice(start, start + bounded);
  return {
    version: catalog.version,
    operation: "search",
    ...(normalizedQuery ? { query: normalizedQuery } : {}),
    total: matching.length,
    truncated: start + methods.length < matching.length,
    count: methods.length,
    methods,
  };
}

function truncateDeclaration(declaration: string): { text: string; truncated: boolean } {
  if (declaration.length <= MAX_DECLARATION_CHARS) {
    return { text: declaration, truncated: false };
  }
  return {
    text: `${declaration.slice(0, MAX_DECLARATION_CHARS)}\n/* declaration truncated at ${MAX_DECLARATION_CHARS} characters */`,
    truncated: true,
  };
}

export function describeWorkbookApi(methodName: string): WorkbookApiDescribeResult {
  const { catalog, declarations, methods } = getParsedCatalog();
  const method = catalog.methods.find((entry) => entry.name === methodName);
  if (!method) {
    throw new Error(`unknown or unavailable Workbook method: ${methodName}`);
  }

  const declarationsOut: WorkbookApiDeclaration[] = [];
  const seen = new Set<string>();
  const queue: Array<{ name: string; depth: number }> = [];
  for (const name of methods.get(method.name)?.references ?? []) {
    queue.push({ name, depth: 0 });
  }
  let totalChars = 0;
  let declarationsTruncated = false;
  let depthTruncated = false;
  while (queue.length > 0) {
    const current = queue.shift();
    if (!current || seen.has(current.name)) {
      continue;
    }
    seen.add(current.name);
    const declaration = declarations.get(current.name);
    if (!declaration) {
      continue;
    }
    if (declarationsOut.length >= MAX_DECLARATIONS) {
      declarationsTruncated = true;
      break;
    }
    if (current.depth > MAX_DECLARATION_DEPTH) {
      depthTruncated = true;
      continue;
    }
    const bounded = truncateDeclaration(declaration.declaration);
    totalChars += bounded.text.length;
    if (totalChars > MAX_TOTAL_DECLARATION_CHARS) {
      declarationsTruncated = true;
      break;
    }
    declarationsOut.push({
      name: declaration.name,
      kind: declaration.kind,
      description: declaration.description,
      declaration: bounded.text,
    });
    if (bounded.truncated) {
      declarationsTruncated = true;
    }
    if (current.depth < MAX_DECLARATION_DEPTH) {
      for (const name of declaration.references) {
        if (!seen.has(name)) {
          queue.push({ name, depth: current.depth + 1 });
        }
      }
    } else if (declaration.references.length > 0) {
      depthTruncated = true;
    }
  }
  const truncationReasons: string[] = [];
  if (declarationsTruncated) {
    truncationReasons.push("declaration size or count limit reached");
  }
  if (depthTruncated) {
    truncationReasons.push("reference depth limit reached");
  }
  return {
    version: catalog.version,
    operation: "describe",
    total: 1,
    truncated: declarationsTruncated || depthTruncated,
    count: 1,
    methods: [method],
    declarations: declarationsOut,
    declarationsTruncated: declarationsTruncated || depthTruncated,
    ...(truncationReasons.length > 0 ? { truncation: truncationReasons.join("; ") } : {}),
  };
}
