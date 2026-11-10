import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { LEDGER_ACCESS } from "./session/ledger-access.js";
import type {
  WorkbookApiDeclaration,
  WorkbookApiMethod,
  WorkbookApiTarget,
} from "./workbook-api-parser.js";
import {
  type ParsedDeclaration,
  type ParsedInterface,
  type ParsedMethod,
  parseExportedDeclarations,
  parseInterface,
} from "./workbook-api-parser.js";

export type {
  WorkbookApiDeclaration,
  WorkbookApiMethod,
  WorkbookApiTarget,
} from "./workbook-api-parser.js";

export const WORKBOOK_API_MAX_LIMIT = 100;

const MAX_DECLARATIONS = 64;
const MAX_DECLARATION_DEPTH = 5;
const MAX_DECLARATION_CHARS = 12_000;
const MAX_TOTAL_DECLARATION_CHARS = 48_000;
/** The interface every session target is reached from. */
const ROOT_KIND = "Workbook";
/**
 * The engine module's interface: its own functions are a target with no keys. Nothing is reached
 * through it (its `Workbook` constructor would make a workbook outside any session).
 */
export const MODULE_KIND = "Formulon";
/**
 * Members other tools own: `dispose` ends the session's handle (`formulon_close_workbook`) and a
 * save needs the session's bookkeeping (`formulon_save_session`).
 */
const SESSION_OWNED_METHODS = /^(dispose|save)/;
/**
 * Signature fragments of members a JSON argument or result cannot carry: raw bytes (which
 * `formulon_save_session` and `formulon_image_operation` handle), callbacks (inline or a
 * `*Callback` alias), maps and iterators.
 */
const UNCALLABLE_SIGNATURE = /Uint8Array|BytesInput|=>|Callback\b|Map<|Iterator</;
/** Method-name prefixes that only read, for methods the ledger does not name. */
const READ_PREFIXES =
  /^(get|list|has|is|evaluate|format|resolve|parse|compare|localize|canonicalize|validate|preview|find)(?![a-z])/;
/** Method-name prefixes that change the workbook, for methods the ledger does not name. */
const WRITE_PREFIXES =
  /^(set|add|remove|update|insert|delete|clear|reset|apply|recalc|rename|move|copy|restore|remap|convert|merge|refresh|project|show)(?![a-z])/;

/** How one call reaches a handle from the workbook: an accessor method taking a key, or a property. */
export type TargetHop = { method: string; key: string } | { property: string };

/** A handle kind `formulon_workbook_call` can address, with the hops that reach it. */
export interface WorkbookApiKind {
  kind: string;
  keys: string[];
  hops: TargetHop[];
  /** The handle's scalar (number or string) properties, which identify it in a result. */
  identity: string[];
}

/**
 * Whether `method` only reads: the ledger's verb when it names the method, else the name's
 * prefix; `undefined` when neither decides. The catalog also reads a method that only hands
 * out another handle; anything still undecided is refused.
 */
export function methodAccess(
  method: string,
  accessors: ReadonlySet<string>,
): "read" | "write" | undefined {
  const recorded = LEDGER_ACCESS[method];
  if (recorded) {
    return recorded;
  }
  if (accessors.has(method) || READ_PREFIXES.test(method)) {
    return "read";
  }
  return WRITE_PREFIXES.test(method) ? "write" : undefined;
}

type PackageMetadata = {
  version?: unknown;
  types?: unknown;
};

export interface WorkbookApiCatalog {
  version: string;
  typesPath: string;
  kinds: readonly WorkbookApiKind[];
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
  /** Parsed method by `<kind>.<method>`. */
  methods: Map<string, ParsedMethod>;
};

let cached: ParsedCatalog | undefined;

const lowerFirst = (name: string) => name.charAt(0).toLowerCase() + name.slice(1);

/** The interface's own members merged over those of every interface it extends. */
function withInherited(
  name: string,
  parse: (name: string) => ParsedInterface | undefined,
): ParsedInterface | undefined {
  const own = parse(name);
  if (!own) {
    return undefined;
  }
  const merged: ParsedInterface = { ...own, methods: new Map(), properties: new Map() };
  for (const base of own.extends) {
    const inherited = withInherited(base, parse);
    for (const [key, value] of inherited?.methods ?? []) {
      merged.methods.set(key, value);
    }
    for (const [key, value] of inherited?.properties ?? []) {
      merged.properties.set(key, value);
    }
  }
  for (const [key, value] of own.methods) {
    merged.methods.set(key, value);
  }
  for (const [key, value] of own.properties) {
    merged.properties.set(key, value);
  }
  return merged;
}

/**
 * Every handle kind reachable from the workbook, nearest route first: an interface with methods
 * that a reachable kind hands out through a one-argument accessor named after it (`sheet(key)`
 * gives `Sheet`) or through a property of exactly that type (`ooxml`).
 */
function discoverKinds(
  parse: (name: string) => ParsedInterface | undefined,
): Map<string, { kind: WorkbookApiKind; members: ParsedInterface }> {
  const root = withInherited(ROOT_KIND, parse);
  if (!root) {
    throw new Error(
      "installed @libraz/formulon declaration file has no exported Workbook interface",
    );
  }
  const found = new Map<string, { kind: WorkbookApiKind; members: ParsedInterface }>([
    [ROOT_KIND, { kind: { kind: ROOT_KIND, keys: [], hops: [], identity: [] }, members: root }],
  ]);
  const queue = [ROOT_KIND];
  while (queue.length > 0) {
    const parent = found.get(queue.shift() as string);
    if (!parent) {
      continue;
    }
    const routes: [string, TargetHop][] = [];
    for (const method of parent.members.methods.values()) {
      const returned = method.returns.trim();
      if (method.arity === 1 && method.name === lowerFirst(returned)) {
        routes.push([returned, { method: method.name, key: method.name }]);
      }
    }
    for (const [property, type] of parent.members.properties) {
      routes.push([type.trim(), { property }]);
    }
    for (const [name, hop] of routes) {
      if (found.has(name)) {
        continue;
      }
      const members = withInherited(name, parse);
      if (!members || members.methods.size === 0) {
        continue;
      }
      const keys = "key" in hop ? [...parent.kind.keys, hop.key] : [...parent.kind.keys];
      const identity = [...members.properties]
        .filter(([, type]) => /^(number|string)$/.test(type.trim()))
        .map(([property]) => property);
      found.set(name, {
        kind: { kind: name, keys, hops: [...parent.kind.hops, hop], identity },
        members,
      });
      queue.push(name);
    }
  }
  return found;
}

function getParsedCatalog(): ParsedCatalog {
  if (cached) {
    return cached;
  }
  const { version, typesPath } = resolveEngineDeclarations();
  const source = readFileSync(typesPath, "utf8");
  const declarations = parseExportedDeclarations(source);
  const parse = (name: string) => parseInterface(source, name, declarations);
  const kinds = discoverKinds(parse);
  const module = parse(MODULE_KIND);
  if (!module) {
    throw new Error(
      "installed @libraz/formulon declaration file has no exported Formulon interface",
    );
  }
  kinds.set(MODULE_KIND, {
    kind: { kind: MODULE_KIND, keys: [], hops: [], identity: [] },
    members: module,
  });
  const accessors = new Set(
    [...kinds.values()].flatMap(({ kind }) =>
      kind.hops.flatMap((hop) => ("method" in hop ? [hop.method] : [])),
    ),
  );
  const methods: WorkbookApiMethod[] = [];
  const parsedMethods = new Map<string, ParsedMethod>();
  const unclassified: string[] = [];
  for (const { kind, members } of kinds.values()) {
    const target: WorkbookApiTarget = { kind: kind.kind, keys: kind.keys };
    for (const method of members.methods.values()) {
      if (SESSION_OWNED_METHODS.test(method.name) || UNCALLABLE_SIGNATURE.test(method.signature)) {
        continue;
      }
      // A module function never changes workbook state, whatever its verb (`setLogMinLevel` sets
      // process-level logging), so it never marks a session dirty.
      const access =
        kind.kind === MODULE_KIND
          ? "read"
          : (methodAccess(method.name, accessors) ??
            (kinds.has(method.returns.trim()) ? "read" : undefined));
      if (access === undefined) {
        unclassified.push(`${kind.kind}.${method.name}`);
        continue;
      }
      methods.push({
        name: method.name,
        target,
        signature: method.signature,
        description: method.description,
        access,
      });
      parsedMethods.set(`${kind.kind}.${method.name}`, method);
    }
  }
  if (unclassified.length > 0) {
    throw new Error(
      `installed @libraz/formulon declares methods neither the ledger nor a name prefix classifies as read or write: ${unclassified.join(", ")}`,
    );
  }
  cached = {
    catalog: { version, typesPath, kinds: [...kinds.values()].map(({ kind }) => kind), methods },
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
          method.target.kind.toLowerCase().includes(needle) ||
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

/**
 * Describes one method, named `Kind.method` (`Sheet.getValue`) or, when only one kind has it,
 * by its bare name.
 */
export function describeWorkbookApi(methodName: string): WorkbookApiDescribeResult {
  const { catalog, declarations, methods } = getParsedCatalog();
  const handleKinds = new Set(catalog.kinds.map((kind) => kind.kind));
  const dot = methodName.indexOf(".");
  const candidates = catalog.methods.filter((entry) =>
    dot < 0
      ? entry.name === methodName
      : entry.target.kind === methodName.slice(0, dot) && entry.name === methodName.slice(dot + 1),
  );
  if (candidates.length === 0) {
    throw new Error(`unknown or unavailable Formulon method: ${methodName}`);
  }
  if (candidates.length > 1) {
    throw new Error(
      `${methodName} exists on several targets; name one of ${candidates.map((entry) => `${entry.target.kind}.${entry.name}`).join(", ")}`,
    );
  }
  const method = candidates[0];

  const declarationsOut: WorkbookApiDeclaration[] = [];
  const seen = new Set<string>();
  const queue: Array<{ name: string; depth: number }> = [];
  for (const name of methods.get(`${method.target.kind}.${method.name}`)?.references ?? []) {
    if (!handleKinds.has(name)) {
      queue.push({ name, depth: 0 });
    }
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
        // A handle is a target of its own, described by its methods; following it would pull in the whole API.
        if (!(seen.has(name) || handleKinds.has(name))) {
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
