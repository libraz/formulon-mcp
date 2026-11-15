import {
  loadedFormulon,
  resultToJson,
  StatusCode,
  statusError,
  type Workbook,
} from "../formulon.js";
import { getWorkbookApiCatalog, MODULE_KIND } from "../workbook-api.js";
import { getSession, publicInfo, touch } from "./state.js";

/**
 * The object a `formulon_workbook_call` addresses: a handle kind of the installed engine
 * (`Workbook`, `Sheet`, `Table`, `Pivot`, `Chart`, ...) and the key each accessor on the way
 * takes (`{ kind: "Pivot", sheet: 0, pivot: "Sales" }`). `formulon_workbook_api` lists every kind
 * with its keys.
 */
export type WorkbookCallTarget = { kind: string } & Record<string, unknown>;

/**
 * Replaces each engine handle in a result with its kind, its identifying scalars (`cacheId`,
 * `objectId`, `index`, `name`) and, when it has one, its properties record.
 */
function describeHandles(value: unknown, kinds: ReadonlyMap<string, readonly string[]>): unknown {
  if (Array.isArray(value)) {
    return value.map((item) => describeHandles(item, kinds));
  }
  if (typeof value === "object" && value !== null) {
    const kind = (value as { constructor?: { name?: string } }).constructor?.name;
    const identity = kind === undefined ? undefined : kinds.get(kind);
    if (kind !== undefined && identity !== undefined) {
      const handle = value as Record<string, unknown>;
      const getProperties = handle.getProperties;
      return {
        handle: kind,
        ...Object.fromEntries(identity.map((property) => [property, handle[property]])),
        ...(typeof getProperties === "function" ? { properties: getProperties.call(value) } : {}),
      };
    }
  }
  return value;
}

/** Calls a method of the installed engine on the handle `target` names, with positional JSON arguments. */
export function callWorkbookMethod(
  id: string,
  target: WorkbookCallTarget | undefined,
  method: string,
  args: unknown[],
) {
  const catalog = getWorkbookApiCatalog();
  const kindName = target?.kind ?? "Workbook";
  const kind = catalog.kinds.find((entry) => entry.kind === kindName);
  if (!kind) {
    throw statusError(
      StatusCode.NotFound,
      `unknown target kind: ${kindName}; one of ${catalog.kinds.map((entry) => entry.kind).join(", ")}`,
    );
  }
  const entry = catalog.methods.find(
    (candidate) => candidate.target.kind === kind.kind && candidate.name === method,
  );
  if (!entry) {
    throw statusError(
      StatusCode.NotFound,
      `${kind.kind}.${method} is not a callable method (see formulon_workbook_api)`,
    );
  }
  const session = getSession(id);
  let receiver: unknown = kind.kind === MODULE_KIND ? loadedFormulon() : session.workbook;
  for (const hop of kind.hops) {
    const current = receiver as Record<string, unknown>;
    if ("property" in hop) {
      receiver = current[hop.property];
      continue;
    }
    const key = target?.[hop.key];
    if (key === undefined) {
      throw statusError(StatusCode.InvalidArgument, `target ${kind.kind} needs the key ${hop.key}`);
    }
    receiver = (current[hop.method] as (key: unknown) => unknown).call(current, key);
  }
  const dirty = entry.access === "write";
  // Mutating calls can fail after partially changing native state, so dirty is
  // recorded before invocation and retained when the call throws.
  if (dirty) {
    touch(session, true);
  }
  const callable = (receiver as Record<string, unknown>)[method] as (
    ...callArgs: unknown[]
  ) => unknown;
  const result = callable.apply(receiver, args);
  touch(session, dirty);
  const handleKinds = new Map(catalog.kinds.map((entry) => [entry.kind, entry.identity]));
  return {
    session: publicInfo(session),
    target: { ...target, kind: kind.kind },
    method,
    result: resultToJson(describeHandles(result, handleKinds)),
  };
}

/**
 * Runs one engine operation of a task tool on the session's workbook and reports it as
 * `{ session, method, result }`, where `method` names the operation. A `write` operation marks
 * the session dirty once it succeeds: each one is a single engine call, which changes nothing
 * when it throws.
 */
export function callLabeled(
  id: string,
  method: string,
  access: "read" | "write",
  operation: (workbook: Workbook) => unknown,
) {
  const session = getSession(id);
  const result = operation(session.workbook);
  touch(session, access === "write");
  return { session: publicInfo(session), method, result: resultToJson(result ?? null) };
}
