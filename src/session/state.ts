import { randomUUID } from "node:crypto";

import {
  assertStatus,
  createOrLoadWorkbook,
  type LossCounters,
  readLosses,
  saveWorkbook,
  type Workbook,
  workbookSummary,
} from "../formulon.js";
export type SessionInfo = {
  id: string;
  sourcePath?: string;
  outputPath?: string;
  createdAt: string;
  updatedAt: string;
  dirty: boolean;
  /** Counters for what the loader could not decode; absent for a clean load. */
  loadLosses?: LossCounters;
};

export type WorkbookSession = SessionInfo & {
  workbook: Workbook;
};

const sessions = new Map<string, WorkbookSession>();
const revisions = new WeakMap<WorkbookSession, number>();

function nowIso(): string {
  return new Date().toISOString();
}

export function touch(session: WorkbookSession, dirty: boolean): void {
  session.updatedAt = nowIso();
  if (dirty) {
    session.dirty = true;
    revisions.set(session, (revisions.get(session) ?? 0) + 1);
  }
}

function revisionOf(session: WorkbookSession): number {
  return revisions.get(session) ?? 0;
}

export function publicInfo(session: WorkbookSession): SessionInfo {
  return {
    id: session.id,
    sourcePath: session.sourcePath,
    outputPath: session.outputPath,
    createdAt: session.createdAt,
    updatedAt: session.updatedAt,
    dirty: session.dirty,
    loadLosses: session.loadLosses,
  };
}
/** Opens a workbook session from disk, or creates a new default workbook when no path is given. */
export async function openSession(
  inputPath?: string,
  id: string = randomUUID(),
): Promise<SessionInfo> {
  if (sessions.has(id)) {
    throw new Error(`session already exists: ${id}`);
  }
  const workbook = await createOrLoadWorkbook(inputPath);
  let registered = false;
  try {
    // The first check races with another open that awaited the same engine
    // work. Re-check after the await so a duplicate id cannot overwrite the
    // existing handle, and always release this handle on any setup failure.
    if (sessions.has(id)) {
      throw new Error(`session already exists: ${id}`);
    }
    // Excel caches formula results, but the engine leaves formula cells blank
    // until recalculated. Recalc once on open so the first read of a formula
    // cell returns its computed value rather than a stale blank.
    assertStatus(workbook.recalc(), "recalc workbook on open");
    const createdAt = nowIso();
    const session: WorkbookSession = {
      id,
      sourcePath: inputPath,
      createdAt,
      updatedAt: createdAt,
      dirty: false,
      // Surface what the reader could not decode at open time: a workbook that
      // lost parts or formulas on load would otherwise save back silently thinned.
      loadLosses: inputPath ? readLosses(workbook) : undefined,
      workbook,
    };
    sessions.set(id, session);
    revisions.set(session, 0);
    registered = true;
    return publicInfo(session);
  } finally {
    if (!registered && workbook.isValid()) {
      workbook.delete();
    }
  }
}

/** Lists open workbook sessions without exposing native handles. */
export function listSessions(): SessionInfo[] {
  return Array.from(sessions.values(), publicInfo);
}

/** Returns an open session or throws when the id is unknown. */
export function getSession(id: string): WorkbookSession {
  const session = sessions.get(id);
  if (!session) {
    throw new Error(`session not found: ${id}`);
  }
  return session;
}

/** Closes a session and releases the underlying Formulon workbook handle. */
export function closeSession(id: string): SessionInfo {
  const session = getSession(id);
  sessions.delete(id);
  session.workbook.delete();
  return publicInfo(session);
}

/** Returns a workbook summary for an open session. */
export function inspectSession(id: string, includeCells: boolean, maxCellsPerSheet: number) {
  const session = getSession(id);
  touch(session, false);
  return {
    session: publicInfo(session),
    workbook: workbookSummary(session.workbook, includeCells, maxCellsPerSheet),
  };
}
/** Saves an open session to disk. */
export async function saveSession(id: string, outputPath?: string) {
  const session = getSession(id);
  const destination = outputPath ?? session.outputPath ?? session.sourcePath;
  if (!destination) {
    throw new Error("outputPath is required for a new workbook session");
  }
  const saveRevision = revisionOf(session);
  const saved = await saveWorkbook(session.workbook, destination);
  session.outputPath = destination;
  // A lossy save leaves the in-memory workbook different from the artifact on
  // disk. Keep the dirty marker set so callers can choose XLSX or inspect the
  // reported counters before treating the session as safely persisted.
  const lossy = Object.entries(saved.losses ?? {}).some(
    ([key, count]) => key !== "renumberedParts" && count > 0,
  );
  session.dirty = lossy || revisionOf(session) !== saveRevision;
  touch(session, false);
  return {
    session: publicInfo(session),
    outputPath: destination,
    bytes: saved.bytes,
    format: saved.format,
    losses: saved.losses,
  };
}
