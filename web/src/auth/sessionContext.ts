import { getSessionId } from './sessionId';

/**
 * Browser-side persistence of the (non-secret) demo context so the user can
 * reopen the app and resume where they left off — even if the API server was
 * restarted and lost its in-memory session.
 *
 * The snapshot mirrors the server's SessionData minus client secrets, which by
 * design never leave the backend. Resources are recreated identically on
 * rehydrate via the existing /api/setup/manual-session endpoint.
 */

export interface PersistedContext {
  demoSuffix?: string;
  blueprint?: unknown;
  agent?: unknown;
  catalog?: unknown;
  accessPackage?: unknown;
  accessPackage2?: unknown;
  lcw?: unknown;
  csa?: unknown;
  caPolicy?: unknown;
}

const PREFIX = 'gov4.context.';

function storageKey(): string {
  return PREFIX + getSessionId();
}

export function loadContext(): PersistedContext {
  try {
    return JSON.parse(localStorage.getItem(storageKey()) ?? '{}') as PersistedContext;
  } catch {
    return {};
  }
}

/**
 * Merge a snapshot into the stored context. Null/undefined fields never
 * overwrite previously stored values, so a momentarily empty server response
 * (e.g. right after a restart) cannot wipe a good snapshot.
 */
export function saveContext(patch: PersistedContext): void {
  const merged: Record<string, unknown> = { ...loadContext() };
  for (const [k, v] of Object.entries(patch)) {
    if (v !== null && v !== undefined) merged[k] = v;
  }
  localStorage.setItem(storageKey(), JSON.stringify(merged));
}

function headers(token: string): Record<string, string> {
  return {
    'content-type': 'application/json',
    authorization: `Bearer ${token}`,
    'x-gov4-session': getSessionId(),
  };
}

/** Pull the server's non-secret session snapshot and persist it locally. */
export async function syncContextToLocal(token: string): Promise<void> {
  try {
    const res = await fetch('/api/setup/session-context', { headers: headers(token) });
    if (!res.ok) return;
    const data = (await res.json()) as { context?: PersistedContext };
    if (data?.context) saveContext(data.context);
  } catch {
    /* best-effort */
  }
}

/** Push the locally persisted context back into the server session. */
export async function rehydrateServer(token: string): Promise<void> {
  const ctx = loadContext();
  if (!ctx || Object.keys(ctx).length === 0) return;
  try {
    await fetch('/api/setup/manual-session', {
      method: 'POST',
      headers: headers(token),
      body: JSON.stringify(ctx),
    });
  } catch {
    /* best-effort */
  }
}
