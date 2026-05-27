/**
 * In-memory session store for PoC artifacts:
 *  - the Agent's client secret (returned only once at creation)
 *  - the IDs of every component created during setup (so later journey
 *    pages can deep-link to them and act on them).
 *
 * Keyed by the SPA-generated session id (sent as `x-gov4-session` header).
 * Cleared whenever the server restarts — by design.
 */

export interface SessionData {
  blueprint?: {
    id: string;
    name: string;
    appId?: string;
    principalId?: string;
    clientSecret?: string;
  };
  agent?: {
    id: string;
    clientId: string;
    name: string;
    clientSecret?: string;
    tenantId?: string;
    sponsorId?: string;
    sponsorUpn?: string;
  };
  catalog?: { id: string; name: string };
  accessPackage?: { id: string; name: string; approverUpn?: string };
  accessPackage2?: { id: string; name: string; approverUpn?: string };
  lcw?: { id: string; name: string };
  csa?: { setName: string; attributeName: string };
  caPolicy?: { id: string; displayName: string };
}

const store = new Map<string, SessionData>();

export function getSession(id: string): SessionData {
  let s = store.get(id);
  if (!s) {
    s = {};
    store.set(id, s);
  }
  return s;
}

export function setSession(id: string, patch: Partial<SessionData>): SessionData {
  const s = getSession(id);
  Object.assign(s, patch);
  return s;
}

export function clearSession(id: string): void {
  store.delete(id);
}
