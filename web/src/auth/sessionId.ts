/**
 * Stable per-browser session id used to key server-side session storage
 * (e.g. the agent's client secret) and the browser-side context snapshot.
 *
 * Stored in localStorage so it survives closing/reopening the tab or browser,
 * letting the user resume the demo where they left off. Any value previously
 * kept in sessionStorage is migrated transparently.
 */
const KEY = 'gov4.sessionId';

export function getSessionId(): string {
  let id = localStorage.getItem(KEY);
  if (!id) {
    const legacy = sessionStorage.getItem(KEY);
    id = legacy ?? crypto.randomUUID();
    localStorage.setItem(KEY, id);
  }
  return id;
}
