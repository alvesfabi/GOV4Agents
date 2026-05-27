/**
 * Stable per-browser-session id used to key server-side session storage
 * (e.g. the agent's client secret). Cleared when sessionStorage is cleared.
 */
const KEY = 'gov4.sessionId';

export function getSessionId(): string {
  let id = sessionStorage.getItem(KEY);
  if (!id) {
    id = crypto.randomUUID();
    sessionStorage.setItem(KEY, id);
  }
  return id;
}
