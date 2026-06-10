import { useApiToken } from '../auth/useApiToken';
import { getSessionId } from '../auth/sessionId';
import { syncContextToLocal } from '../auth/sessionContext';

export class ApiError extends Error {
  status: number;
  body: unknown;
  constructor(status: number, body: unknown, message: string) {
    super(message);
    this.status = status;
    this.body = body;
  }
}

/**
 * Hook returning a typed `apiFetch` that automatically attaches the API
 * bearer token and the session id header.
 */
export function useApi() {
  const getToken = useApiToken();
  return async function apiFetch<T = unknown>(
    path: string,
    init: RequestInit = {},
  ): Promise<T> {
    const token = await getToken();
    const res = await fetch(path, {
      ...init,
      headers: {
        'content-type': 'application/json',
        authorization: `Bearer ${token}`,
        'x-gov4-session': getSessionId(),
        ...(init.headers ?? {}),
      },
    });
    const text = await res.text();
    const body = text ? safeJson(text) : undefined;
    if (!res.ok) {
      throw new ApiError(res.status, body, `API ${res.status}: ${res.statusText}`);
    }
    // After a successful setup mutation, snapshot the (non-secret) context to
    // localStorage so the demo can be resumed after a server restart.
    const method = (init.method ?? 'GET').toUpperCase();
    if (
      method === 'POST' &&
      path.startsWith('/api/setup/') &&
      !path.endsWith('/manual-session')
    ) {
      void syncContextToLocal(token);
    }
    return body as T;
  };
}

function safeJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}
