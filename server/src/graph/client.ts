import { acquireGraphTokenOnBehalfOf } from '../auth/msal.js';

const GRAPH_BASE = 'https://graph.microsoft.com';

export class GraphError extends Error {
  status: number;
  body: unknown;
  constructor(status: number, body: unknown, message: string) {
    super(message);
    this.status = status;
    this.body = body;
  }
}

export interface GraphCallOptions {
  method?: 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';
  body?: unknown;
  version?: 'v1.0' | 'beta';
  /**
   * Some endpoints (e.g. servicePrincipal creation) are only consistent after
   * a brief delay. Allow callers to disable error-throwing on specific codes.
   */
  expectStatuses?: number[];
}

/**
 * High-level Graph client built on top of OBO tokens. Each method takes the
 * caller's SPA access token, exchanges it for a Graph token with the requested
 * scopes, then makes the call.
 */
export class GraphClient {
  constructor(private userAccessToken: string) {}

  async call<T = unknown>(
    scopes: string[],
    path: string,
    options: GraphCallOptions = {},
  ): Promise<T> {
    const { method = 'GET', body, version = 'v1.0', expectStatuses } = options;
    const tokenResult = await acquireGraphTokenOnBehalfOf(this.userAccessToken, scopes);
    const url = path.startsWith('http') ? path : `${GRAPH_BASE}/${version}${path}`;
    const res = await fetch(url, {
      method,
      headers: {
        authorization: `Bearer ${tokenResult.accessToken}`,
        'content-type': 'application/json',
      },
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    const text = await res.text();
    const parsed = text ? safeJson(text) : undefined;
    if (!res.ok && !expectStatuses?.includes(res.status)) {
      const wwwAuth = res.headers.get('www-authenticate');
      if (wwwAuth) console.log(`[Graph ${method} ${path}] WWW-Authenticate:`, wwwAuth);
      throw new GraphError(res.status, parsed, `Graph ${method} ${path} → ${res.status}`);
    }
    return parsed as T;
  }
}

function safeJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}
