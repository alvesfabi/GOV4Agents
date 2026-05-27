import { ConfidentialClientApplication, type AuthenticationResult } from '@azure/msal-node';
import { env } from '../config.js';

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: env.clientId,
    clientSecret: env.clientSecret,
    authority: `https://login.microsoftonline.com/${env.tenantId}`,
  },
});

/**
 * Acquire a Microsoft Graph token for THIS app using client-credentials.
 * Used for endpoints (like Conditional Access policy creation) where Graph
 * rejects the OBO delegated token even when the user is a Global Admin and
 * the delegated scope is consented. The app reg must have the corresponding
 * application permissions granted.
 */
export async function acquireGraphTokenAsApp(): Promise<AuthenticationResult> {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'],
  });
  if (!result) throw new Error('App client-credentials request returned no result');
  try {
    const payload = JSON.parse(
      Buffer.from(result.accessToken.split('.')[1], 'base64').toString('utf8'),
    );
    console.log('[App-Graph] roles:', payload.roles, 'aud:', payload.aud, 'appid:', payload.appid, 'idtyp:', payload.idtyp);
  } catch {
    /* ignore */
  }
  return result;
}

/**
 * Exchange the SPA's access token (audience = our API) for a Microsoft Graph
 * token using the on-behalf-of flow. Scopes are Graph delegated scopes.
 */
export async function acquireGraphTokenOnBehalfOf(
  userAccessToken: string,
  _scopes: string[],
): Promise<AuthenticationResult> {
  // We request `.default` instead of named scopes so Entra issues a Graph token
  // with whatever delegated permissions admin has already consented to on the
  // app registration. This avoids AADSTS65001 when individual preview scopes
  // (e.g. AgentIdentityBlueprint.*) are present but not yet consentable per
  // scope. The Graph API itself enforces the actual permission requirement.
  const result = await cca.acquireTokenOnBehalfOf({
    oboAssertion: userAccessToken,
    scopes: ['https://graph.microsoft.com/.default'],
  });
  if (!result) throw new Error('OBO exchange returned no result');
  try {
    const payload = JSON.parse(
      Buffer.from(result.accessToken.split('.')[1], 'base64').toString('utf8'),
    );
    console.log('[OBO] Graph token scp:', payload.scp);
    console.log('[OBO] amr:', payload.amr, 'acrs:', payload.acrs, 'acr:', payload.acr, 'wids count:', payload.wids?.length);
    console.log('[OBO] aud:', payload.aud, 'iss:', payload.iss, 'tid:', payload.tid, 'appid:', payload.appid, 'idtyp:', payload.idtyp, 'wids:', payload.wids);
  } catch {
    /* ignore */
  }
  return result;
}

/**
 * Acquire a token for an arbitrary client app using client-credentials.
 * Used to authenticate as the **created Agent ID** during the Govern journey.
 */
export async function acquireAgentToken(params: {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  scopes?: string[];
}): Promise<AuthenticationResult> {
  const agent = new ConfidentialClientApplication({
    auth: {
      clientId: params.clientId,
      clientSecret: params.clientSecret,
      authority: `https://login.microsoftonline.com/${params.tenantId}`,
    },
  });
  const result = await agent.acquireTokenByClientCredential({
    scopes: params.scopes ?? ['https://graph.microsoft.com/.default'],
  });
  if (!result) throw new Error('Client-credentials request returned no result');
  return result;
}

/**
 * Two-step authentication flow for an Agent ID:
 *
 * 1. Acquire a token for the **blueprint application** using client-credentials
 *    (its secret), with audience = api://AzureADTokenExchange. This yields a
 *    JWT we can use as a federated identity assertion.
 * 2. Use that assertion as `client_assertion` against the **agent's** appId in
 *    a client_credentials request, producing a token issued to the agent.
 *
 * Returns the raw token response from Entra (access_token / expires_in / etc).
 */
export async function acquireAgentTokenViaFic(params: {
  tenantId: string;
  blueprintAppId: string;
  blueprintSecret: string;
  agentAppId: string;
  scope?: string;
}): Promise<{
  blueprint: Record<string, unknown>;
  agent: Record<string, unknown>;
}> {
  const tokenUrl = `https://login.microsoftonline.com/${params.tenantId}/oauth2/v2.0/token`;

  const bpForm = new URLSearchParams({
    client_id: params.blueprintAppId,
    client_secret: params.blueprintSecret,
    scope: 'api://AzureADTokenExchange/.default',
    grant_type: 'client_credentials',
    fmi_path: params.agentAppId,
    fmipath: params.agentAppId,
  });
  const bpRes = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: bpForm.toString(),
  });
  const bpJson = (await bpRes.json()) as Record<string, unknown>;
  if (!bpRes.ok || !bpJson.access_token) {
    throw new Error(
      `Blueprint client-credentials failed: ${JSON.stringify(bpJson)}`,
    );
  }

  const agentForm = new URLSearchParams({
    client_id: params.agentAppId,
    client_assertion_type:
      'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
    client_assertion: String(bpJson.access_token),
    scope: params.scope ?? 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const agentRes = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: agentForm.toString(),
  });
  const agentJson = (await agentRes.json()) as Record<string, unknown>;
  if (!agentRes.ok || !agentJson.access_token) {
    throw new Error(
      `Agent FIC exchange failed: ${JSON.stringify(agentJson)}`,
    );
  }

  return { blueprint: bpJson, agent: agentJson };
}
