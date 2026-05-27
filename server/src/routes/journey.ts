import { Router } from 'express';
import { requireBearer, type AuthedRequest } from '../auth/bearer.js';
import { attachSession, type SessionedRequest } from '../store/middleware.js';
import { acquireAgentTokenViaFic } from '../auth/msal.js';
import { GraphClient, GraphError } from '../graph/client.js';
import { GraphScopes } from '../graph/constants.js';
import { env } from '../config.js';

export const agentRouter = Router();
agentRouter.use(requireBearer, attachSession);

type Req = AuthedRequest & SessionedRequest;

/**
 * Two-step Agent ID authentication:
 *  1. Blueprint client-credentials (using the blueprint's secret) →
 *     api://AzureADTokenExchange token.
 *  2. That token is presented as a JWT-bearer client_assertion against the
 *     agent's appId, yielding the agent's Graph access token. The blueprint's
 *     secret never leaves the backend; the SPA only sees the resulting agent
 *     token (which the user is meant to inspect/decode).
 */
agentRouter.post('/token', async (req: Req, res) => {
  try {
    const blueprint = req.sessionData?.blueprint;
    const agent = req.sessionData?.agent;
    if (!blueprint?.appId || !blueprint.clientSecret) {
      return res
        .status(400)
        .json({ error: 'blueprint or its client secret missing in session' });
    }
    if (!agent?.clientId) {
      return res.status(400).json({ error: 'agent not created yet' });
    }
    const { agent: agentTok } = await acquireAgentTokenViaFic({
      tenantId: agent.tenantId ?? env.tenantId,
      blueprintAppId: blueprint.appId,
      blueprintSecret: blueprint.clientSecret,
      agentAppId: agent.clientId,
    });
    try {
      const p = JSON.parse(
        Buffer.from(String(agentTok.access_token).split('.')[1], 'base64').toString('utf8'),
      );
      console.log('[AGENT TOKEN] scp:', p.scp, 'roles:', p.roles, 'aud:', p.aud, 'appid:', p.appid, 'idtyp:', p.idtyp);
    } catch {
      /* ignore */
    }
    return res.json({
      accessToken: agentTok.access_token,
      expiresIn: agentTok.expires_in,
      tokenType: agentTok.token_type,
      scope: agentTok.scope,
    });
  } catch (err) {
    return failure(res, err);
  }
});

/**
 * Proxies an arbitrary Microsoft Graph call as the agent. Acquires the agent
 * token via the same FIC exchange, then calls Graph with it.
 */
agentRouter.post('/graph', async (req: Req, res) => {
  try {
    const blueprint = req.sessionData?.blueprint;
    const agent = req.sessionData?.agent;
    if (!blueprint?.appId || !blueprint.clientSecret) {
      return res
        .status(400)
        .json({ error: 'blueprint or its client secret missing in session' });
    }
    if (!agent?.clientId) {
      return res.status(400).json({ error: 'agent not created yet' });
    }
    const path: string = req.body?.path;
    if (!path) return res.status(400).json({ error: 'path required' });
    const requestedScope: string | undefined = req.body?.scope;

    const { agent: agentTok } = await acquireAgentTokenViaFic({
      tenantId: agent.tenantId ?? env.tenantId,
      blueprintAppId: blueprint.appId,
      blueprintSecret: blueprint.clientSecret,
      agentAppId: agent.clientId,
      scope: requestedScope,
    });
    try {
      const p = JSON.parse(
        Buffer.from(String(agentTok.access_token).split('.')[1], 'base64').toString('utf8'),
      );
      console.log('[AGENT GRAPH CALL] path:', path, 'scp:', p.scp, 'roles:', p.roles);
    } catch {
      /* ignore */
    }
    const url = path.startsWith('http')
      ? path
      : `https://graph.microsoft.com/v1.0${path.startsWith('/') ? path : `/${path}`}`;
    const r = await fetch(url, {
      headers: { authorization: `Bearer ${agentTok.access_token}` },
    });
    const text = await r.text();
    let body: unknown = text;
    try {
      body = text ? JSON.parse(text) : undefined;
    } catch {
      /* leave as text */
    }
    return res.status(r.status).json({ status: r.status, ok: r.ok, body });
  } catch (err) {
    return failure(res, err);
  }
});

export const governRouter = Router();
governRouter.use(requireBearer, attachSession);

/**
 * Runs the offboarding lifecycle workflow against the signed-in user. This
 * triggers the agent-sponsorship-transfer task, which is the demo's "sponsor
 * leaves the org" scenario.
 */
governRouter.post('/run-offboarding', async (req: Req, res) => {
  try {
    const lcw = req.sessionData?.lcw;
    if (!lcw) return res.status(400).json({ error: 'lifecycle workflow not created yet' });
    const g = new GraphClient(req.userAccessToken!);

    // Resolve target user(s). Default to the agent's sponsor (the user who
    // was specified during Agent ID creation). Allow override via request
    // body if explicitly provided.
    const subjectIds: string[] = req.body?.subjectIds ?? [];
    if (subjectIds.length === 0) {
      const sponsorId = req.sessionData?.agent?.sponsorId;
      if (sponsorId) {
        subjectIds.push(sponsorId);
      } else {
        const me = await g.call<{ id: string }>(GraphScopes.user, '/me');
        subjectIds.push(me.id);
      }
    }

    const run = await g.call(
      GraphScopes.lcw,
      `/identityGovernance/lifecycleWorkflows/workflows/${lcw.id}/activate`,
      {
        method: 'POST',
        body: {
          subjects: subjectIds.map((id) => ({ id })),
        },
      },
    );
    return res.json({ ok: true, run, subjectIds });
  } catch (err) {
    return failure(res, err);
  }
});

export const protectRouter = Router();
protectRouter.use(requireBearer, attachSession);

/**
 * Patches the Conditional Access policy state to fully enabled.
 */
protectRouter.post('/enable-ca', async (req: Req, res) => {
  try {
    const ca = req.sessionData?.caPolicy;
    if (!ca) return res.status(400).json({ error: 'CA policy not created yet' });
    const g = new GraphClient(req.userAccessToken!);
    await g.call(
      GraphScopes.conditionalAccess,
      `/identity/conditionalAccess/policies/${ca.id}`,
      { method: 'PATCH', body: { state: 'enabled' } },
    );
    return res.json({ ok: true, state: 'enabled', policyId: ca.id });
  } catch (err) {
    return failure(res, err);
  }
});

/**
 * Sets the agent's CSA value (AgentsCSA/TAG = approved) so it satisfies the
 * Conditional Access filter and can authenticate again.
 */
protectRouter.post('/set-csa', async (req: Req, res) => {
  try {
    const agent = req.sessionData?.agent;
    const csa = req.sessionData?.csa;
    if (!agent || !csa) {
      return res.status(400).json({ error: 'agent or CSA not created yet' });
    }
    const value: string = req.body?.value ?? 'approved';
    const g = new GraphClient(req.userAccessToken!);
    await g.call(
      GraphScopes.csaAssignment.concat(GraphScopes.directory),
      `/servicePrincipals/${agent.id}`,
      {
        method: 'PATCH',
        body: {
          customSecurityAttributes: {
            [csa.setName]: {
              '@odata.type': '#Microsoft.DirectoryServices.CustomSecurityAttributeValue',
              [`${csa.attributeName}@odata.type`]: '#String',
              [csa.attributeName]: value,
            },
          },
        },
      },
    );
    return res.json({ ok: true, [csa.attributeName]: value });
  } catch (err) {
    return failure(res, err);
  }
});

function failure(res: import('express').Response, err: unknown): void {
  if (err instanceof GraphError) {
    res.status(err.status).json({ error: err.message, body: err.body });
    return;
  }
  if (err instanceof Error) {
    res.status(500).json({ error: err.message });
    return;
  }
  res.status(500).json({ error: 'unknown error' });
}
