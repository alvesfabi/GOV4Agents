import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import { env } from './config.js';
import { requireBearer, type AuthedRequest } from './auth/bearer.js';
import { attachSession, type SessionedRequest } from './store/middleware.js';
import { acquireGraphTokenOnBehalfOf } from './auth/msal.js';
import { setupRouter } from './routes/setup.js';
import { agentRouter, governRouter, protectRouter } from './routes/journey.js';

const app = express();

app.use(
  cors({
    origin: env.spaOrigin,
    credentials: true,
    allowedHeaders: ['authorization', 'content-type', 'x-gov4-session'],
  }),
);
app.use(express.json({ limit: '1mb' }));

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, service: 'gov4agents-server' });
});

/**
 * Diagnostic: exchanges the caller's SPA token via OBO for a Graph token using
 * the requested scope(s). Returns only metadata (not the token) to keep this
 * safe to call from the browser during development.
 */
app.post(
  '/api/auth/probe',
  requireBearer,
  attachSession,
  async (req: AuthedRequest & SessionedRequest, res) => {
    try {
      const scopes: string[] = req.body?.scopes ?? ['User.Read'];
      const result = await acquireGraphTokenOnBehalfOf(req.userAccessToken!, scopes);
      res.json({
        ok: true,
        scopes: result.scopes,
        expiresOn: result.expiresOn,
        account: result.account?.username ?? null,
        tenantId: result.tenantId,
      });
    } catch (err) {
      res.status(400).json({ ok: false, error: serializeError(err) });
    }
  },
);

app.use('/api/setup', setupRouter);
app.use('/api/agent', agentRouter);
app.use('/api/govern', governRouter);
app.use('/api/protect', protectRouter);

function serializeError(err: unknown): unknown {
  if (err instanceof Error) {
    return { name: err.name, message: err.message };
  }
  return err;
}

app.listen(env.port, () => {
  console.log(`[gov4agents-server] listening on http://localhost:${env.port}`);
});
