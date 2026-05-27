import type { Request, Response, NextFunction } from 'express';
import { getSession, type SessionData } from './session.js';

export interface SessionedRequest extends Request {
  sessionId?: string;
  sessionData?: SessionData;
}

const HEADER = 'x-gov4-session';

export function attachSession(
  req: SessionedRequest,
  res: Response,
  next: NextFunction,
): void {
  const id = req.header(HEADER);
  if (!id) {
    res.status(400).json({ error: `Missing ${HEADER} header` });
    return;
  }
  req.sessionId = id;
  req.sessionData = getSession(id);
  next();
}
