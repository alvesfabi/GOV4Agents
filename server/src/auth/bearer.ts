import type { Request, Response, NextFunction } from 'express';

export interface AuthedRequest extends Request {
  userAccessToken?: string;
}

export function requireBearer(req: AuthedRequest, res: Response, next: NextFunction): void {
  const header = req.headers.authorization;
  if (!header?.startsWith('Bearer ')) {
    res.status(401).json({ error: 'Missing bearer token' });
    return;
  }
  req.userAccessToken = header.substring('Bearer '.length).trim();
  next();
}
