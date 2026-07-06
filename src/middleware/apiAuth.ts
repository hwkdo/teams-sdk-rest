import type { NextFunction, Request, Response } from 'express';
import type { Config } from '../config.js';

export function createApiAuth(config: Config) {
  return (req: Request, res: Response, next: NextFunction): void => {
    const header = req.headers.authorization;

    if (!header?.startsWith('Bearer ')) {
      res.status(401).json({ success: false, error: 'Missing or invalid Authorization header' });
      return;
    }

    const token = header.slice('Bearer '.length);

    if (token !== config.API_KEY) {
      res.status(403).json({ success: false, error: 'Invalid API key' });
      return;
    }

    next();
  };
}
