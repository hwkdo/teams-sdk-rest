import type { NextFunction, Request, Response } from 'express';

export function errorHandler(err: unknown, _req: Request, res: Response, _next: NextFunction): void {
  console.error(err);

  const message = err instanceof Error ? err.message : 'Internal server error';
  const status = message.includes('not found') ? 404 : 500;

  res.status(status).json({
    success: false,
    error: message,
  });
}
