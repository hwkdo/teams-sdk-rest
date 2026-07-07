import { inspect } from 'node:util';
import type { NextFunction, Request, Response } from 'express';

type UpstreamError = {
  status?: number;
  statusCode?: number;
  response?: {
    status?: number;
    data?: unknown;
  };
  message?: string;
};

function getUpstreamStatus(err: UpstreamError): number | undefined {
  return err.response?.status ?? err.status ?? err.statusCode;
}

function getUpstreamData(err: UpstreamError): unknown {
  return err.response?.data;
}

export function errorHandler(err: unknown, _req: Request, res: Response, _next: NextFunction): void {
  console.error(inspect(err, { depth: null, colors: false }));

  const upstream = (typeof err === 'object' && err !== null ? err : {}) as UpstreamError;
  const upstreamStatus = getUpstreamStatus(upstream);
  const upstreamData = getUpstreamData(upstream);
  const message = err instanceof Error ? err.message : 'Internal server error';

  let status = 500;
  if (message.includes('not found') || message.includes('No conversation found')) {
    status = 404;
  } else if (upstreamStatus && upstreamStatus >= 400 && upstreamStatus < 600) {
    status = upstreamStatus;
  }

  const body: Record<string, unknown> = {
    success: false,
    error: message,
  };

  if (upstreamStatus) {
    body.upstreamStatus = upstreamStatus;
  }

  if (upstreamData !== undefined) {
    body.details = upstreamData;
  }

  if (upstreamStatus === 403) {
    body.hint =
      'Bot Connector returned 403 Forbidden. The bot app is most likely not installed in the target team/chat. Install the Teams app in that team (or 1:1 chat) before sending.';
  }

  res.status(status).json(body);
}
