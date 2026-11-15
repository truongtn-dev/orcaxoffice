import { randomUUID } from 'crypto';
import { NextFunction, Request, Response } from 'express';
import { CORRELATION_ID_HEADER } from '@orcax/api-contracts';
import { isUuid } from '../common/ids/uuid';

export function resolveCorrelationId(incoming?: string, existing?: string): string {
  if (existing && isUuid(existing)) {
    return existing;
  }
  if (incoming && isUuid(incoming)) {
    return incoming;
  }
  return randomUUID();
}

export function correlationMiddleware(req: Request, res: Response, next: NextFunction): void {
  const header = req.header(CORRELATION_ID_HEADER);
  const current = (req as Request & { id?: string }).id;
  const correlationId = resolveCorrelationId(header, current);
  (req as Request & { id?: string }).id = correlationId;
  res.setHeader(CORRELATION_ID_HEADER, correlationId);
  next();
}
