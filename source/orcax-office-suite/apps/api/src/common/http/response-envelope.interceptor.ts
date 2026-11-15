import { CallHandler, ExecutionContext, Injectable, NestInterceptor } from '@nestjs/common';
import { API_VERSION, CORRELATION_ID_HEADER, success } from '@orcax/api-contracts';
import { Request } from 'express';
import { Observable, map } from 'rxjs';
import { resolveCorrelationId } from '../../observability/correlation';
import { isApiEnvelope } from './is-api-envelope';

@Injectable()
export class ResponseEnvelopeInterceptor implements NestInterceptor {
  intercept(context: ExecutionContext, next: CallHandler): Observable<unknown> {
    const request = context.switchToHttp().getRequest<Request & { id?: string }>();
    const path = (request.path ?? request.url ?? '').split('?')[0];
    if (path === '/health' || path.startsWith('/health/')) {
      return next.handle();
    }

    const header = request.headers[CORRELATION_ID_HEADER];
    const incoming = Array.isArray(header) ? header[0] : header;
    const correlationId = resolveCorrelationId(incoming, request.id);
    return next.handle().pipe(
      map((data: unknown) => {
        if (isApiEnvelope(data)) {
          return data;
        }
        return success(data ?? null, { correlationId, apiVersion: API_VERSION });
      }),
    );
  }
}
