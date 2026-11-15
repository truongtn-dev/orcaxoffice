import {
  ArgumentsHost,
  Catch,
  ExceptionFilter,
  HttpException,
  HttpStatus,
  Logger,
} from '@nestjs/common';
import { API_VERSION, CORRELATION_ID_HEADER, ErrorCodes, failure } from '@orcax/api-contracts';
import { Request, Response } from 'express';
import { resolveCorrelationId } from '../../observability/correlation';

const STATUS_CODES: Readonly<Record<number, string>> = {
  [HttpStatus.BAD_REQUEST]: ErrorCodes.BAD_REQUEST,
  [HttpStatus.UNAUTHORIZED]: ErrorCodes.UNAUTHORIZED,
  [HttpStatus.FORBIDDEN]: ErrorCodes.FORBIDDEN,
  [HttpStatus.NOT_FOUND]: ErrorCodes.NOT_FOUND,
  [HttpStatus.CONFLICT]: ErrorCodes.CONFLICT,
  [HttpStatus.SERVICE_UNAVAILABLE]: ErrorCodes.SERVICE_UNAVAILABLE,
};

@Catch()
export class ApiExceptionFilter implements ExceptionFilter {
  private readonly logger = new Logger(ApiExceptionFilter.name);

  catch(exception: unknown, host: ArgumentsHost): void {
    const http = host.switchToHttp();
    const response = http.getResponse<Response>();
    const request = http.getRequest<Request & { id?: string }>();
    const header = request.headers[CORRELATION_ID_HEADER];
    const incoming = Array.isArray(header) ? header[0] : header;
    const correlationId = resolveCorrelationId(incoming, request.id);
    response.setHeader(CORRELATION_ID_HEADER, correlationId);

    if (exception instanceof HttpException) {
      const status = exception.getStatus();
      const body = exception.getResponse();
      let code = STATUS_CODES[status] ?? ErrorCodes.INTERNAL_ERROR;
      let message = 'Yêu cầu không thể xử lý.';
      let details: unknown = null;

      if (typeof body === 'string') {
        message = body;
      } else if (typeof body === 'object' && body !== null && 'message' in body) {
        const raw = (body as { message?: unknown }).message;
        if (Array.isArray(raw)) {
          code = ErrorCodes.VALIDATION_FAILED;
          message = 'Dữ liệu không hợp lệ.';
          details = raw.filter((item) => typeof item === 'string').slice(0, 20);
        } else if (typeof raw === 'string') {
          message = raw;
        }
        const explicit = (body as { code?: unknown }).code;
        if (!Array.isArray(raw) && typeof explicit === 'string') {
          code = explicit;
        }
      }

      response
        .status(status)
        .json(failure({ code, message, details }, { correlationId, apiVersion: API_VERSION }));
      return;
    }

    const name = exception instanceof Error ? exception.name : 'Error';
    this.logger.error(`Unhandled exception name=${name} correlationId=${correlationId}`);
    response.status(HttpStatus.INTERNAL_SERVER_ERROR).json(
      failure(
        {
          code: ErrorCodes.INTERNAL_ERROR,
          message: 'Đã xảy ra lỗi nội bộ.',
          details: null,
        },
        { correlationId, apiVersion: API_VERSION },
      ),
    );
  }
}
