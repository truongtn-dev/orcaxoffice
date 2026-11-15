import { INestApplication, RequestMethod, ValidationPipe } from '@nestjs/common';
import { API_VERSION } from '@orcax/api-contracts';
import { ApiExceptionFilter } from '../common/http/api-exception.filter';
import { ResponseEnvelopeInterceptor } from '../common/http/response-envelope.interceptor';
import { correlationMiddleware } from '../observability/correlation';

export function configureHttp(app: INestApplication, options?: { webOrigin?: string }): void {
  const server = app.getHttpAdapter().getInstance() as { disable?: (setting: string) => void };
  server.disable?.('x-powered-by');
  app.use(correlationMiddleware);
  app.setGlobalPrefix(`api/${API_VERSION}`, {
    exclude: [
      { path: 'health', method: RequestMethod.GET },
      { path: 'health/ready', method: RequestMethod.GET },
    ],
  });
  app.useGlobalPipes(
    new ValidationPipe({
      whitelist: true,
      forbidNonWhitelisted: true,
      transform: true,
    }),
  );
  app.useGlobalFilters(new ApiExceptionFilter());
  app.useGlobalInterceptors(new ResponseEnvelopeInterceptor());
  const origin = options?.webOrigin?.trim();
  if (origin) {
    app.enableCors({ origin, credentials: true });
  }
  app.enableShutdownHooks();
}
