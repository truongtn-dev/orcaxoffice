import { createRequire } from 'module';
import { Module } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { LoggerModule } from 'nestjs-pino';
import { CORRELATION_ID_HEADER } from '@orcax/api-contracts';
import { resolveCorrelationId } from './correlation';

const nodeRequire = createRequire(__filename);

function prettyTransport(): { target: string; options: { singleLine: boolean } } | undefined {
  try {
    nodeRequire.resolve('pino-pretty');
    return { target: 'pino-pretty', options: { singleLine: true } };
  } catch {
    return undefined;
  }
}

@Module({
  imports: [
    LoggerModule.forRootAsync({
      inject: [ConfigService],
      useFactory: (config: ConfigService) => ({
        pinoHttp: {
          level: config.getOrThrow<string>('logLevel'),
          genReqId: (req, res) => {
            const header = req.headers[CORRELATION_ID_HEADER];
            const incoming = Array.isArray(header) ? header[0] : header;
            const existing = typeof req.id === 'string' ? req.id : undefined;
            const correlationId = resolveCorrelationId(incoming, existing);
            res.setHeader(CORRELATION_ID_HEADER, correlationId);
            return correlationId;
          },
          redact: {
            paths: [
              'req.headers.authorization',
              'req.headers.cookie',
              'password',
              'token',
              'secret',
              'accessToken',
              'refreshToken',
            ],
            censor: '[redacted]',
          },
          serializers: {
            req(req) {
              const url = typeof req.url === 'string' ? req.url.split('?')[0] : req.url;
              return { id: req.id, method: req.method, url };
            },
            res(res) {
              return { statusCode: res.statusCode };
            },
          },
          autoLogging: {
            ignore: (req) => (req.url ?? '').split('?')[0] === '/health',
          },
          transport: config.get('nodeEnv') === 'development' ? prettyTransport() : undefined,
        },
      }),
    }),
  ],
})
export class ObservabilityModule {}
