import 'reflect-metadata';
import { NestFactory } from '@nestjs/core';
import { ConfigService } from '@nestjs/config';
import { Logger } from 'nestjs-pino';
import { AppModule } from './app.module';
import { configureHttp } from './bootstrap/configure-http';

async function bootstrap(): Promise<void> {
  const app = await NestFactory.create(AppModule, { bufferLogs: true });
  const logger = app.get(Logger);
  app.useLogger(logger);
  const config = app.get(ConfigService);
  configureHttp(app, { webOrigin: config.get<string>('webOrigin') });
  const port = config.getOrThrow<number>('apiPort');
  await app.listen(port);
  logger.log(`API listening on port ${port}`);
}

bootstrap().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : 'API failed to start';
  process.stderr.write(`${message}\n`);
  process.exit(1);
});
