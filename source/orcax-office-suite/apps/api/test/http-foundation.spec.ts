import 'reflect-metadata';
import assert from 'node:assert/strict';
import { after, before, describe, it } from 'node:test';
import {
  Body,
  Controller,
  Get,
  HttpException,
  HttpStatus,
  INestApplication,
  Logger,
  NotFoundException,
  Post,
} from '@nestjs/common';
import { Test } from '@nestjs/testing';
import { HealthCheckService, TypeOrmHealthIndicator } from '@nestjs/terminus';
import { IsNotEmpty, IsString } from 'class-validator';
import request from 'supertest';
import { CORRELATION_ID_HEADER, ErrorCodes } from '@orcax/api-contracts';
import { configureHttp } from '../src/bootstrap/configure-http';
import { HealthController } from '../src/health/health.controller';

class SampleBody {
  @IsString()
  @IsNotEmpty()
  name!: string;
}

@Controller('samples')
class SampleController {
  @Get()
  show(): { ready: boolean } {
    return { ready: true };
  }

  @Post()
  create(@Body() body: SampleBody): { accepted: string } {
    return { accepted: body.name };
  }

  @Get('missing')
  missing(): never {
    throw new NotFoundException('Không tìm thấy bản ghi.');
  }

  @Get('boom')
  boom(): never {
    throw new Error('database password=should-not-leak');
  }

  @Get('unverified')
  unverified(): never {
    throw new HttpException(
      { code: ErrorCodes.EMAIL_NOT_VERIFIED, message: 'Email chưa được xác minh.' },
      HttpStatus.FORBIDDEN,
    );
  }
}

describe('HTTP foundation', () => {
  let app: INestApplication;

  before(async () => {
    const moduleRef = await Test.createTestingModule({
      controllers: [SampleController, HealthController],
      providers: [
        { provide: HealthCheckService, useValue: { check: async () => ({}) } },
        { provide: TypeOrmHealthIndicator, useValue: { pingCheck: async () => ({}) } },
      ],
    }).compile();
    app = moduleRef.createNestApplication();
    configureHttp(app);
    await app.init();
    Logger.overrideLogger({
      log() {},
      error() {},
      warn() {},
      debug() {},
      verbose() {},
      fatal() {},
    });
  });

  after(async () => {
    await app.close();
  });

  it('wraps successful JSON and keeps the caller correlation id', async () => {
    const correlationId = '55555555-5555-4555-8555-555555555555';
    const response = await request(app.getHttpServer())
      .get('/api/v1/samples')
      .set(CORRELATION_ID_HEADER, correlationId)
      .expect(200);

    assert.deepEqual(response.body.data, { ready: true });
    assert.equal(response.body.error, null);
    assert.equal(response.body.meta.correlationId, correlationId);
    assert.equal(response.body.meta.apiVersion, 'v1');
    assert.equal(response.headers[CORRELATION_ID_HEADER], correlationId);
  });

  it('validates DTOs and rejects unknown fields', async () => {
    const invalid = await request(app.getHttpServer()).post('/api/v1/samples').send({}).expect(400);
    assert.equal(invalid.body.error.code, 'VALIDATION_FAILED');
    assert.equal(invalid.body.error.message, 'Dữ liệu không hợp lệ.');

    await request(app.getHttpServer())
      .post('/api/v1/samples')
      .send({ name: 'A', extra: true })
      .expect(400);

    const created = await request(app.getHttpServer())
      .post('/api/v1/samples')
      .send({ name: 'A' })
      .expect(201);
    assert.deepEqual(created.body.data, { accepted: 'A' });
  });

  it('maps known HTTP errors and hides unexpected failures', async () => {
    const missing = await request(app.getHttpServer()).get('/api/v1/samples/missing').expect(404);
    assert.equal(missing.body.error.code, 'NOT_FOUND');
    assert.equal(missing.body.error.message, 'Không tìm thấy bản ghi.');

    const boom = await request(app.getHttpServer()).get('/api/v1/samples/boom').expect(500);
    assert.equal(boom.body.error.code, 'INTERNAL_ERROR');
    assert.equal(boom.body.error.message, 'Đã xảy ra lỗi nội bộ.');
    assert.equal(JSON.stringify(boom.body).includes('should-not-leak'), false);

    const unverified = await request(app.getHttpServer())
      .get('/api/v1/samples/unverified')
      .expect(403);
    assert.equal(unverified.body.error.code, 'EMAIL_NOT_VERIFIED');
  });

  it('leaves health probes outside the API prefix and envelope', async () => {
    const health = await request(app.getHttpServer()).get('/health').expect(200);
    assert.deepEqual(health.body, { status: 'ok' });
    await request(app.getHttpServer()).get('/api/v1/health').expect(404);
    await request(app.getHttpServer()).get('/samples').expect(404);
  });
});
