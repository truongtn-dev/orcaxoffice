import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { readDatabaseOptionsFromEnv } from '../src/database/database-options';
import { validateEnv } from '../src/config/validate-env';

const validEnv = {
  NODE_ENV: 'test',
  API_PORT: '3001',
  LOG_LEVEL: 'silent',
  WEB_ORIGIN: 'http://localhost:3000',
  DATABASE_HOST: 'localhost',
  DATABASE_PORT: '5432',
  DATABASE_USER: 'orcax',
  DATABASE_PASSWORD: 'super-secret-value',
  DATABASE_NAME: 'orcax_office',
  DATABASE_SSL: 'false',
  AUTH_TOKEN_SECRET: 'unit-test-token-secret-32chars-min',
  PAYMENT_WEBHOOK_SECRET: 'unit-test-payment-webhook',
  EMAIL_TRANSPORT: 'log',
};

describe('environment validation', () => {
  it('parses the foundation settings', () => {
    const config = validateEnv(validEnv);
    assert.equal(config.nodeEnv, 'test');
    assert.equal(config.apiPort, 3001);
    assert.equal(config.database.ssl, false);
    assert.equal(config.emailTransport, 'log');
    assert.equal(config.accessTtlSeconds, 900);
    assert.deepEqual(
      {
        host: config.database.host,
        port: config.database.port,
        user: config.database.user,
        name: config.database.name,
      },
      {
        host: 'localhost',
        port: 5432,
        user: 'orcax',
        name: 'orcax_office',
      },
    );
  });

  it('does not echo a password when validation fails', () => {
    assert.throws(() => validateEnv({ ...validEnv, DATABASE_NAME: '' }), /DATABASE_NAME/);
    try {
      validateEnv({ ...validEnv, DATABASE_NAME: '' });
    } catch (error) {
      assert.ok(error instanceof Error);
      assert.equal(error.message.includes('super-secret-value'), false);
    }
  });

  it('rejects a log mailbox in production and requires SMTP settings', () => {
    assert.throws(
      () => validateEnv({ ...validEnv, NODE_ENV: 'production', EMAIL_TRANSPORT: 'log' }),
      /EMAIL_TRANSPORT/,
    );
    assert.throws(
      () => validateEnv({ ...validEnv, EMAIL_TRANSPORT: 'smtp', SMTP_HOST: '' }),
      /SMTP_HOST/,
    );
  });

  it('requires database settings and keeps synchronize disabled', () => {
    assert.throws(() => readDatabaseOptionsFromEnv({}), /DATABASE_HOST/);
    const options = readDatabaseOptionsFromEnv({ ...validEnv, DATABASE_SSL: 'true' });
    assert.equal(options.type, 'postgres');
    if (options.type !== 'postgres') {
      return;
    }
    assert.equal(options.synchronize, false);
    assert.deepEqual(options.ssl, { rejectUnauthorized: true });
  });
});
