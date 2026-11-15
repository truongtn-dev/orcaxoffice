import type { DataSourceOptions } from 'typeorm';

export const MIGRATIONS_TABLE = 'schema_migrations';

const REQUIRED_DATABASE_KEYS = [
  'DATABASE_HOST',
  'DATABASE_PORT',
  'DATABASE_USER',
  'DATABASE_PASSWORD',
  'DATABASE_NAME',
] as const;

export function readDatabaseOptionsFromEnv(env: Record<string, unknown>): DataSourceOptions {
  const missing = REQUIRED_DATABASE_KEYS.filter((key) => {
    const value = env[key];
    return value === undefined || value === null || value === '';
  });
  if (missing.length > 0) {
    throw new Error(`Missing database environment: ${missing.join(', ')}`);
  }

  const rawPort = env.DATABASE_PORT;
  const port = typeof rawPort === 'number' ? rawPort : Number(rawPort);
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error('DATABASE_PORT must be an integer from 1 to 65535');
  }

  const ssl = env.DATABASE_SSL === true || env.DATABASE_SSL === 'true';
  return {
    type: 'postgres',
    host: String(env.DATABASE_HOST),
    port,
    username: String(env.DATABASE_USER),
    password: String(env.DATABASE_PASSWORD),
    database: String(env.DATABASE_NAME),
    ssl: ssl ? { rejectUnauthorized: true } : false,
    synchronize: false,
  };
}
