import { envValidationSchema } from './env.schema';
import { readDatabaseOptionsFromEnv } from '../database/database-options';

export interface AppConfig {
  nodeEnv: 'development' | 'test' | 'production';
  apiPort: number;
  logLevel: string;
  webOrigin: string;
  authTokenSecret: string;
  accessTtlSeconds: number;
  sessionTtlDays: number;
  emailTransport: 'log' | 'smtp';
  emailFrom: string;
  smtpHost: string;
  smtpPort: number;
  smtpUser: string;
  smtpPassword: string;
  paymentWebhookSecret: string;
  database: {
    host: string;
    port: number;
    user: string;
    password: string;
    name: string;
    ssl: boolean;
  };
}

export function validateEnv(env: Record<string, unknown>): AppConfig {
  const { error, value } = envValidationSchema.validate(env, {
    abortEarly: false,
    allowUnknown: true,
    convert: true,
  });
  if (error) {
    throw new Error(`Environment validation failed: ${error.message}`);
  }

  const database = readDatabaseOptionsFromEnv(value as Record<string, unknown>);
  if (database.type !== 'postgres') {
    throw new Error('Environment validation failed: database driver must be postgres');
  }

  const emailTransport = value.EMAIL_TRANSPORT as AppConfig['emailTransport'];
  if (value.NODE_ENV === 'production' && emailTransport === 'log') {
    throw new Error('Environment validation failed: EMAIL_TRANSPORT must be smtp in production');
  }
  const smtpHost = String(value.SMTP_HOST ?? '');
  const smtpPort = Number(value.SMTP_PORT ?? 0);
  if (emailTransport === 'smtp' && (smtpHost.length === 0 || !Number.isInteger(smtpPort))) {
    throw new Error('Environment validation failed: SMTP_HOST and SMTP_PORT are required');
  }

  return {
    nodeEnv: value.NODE_ENV as AppConfig['nodeEnv'],
    apiPort: value.API_PORT as number,
    logLevel: value.LOG_LEVEL as string,
    webOrigin: (value.WEB_ORIGIN as string | undefined) ?? '',
    authTokenSecret: String(value.AUTH_TOKEN_SECRET),
    accessTtlSeconds: Number(value.AUTH_ACCESS_TTL_SECONDS),
    sessionTtlDays: Number(value.AUTH_SESSION_TTL_DAYS),
    emailTransport,
    emailFrom: String(value.EMAIL_FROM),
    smtpHost,
    smtpPort,
    smtpUser: String(value.SMTP_USER ?? ''),
    smtpPassword: String(value.SMTP_PASSWORD ?? ''),
    paymentWebhookSecret: String(value.PAYMENT_WEBHOOK_SECRET),
    database: {
      host: String(database.host),
      port: Number(database.port),
      user: String(database.username),
      password: String(database.password),
      name: String(database.database),
      ssl: database.ssl !== false,
    },
  };
}
