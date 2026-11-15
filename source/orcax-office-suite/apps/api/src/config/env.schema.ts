import Joi from 'joi';

export const envValidationSchema = Joi.object({
  NODE_ENV: Joi.string().valid('development', 'test', 'production').default('development'),
  API_PORT: Joi.number().port().default(3001),
  LOG_LEVEL: Joi.string()
    .valid('fatal', 'error', 'warn', 'info', 'debug', 'trace', 'silent')
    .default('info'),
  WEB_ORIGIN: Joi.string()
    .uri({ scheme: ['http', 'https'] })
    .allow('')
    .optional(),
  DATABASE_HOST: Joi.string().min(1).required(),
  DATABASE_PORT: Joi.number().port().required(),
  DATABASE_USER: Joi.string().min(1).required(),
  DATABASE_PASSWORD: Joi.string().min(1).required(),
  DATABASE_NAME: Joi.string().min(1).required(),
  DATABASE_SSL: Joi.boolean().default(false),
  AUTH_TOKEN_SECRET: Joi.string().min(32).required(),
  AUTH_ACCESS_TTL_SECONDS: Joi.number().integer().min(60).max(3600).default(900),
  AUTH_SESSION_TTL_DAYS: Joi.number().integer().min(1).max(90).default(14),
  EMAIL_TRANSPORT: Joi.string().valid('log', 'smtp').default('log'),
  EMAIL_FROM: Joi.string().min(3).default('noreply@localhost'),
  SMTP_HOST: Joi.string().allow('').optional(),
  SMTP_PORT: Joi.number().port().optional(),
  SMTP_USER: Joi.string().allow('').optional(),
  SMTP_PASSWORD: Joi.string().allow('').optional(),
  PAYMENT_WEBHOOK_SECRET: Joi.string().min(16).required(),
}).unknown(true);
