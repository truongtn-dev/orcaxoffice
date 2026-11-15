import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { repoEnvFilePath } from './config/env-file-path';
import { validateEnv } from './config/validate-env';
import { DatabaseModule } from './database/database.module';
import { HealthModule } from './health/health.module';
import { AuditModule } from './modules/audit/audit.module';
import { SHELL_MODULES } from './modules/register-shells';
import { ObservabilityModule } from './observability/observability.module';

const configModule = ConfigModule.forRoot({
  isGlobal: true,
  envFilePath: repoEnvFilePath(),
  validate: validateEnv,
});
// forRoot starts loading immediately. Bootstrap still awaits the same promise.
void Promise.resolve(configModule).catch(() => undefined);

@Module({
  imports: [
    configModule,
    ObservabilityModule,
    DatabaseModule,
    AuditModule,
    HealthModule,
    ...SHELL_MODULES,
  ],
})
export class AppModule {}
