import { Module } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { TypeOrmModule } from '@nestjs/typeorm';
import type { AppConfig } from '../config/validate-env';

@Module({
  imports: [
    TypeOrmModule.forRootAsync({
      inject: [ConfigService],
      useFactory: (config: ConfigService) => {
        const database = config.getOrThrow<AppConfig['database']>('database');
        return {
          type: 'postgres' as const,
          host: database.host,
          port: database.port,
          username: database.user,
          password: database.password,
          database: database.name,
          ssl: database.ssl ? { rejectUnauthorized: true } : false,
          autoLoadEntities: true,
          synchronize: false,
          migrationsRun: false,
          logging: false,
          retryAttempts: 2,
        };
      },
    }),
  ],
})
export class DatabaseModule {}
