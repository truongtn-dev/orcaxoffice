import { config as loadDotenv } from 'dotenv';
import { DataSource } from 'typeorm';
import { join } from 'path';
import { repoEnvFilePath } from '../config/env-file-path';
import { MIGRATIONS_TABLE, readDatabaseOptionsFromEnv } from './database-options';

loadDotenv({ path: repoEnvFilePath() });

export const AppDataSource = new DataSource({
  ...readDatabaseOptionsFromEnv(process.env),
  entities: [join(__dirname, '..', '**', '*.entity.ts')],
  migrations: [join(__dirname, 'migrations', '*.ts')],
  migrationsTableName: MIGRATIONS_TABLE,
  logging: false,
});
