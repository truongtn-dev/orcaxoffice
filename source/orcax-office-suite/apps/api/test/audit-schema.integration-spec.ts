import assert from 'node:assert/strict';
import { after, before, describe, it } from 'node:test';
import { randomUUID } from 'crypto';
import { DataSource } from 'typeorm';
import { CreateAuditEvents1763121600000 } from '../src/database/migrations/1763121600000-create-audit-events';
import { MIGRATIONS_TABLE, readDatabaseOptionsFromEnv } from '../src/database/database-options';
import { AuditEvent } from '../src/modules/audit/audit-event.entity';

describe('audit_events schema', () => {
  let dataSource: DataSource;

  before(
    async () => {
      dataSource = new DataSource({
        ...readDatabaseOptionsFromEnv(process.env),
        entities: [AuditEvent],
        migrations: [CreateAuditEvents1763121600000],
        migrationsTableName: MIGRATIONS_TABLE,
        logging: false,
      });
      await dataSource.initialize();
      await dataSource.runMigrations();
    },
    { timeout: 30_000 },
  );

  after(async () => {
    if (dataSource?.isInitialized) {
      await dataSource.destroy();
    }
  });

  it('accepts an insert and rejects update, delete, and truncate', async () => {
    const id = randomUUID();
    const correlationId = randomUUID();
    await dataSource.query(
      `INSERT INTO audit_events
        (id, action, target_type, occurred_at, outcome, correlation_id)
       VALUES ($1, $2, $3, NOW(), $4, $5)`,
      [id, 'foundation.probe', 'audit_event', 'success', correlationId],
    );

    await assert.rejects(
      () => dataSource.query(`UPDATE audit_events SET action = $1 WHERE id = $2`, ['changed', id]),
      /append-only/,
    );
    await assert.rejects(
      () => dataSource.query(`DELETE FROM audit_events WHERE id = $1`, [id]),
      /append-only/,
    );
    await assert.rejects(() => dataSource.query(`TRUNCATE audit_events`), /append-only/);
    await assert.rejects(() =>
      dataSource.query(
        `INSERT INTO audit_events
          (id, action, target_type, occurred_at, outcome, correlation_id)
         VALUES ($1, $2, $3, NOW(), $4, $5)`,
        [randomUUID(), 'foundation.probe', 'audit_event', 'maybe', randomUUID()],
      ),
    );
  });
});
