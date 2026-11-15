import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { QueryRunner } from 'typeorm';
import { CreateAuditEvents1763121600000 } from '../src/database/migrations/1763121600000-create-audit-events';

describe('audit_events migration', () => {
  it('creates an append-only table and can drop it', async () => {
    const up: string[] = [];
    const down: string[] = [];
    const migration = new CreateAuditEvents1763121600000();
    await migration.up({
      query: async (sql: string) => {
        up.push(sql);
      },
    } as unknown as QueryRunner);
    await migration.down({
      query: async (sql: string) => {
        down.push(sql);
      },
    } as unknown as QueryRunner);

    const created = up.join('\n');
    assert.match(created, /CREATE TABLE "audit_events"/);
    assert.match(created, /audit_events is append-only/);
    assert.match(created, /BEFORE UPDATE OR DELETE/);
    assert.match(created, /BEFORE TRUNCATE/);
    assert.match(down.join('\n'), /DROP TABLE IF EXISTS "audit_events"/);
  });
});
