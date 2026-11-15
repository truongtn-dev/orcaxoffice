import assert from 'node:assert/strict';
import { after, before, describe, it } from 'node:test';
import { randomUUID } from 'crypto';
import { DataSource } from 'typeorm';
import { CreateAuditEvents1763121600000 } from '../src/database/migrations/1763121600000-create-audit-events';
import { IdentityWorkspace1763125200000 } from '../src/database/migrations/1763125200000-identity-workspace';
import { E03PersonalAuditRead1763128800000 } from '../src/database/migrations/1763128800000-e03-personal-audit-read';
import { MIGRATIONS_TABLE, readDatabaseOptionsFromEnv } from '../src/database/database-options';
import { PERMISSIONS, ROLE_PERMISSIONS, ROLES } from '../src/modules/authorization/catalog';

describe('identity workspace schema', () => {
  let dataSource: DataSource;

  before(
    async () => {
      dataSource = new DataSource({
        ...readDatabaseOptionsFromEnv(process.env),
        migrations: [
          CreateAuditEvents1763121600000,
          IdentityWorkspace1763125200000,
          E03PersonalAuditRead1763128800000,
        ],
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

  it('seeds the same role permissions the runtime catalog uses', async () => {
    const rows = await dataSource.query(
      `SELECT r.code AS role, p.code AS permission
       FROM role_permissions rp
       JOIN roles r ON r.id = rp.role_id
       JOIN permissions p ON p.id = rp.permission_id`,
    );
    const seeded = new Map<string, Set<string>>();
    for (const row of rows as { role: string; permission: string }[]) {
      const current = seeded.get(row.role) ?? new Set<string>();
      current.add(row.permission);
      seeded.set(row.role, current);
    }
    assert.equal(seeded.size, ROLES.length);
    for (const role of ROLES) {
      assert.deepEqual(
        [...(seeded.get(role.code) ?? [])].sort(),
        [...ROLE_PERMISSIONS[role.code]].sort(),
      );
    }
    const permissionCount = await dataSource.query(
      `SELECT COUNT(*)::int AS count FROM permissions`,
    );
    assert.equal(permissionCount[0].count, PERMISSIONS.length);
  });

  it('separates a personal workspace from an organization workspace', async () => {
    const userId = randomUUID();
    const now = new Date();
    await dataSource.query(
      `INSERT INTO users (id, email, password_hash, display_name, status, created_at, updated_at)
       VALUES ($1, $2, 'hash', 'Người dùng', 'active', $3, $3)`,
      [userId, `${userId}@example.com`, now],
    );

    const personalId = randomUUID();
    await dataSource.query(
      `INSERT INTO workspaces
        (id, kind, name, status, owner_user_id, organization_id, created_at, updated_at)
       VALUES ($1, 'personal', 'Cá nhân', 'active', $2, NULL, $3, $3)`,
      [personalId, userId, now],
    );

    await assert.rejects(() =>
      dataSource.query(
        `INSERT INTO workspaces
          (id, kind, name, status, owner_user_id, organization_id, created_at, updated_at)
         VALUES ($1, 'personal', 'Cá nhân 2', 'active', $2, NULL, $3, $3)`,
        [randomUUID(), userId, now],
      ),
    );

    const organizationId = randomUUID();
    await dataSource.query(
      `INSERT INTO organizations (id, name, status, created_by, created_at, updated_at)
       VALUES ($1, 'Công ty', 'active', $2, $3, $3)`,
      [organizationId, userId, now],
    );
    await assert.rejects(() =>
      dataSource.query(
        `INSERT INTO workspaces
          (id, kind, name, status, owner_user_id, organization_id, created_at, updated_at)
         VALUES ($1, 'organization', 'Công ty', 'active', $2, $3, $4, $4)`,
        [randomUUID(), userId, organizationId, now],
      ),
    );
    await assert.rejects(() =>
      dataSource.query(`UPDATE workspaces SET status = 'suspended' WHERE id = $1`, [personalId]),
    );

    const organizationWorkspaceId = randomUUID();
    await dataSource.query(
      `INSERT INTO workspaces
        (id, kind, name, status, owner_user_id, organization_id, created_at, updated_at)
       VALUES ($1, 'organization', 'Công ty', 'active', NULL, $2, $3, $3)`,
      [organizationWorkspaceId, organizationId, now],
    );
    const stored = await dataSource.query(
      `SELECT kind, owner_user_id, organization_id FROM workspaces WHERE id = $1`,
      [organizationWorkspaceId],
    );
    assert.equal(stored[0].kind, 'organization');
    assert.equal(stored[0].owner_user_id, null);
    assert.equal(stored[0].organization_id, organizationId);
  });
});
