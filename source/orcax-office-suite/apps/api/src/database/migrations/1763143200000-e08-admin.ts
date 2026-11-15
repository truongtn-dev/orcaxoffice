import { MigrationInterface, QueryRunner } from 'typeorm';
import { PERMISSIONS, ROLE_PERMISSIONS, ROLES } from '../../modules/authorization/catalog';

const ADMIN_PERMISSION_CODES = [
  'workspace.department.manage',
  'workspace.config.read',
  'workspace.config.manage',
] as const;

/** Also backfill admin role grants introduced in E08 (update + role). */
const ADMIN_ROLE_BACKFILL = [
  'workspace.update',
  'workspace.member.role',
  'workspace.department.manage',
  'workspace.config.read',
  'workspace.config.manage',
] as const;

export class E08Admin1763143200000 implements MigrationInterface {
  name = 'E08Admin1763143200000';

  public async up(queryRunner: QueryRunner): Promise<void> {
    for (const permission of PERMISSIONS.filter((item) =>
      (ADMIN_PERMISSION_CODES as readonly string[]).includes(item.code),
    )) {
      await queryRunner.query(
        `INSERT INTO "permissions" ("id", "code")
         SELECT $1, $2
         WHERE NOT EXISTS (SELECT 1 FROM "permissions" WHERE "id" = $1 OR "code" = $2)`,
        [permission.id, permission.code],
      );
    }

    const permissionId = new Map(PERMISSIONS.map((item) => [item.code, item.id]));
    for (const role of ROLES) {
      for (const code of ROLE_PERMISSIONS[role.code]) {
        const shouldInsert =
          (ADMIN_PERMISSION_CODES as readonly string[]).includes(code) ||
          (role.code === 'organization_administrator' &&
            (ADMIN_ROLE_BACKFILL as readonly string[]).includes(code));
        if (!shouldInsert) {
          continue;
        }
        await queryRunner.query(
          `INSERT INTO "role_permissions" ("role_id", "permission_id")
           SELECT $1, $2
           WHERE NOT EXISTS (
             SELECT 1 FROM "role_permissions"
             WHERE "role_id" = $1 AND "permission_id" = $2
           )`,
          [role.id, permissionId.get(code)],
        );
      }
    }

    await queryRunner.query(`CREATE TABLE "departments" (
      "id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "parent_id" uuid,
      "code" character varying(64) NOT NULL,
      "name" character varying(160) NOT NULL,
      "status" character varying(32) NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_departments" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_departments_status" CHECK ("status" IN ('active', 'archived')),
      CONSTRAINT "FK_departments_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_departments_parent" FOREIGN KEY ("parent_id") REFERENCES "departments"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE UNIQUE INDEX "UQ_departments_workspace_code" ON "departments" ("workspace_id", "code")`,
    );
    await queryRunner.query(
      `CREATE INDEX "IDX_departments_workspace_id" ON "departments" ("workspace_id")`,
    );

    await queryRunner.query(`CREATE TABLE "org_groups" (
      "id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "code" character varying(64) NOT NULL,
      "name" character varying(160) NOT NULL,
      "status" character varying(32) NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_org_groups" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_org_groups_status" CHECK ("status" IN ('active', 'archived')),
      CONSTRAINT "FK_org_groups_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE UNIQUE INDEX "UQ_org_groups_workspace_code" ON "org_groups" ("workspace_id", "code")`,
    );
    await queryRunner.query(
      `CREATE INDEX "IDX_org_groups_workspace_id" ON "org_groups" ("workspace_id")`,
    );

    await queryRunner.query(`CREATE TABLE "department_members" (
      "department_id" uuid NOT NULL,
      "user_id" uuid NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_department_members" PRIMARY KEY ("department_id", "user_id"),
      CONSTRAINT "FK_department_members_department" FOREIGN KEY ("department_id") REFERENCES "departments"("id") ON DELETE CASCADE
    )`);

    await queryRunner.query(`CREATE TABLE "group_members" (
      "group_id" uuid NOT NULL,
      "user_id" uuid NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_group_members" PRIMARY KEY ("group_id", "user_id"),
      CONSTRAINT "FK_group_members_group" FOREIGN KEY ("group_id") REFERENCES "org_groups"("id") ON DELETE CASCADE
    )`);

    await queryRunner.query(`CREATE TABLE "profile_distributions" (
      "id" uuid NOT NULL,
      "profile_id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "target_type" character varying(32) NOT NULL,
      "target_id" uuid NOT NULL,
      "distributed_by" uuid NOT NULL,
      "distributed_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_profile_distributions" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_profile_distributions_target" CHECK ("target_type" IN ('department', 'group')),
      CONSTRAINT "FK_profile_distributions_profile" FOREIGN KEY ("profile_id") REFERENCES "profiles"("id") ON DELETE CASCADE,
      CONSTRAINT "FK_profile_distributions_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_profile_distributions_profile_id" ON "profile_distributions" ("profile_id")`,
    );
    await queryRunner.query(
      `CREATE INDEX "IDX_profile_distributions_workspace_id" ON "profile_distributions" ("workspace_id")`,
    );

    await queryRunner.query(`CREATE TABLE "workspace_settings" (
      "id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "key" character varying(64) NOT NULL,
      "value" jsonb NOT NULL,
      "version" integer NOT NULL,
      "updated_by" uuid NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_workspace_settings" PRIMARY KEY ("id"),
      CONSTRAINT "FK_workspace_settings_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE UNIQUE INDEX "UQ_workspace_settings_workspace_key" ON "workspace_settings" ("workspace_id", "key")`,
    );

    await queryRunner.query(`CREATE TABLE "workspace_setting_history" (
      "id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "key" character varying(64) NOT NULL,
      "value" jsonb NOT NULL,
      "version" integer NOT NULL,
      "changed_by" uuid NOT NULL,
      "changed_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_workspace_setting_history" PRIMARY KEY ("id"),
      CONSTRAINT "FK_workspace_setting_history_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_workspace_setting_history_workspace_id" ON "workspace_setting_history" ("workspace_id")`,
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(`DROP TABLE IF EXISTS "workspace_setting_history"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "workspace_settings"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "profile_distributions"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "group_members"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "department_members"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "org_groups"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "departments"`);

    const permissionId = new Map(PERMISSIONS.map((item) => [item.code, item.id]));
    for (const role of ROLES) {
      for (const code of ADMIN_PERMISSION_CODES) {
        if (!(ROLE_PERMISSIONS[role.code] as readonly string[]).includes(code)) {
          continue;
        }
        await queryRunner.query(
          `DELETE FROM "role_permissions" WHERE "role_id" = $1 AND "permission_id" = $2`,
          [role.id, permissionId.get(code)],
        );
      }
    }
    for (const code of ADMIN_PERMISSION_CODES) {
      await queryRunner.query(`DELETE FROM "permissions" WHERE "id" = $1`, [
        permissionId.get(code),
      ]);
    }
  }
}
