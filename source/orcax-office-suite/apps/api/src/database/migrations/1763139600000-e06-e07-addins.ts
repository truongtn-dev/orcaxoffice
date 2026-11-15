import { MigrationInterface, QueryRunner } from 'typeorm';
import { PERMISSIONS, ROLE_PERMISSIONS, ROLES } from '../../modules/authorization/catalog';

const ADDIN_PERMISSION_CODES = ['addin.word.use', 'addin.excel.use'] as const;

export class E06E07Addins1763139600000 implements MigrationInterface {
  name = 'E06E07Addins1763139600000';

  public async up(queryRunner: QueryRunner): Promise<void> {
    for (const permission of PERMISSIONS.filter((item) =>
      (ADDIN_PERMISSION_CODES as readonly string[]).includes(item.code),
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
        if (!(ADDIN_PERMISSION_CODES as readonly string[]).includes(code)) {
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

    await queryRunner.query(`CREATE TABLE "validation_runs" (
      "id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "user_id" uuid NOT NULL,
      "host" character varying(16) NOT NULL,
      "profile_id" uuid NOT NULL,
      "profile_version_id" uuid NOT NULL,
      "profile_version" integer NOT NULL,
      "profile_code" character varying(64) NOT NULL,
      "summary_pass" integer NOT NULL,
      "summary_fail" integer NOT NULL,
      "summary_warning" integer NOT NULL,
      "summary_corrected" integer NOT NULL,
      "adapter" character varying(32) NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_validation_runs" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_validation_runs_host" CHECK ("host" IN ('word', 'excel')),
      CONSTRAINT "FK_validation_runs_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_validation_runs_workspace_id" ON "validation_runs" ("workspace_id")`,
    );

    await queryRunner.query(`CREATE TABLE "print_runs" (
      "id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "user_id" uuid NOT NULL,
      "profile_id" uuid NOT NULL,
      "profile_version_id" uuid NOT NULL,
      "profile_version" integer NOT NULL,
      "profile_code" character varying(64) NOT NULL,
      "paper_size" character varying(32) NOT NULL,
      "orientation" character varying(32) NOT NULL,
      "applied" boolean NOT NULL,
      "pdf_exported" boolean NOT NULL,
      "adapter" character varying(32) NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_print_runs" PRIMARY KEY ("id"),
      CONSTRAINT "FK_print_runs_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_print_runs_workspace_id" ON "print_runs" ("workspace_id")`,
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(`DROP TABLE IF EXISTS "print_runs"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "validation_runs"`);

    const permissionId = new Map(PERMISSIONS.map((item) => [item.code, item.id]));
    for (const role of ROLES) {
      for (const code of ADDIN_PERMISSION_CODES) {
        if (!(ROLE_PERMISSIONS[role.code] as readonly string[]).includes(code)) {
          continue;
        }
        await queryRunner.query(
          `DELETE FROM "role_permissions" WHERE "role_id" = $1 AND "permission_id" = $2`,
          [role.id, permissionId.get(code)],
        );
      }
    }
    for (const code of ADDIN_PERMISSION_CODES) {
      await queryRunner.query(`DELETE FROM "permissions" WHERE "id" = $1`, [
        permissionId.get(code),
      ]);
    }
  }
}
