import { MigrationInterface, QueryRunner } from 'typeorm';

const PERSONAL_OWNER_ID = '10000000-0000-4000-8000-000000000001';
const AUDIT_READ_ID = '20000000-0000-4000-8000-000000000007';

export class E03PersonalAuditRead1763128800000 implements MigrationInterface {
  name = 'E03PersonalAuditRead1763128800000';

  public async up(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(
      `INSERT INTO "role_permissions" ("role_id", "permission_id")
       SELECT $1, $2
       WHERE NOT EXISTS (
         SELECT 1 FROM "role_permissions"
         WHERE "role_id" = $1 AND "permission_id" = $2
       )`,
      [PERSONAL_OWNER_ID, AUDIT_READ_ID],
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(
      `DELETE FROM "role_permissions" WHERE "role_id" = $1 AND "permission_id" = $2`,
      [PERSONAL_OWNER_ID, AUDIT_READ_ID],
    );
  }
}
