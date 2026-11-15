import { MigrationInterface, QueryRunner } from 'typeorm';
import { PERMISSIONS, ROLE_PERMISSIONS, ROLES } from '../../modules/authorization/catalog';
import { PLANS, PRODUCTS } from '../../modules/licensing/license-catalog';

const LICENSE_PERMISSION_CODES = [
  'license.read',
  'license.issue',
  'license.manage',
  'license.activate',
] as const;

export class E04Licensing1763132400000 implements MigrationInterface {
  name = 'E04Licensing1763132400000';

  public async up(queryRunner: QueryRunner): Promise<void> {
    for (const permission of PERMISSIONS.filter((item) =>
      (LICENSE_PERMISSION_CODES as readonly string[]).includes(item.code),
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
        if (!(LICENSE_PERMISSION_CODES as readonly string[]).includes(code)) {
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

    await queryRunner.query(`CREATE TABLE "products" (
      "id" uuid NOT NULL,
      "code" character varying(64) NOT NULL,
      "name" character varying(160) NOT NULL,
      "status" character varying(32) NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_products" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_products_status" CHECK ("status" IN ('active', 'retired'))
    )`);
    await queryRunner.query(`CREATE UNIQUE INDEX "UQ_products_code" ON "products" ("code")`);

    await queryRunner.query(`CREATE TABLE "plans" (
      "id" uuid NOT NULL,
      "product_id" uuid NOT NULL,
      "code" character varying(64) NOT NULL,
      "name" character varying(160) NOT NULL,
      "license_type" character varying(32) NOT NULL,
      "workspace_kind" character varying(32) NOT NULL,
      "device_limit" integer NOT NULL,
      "seat_limit" integer NOT NULL,
      "term_days" integer,
      "grace_days" integer NOT NULL,
      "status" character varying(32) NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_plans" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_plans_license_type" CHECK ("license_type" IN ('subscription', 'perpetual')),
      CONSTRAINT "CHK_plans_workspace_kind" CHECK ("workspace_kind" IN ('personal', 'organization')),
      CONSTRAINT "CHK_plans_status" CHECK ("status" IN ('active', 'retired')),
      CONSTRAINT "CHK_plans_device_limit" CHECK ("device_limit" >= 1),
      CONSTRAINT "CHK_plans_seat_limit" CHECK ("seat_limit" >= 1),
      CONSTRAINT "FK_plans_product" FOREIGN KEY ("product_id") REFERENCES "products"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(`CREATE UNIQUE INDEX "UQ_plans_code" ON "plans" ("code")`);

    await queryRunner.query(`CREATE TABLE "plan_entitlements" (
      "plan_id" uuid NOT NULL,
      "entitlement_code" character varying(64) NOT NULL,
      CONSTRAINT "PK_plan_entitlements" PRIMARY KEY ("plan_id", "entitlement_code"),
      CONSTRAINT "FK_plan_entitlements_plan" FOREIGN KEY ("plan_id") REFERENCES "plans"("id") ON DELETE CASCADE
    )`);

    await queryRunner.query(`CREATE TABLE "licenses" (
      "id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "plan_id" uuid NOT NULL,
      "status" character varying(32) NOT NULL,
      "issued_by" uuid NOT NULL,
      "starts_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "ends_at" TIMESTAMP WITH TIME ZONE,
      "revoked_at" TIMESTAMP WITH TIME ZONE,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_licenses" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_licenses_status" CHECK ("status" IN ('active', 'revoked')),
      CONSTRAINT "FK_licenses_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_licenses_plan" FOREIGN KEY ("plan_id") REFERENCES "plans"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_licenses_issued_by" FOREIGN KEY ("issued_by") REFERENCES "users"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `COMMENT ON TABLE "licenses" IS 'Tenant-owned license rows. Scope by workspace_id.'`,
    );
    await queryRunner.query(
      `CREATE INDEX "IDX_licenses_workspace_id" ON "licenses" ("workspace_id")`,
    );

    await queryRunner.query(`CREATE TABLE "license_activations" (
      "id" uuid NOT NULL,
      "license_id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "user_id" uuid NOT NULL,
      "device_key_hash" character varying(128) NOT NULL,
      "device_label" character varying(160) NOT NULL,
      "activated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "last_seen_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "revoked_at" TIMESTAMP WITH TIME ZONE,
      CONSTRAINT "PK_license_activations" PRIMARY KEY ("id"),
      CONSTRAINT "FK_license_activations_license" FOREIGN KEY ("license_id") REFERENCES "licenses"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_license_activations_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_license_activations_user" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE UNIQUE INDEX "UQ_license_activations_live_device"
       ON "license_activations" ("license_id", "device_key_hash")
       WHERE "revoked_at" IS NULL`,
    );
    await queryRunner.query(
      `CREATE INDEX "IDX_license_activations_workspace_id" ON "license_activations" ("workspace_id")`,
    );

    const now = new Date();
    for (const product of PRODUCTS) {
      await queryRunner.query(
        `INSERT INTO "products" ("id", "code", "name", "status", "created_at")
         VALUES ($1, $2, $3, 'active', $4)`,
        [product.id, product.code, product.name, now],
      );
    }
    for (const plan of PLANS) {
      await queryRunner.query(
        `INSERT INTO "plans"
          ("id", "product_id", "code", "name", "license_type", "workspace_kind",
           "device_limit", "seat_limit", "term_days", "grace_days", "status", "created_at")
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,'active',$11)`,
        [
          plan.id,
          plan.productId,
          plan.code,
          plan.name,
          plan.licenseType,
          plan.workspaceKind,
          plan.deviceLimit,
          plan.seatLimit,
          plan.termDays,
          plan.graceDays,
          now,
        ],
      );
      for (const entitlement of plan.entitlements) {
        await queryRunner.query(
          `INSERT INTO "plan_entitlements" ("plan_id", "entitlement_code") VALUES ($1, $2)`,
          [plan.id, entitlement],
        );
      }
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(`DROP TABLE IF EXISTS "license_activations"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "licenses"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "plan_entitlements"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "plans"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "products"`);

    const permissionId = new Map(PERMISSIONS.map((item) => [item.code, item.id]));
    for (const role of ROLES) {
      for (const code of LICENSE_PERMISSION_CODES) {
        if (!(ROLE_PERMISSIONS[role.code] as readonly string[]).includes(code)) {
          continue;
        }
        await queryRunner.query(
          `DELETE FROM "role_permissions" WHERE "role_id" = $1 AND "permission_id" = $2`,
          [role.id, permissionId.get(code)],
        );
      }
    }
    for (const code of LICENSE_PERMISSION_CODES) {
      await queryRunner.query(`DELETE FROM "permissions" WHERE "code" = $1`, [code]);
    }
  }
}
