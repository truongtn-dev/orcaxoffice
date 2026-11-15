import { MigrationInterface, QueryRunner } from 'typeorm';
import { PERMISSIONS, ROLE_PERMISSIONS, ROLES } from '../../modules/authorization/catalog';

const UP_STATEMENTS = [
  `CREATE TABLE "users" (
    "id" uuid NOT NULL,
    "email" character varying(320) NOT NULL,
    "password_hash" character varying(255) NOT NULL,
    "display_name" character varying(120) NOT NULL,
    "status" character varying(32) NOT NULL,
    "email_verified_at" TIMESTAMP WITH TIME ZONE,
    "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    CONSTRAINT "PK_users" PRIMARY KEY ("id"),
    CONSTRAINT "CHK_users_status" CHECK ("status" IN ('pending', 'active', 'suspended'))
  )`,
  `CREATE UNIQUE INDEX "UQ_users_email" ON "users" ("email")`,
  `CREATE TABLE "roles" (
    "id" uuid NOT NULL,
    "code" character varying(64) NOT NULL,
    "kind" character varying(32) NOT NULL,
    CONSTRAINT "PK_roles" PRIMARY KEY ("id"),
    CONSTRAINT "CHK_roles_kind" CHECK ("kind" IN ('personal', 'organization'))
  )`,
  `CREATE UNIQUE INDEX "UQ_roles_code" ON "roles" ("code")`,
  `CREATE TABLE "permissions" (
    "id" uuid NOT NULL,
    "code" character varying(64) NOT NULL,
    CONSTRAINT "PK_permissions" PRIMARY KEY ("id")
  )`,
  `CREATE UNIQUE INDEX "UQ_permissions_code" ON "permissions" ("code")`,
  `CREATE TABLE "role_permissions" (
    "role_id" uuid NOT NULL,
    "permission_id" uuid NOT NULL,
    CONSTRAINT "PK_role_permissions" PRIMARY KEY ("role_id", "permission_id"),
    CONSTRAINT "FK_role_permissions_role" FOREIGN KEY ("role_id") REFERENCES "roles"("id") ON DELETE RESTRICT,
    CONSTRAINT "FK_role_permissions_permission" FOREIGN KEY ("permission_id") REFERENCES "permissions"("id") ON DELETE RESTRICT
  )`,
  `CREATE TABLE "organizations" (
    "id" uuid NOT NULL,
    "name" character varying(160) NOT NULL,
    "status" character varying(32) NOT NULL,
    "created_by" uuid NOT NULL,
    "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    CONSTRAINT "PK_organizations" PRIMARY KEY ("id"),
    CONSTRAINT "CHK_organizations_status" CHECK ("status" IN ('active', 'suspended', 'archived')),
    CONSTRAINT "FK_organizations_created_by" FOREIGN KEY ("created_by") REFERENCES "users"("id") ON DELETE RESTRICT
  )`,
  `CREATE TABLE "workspaces" (
    "id" uuid NOT NULL,
    "kind" character varying(32) NOT NULL,
    "name" character varying(160) NOT NULL,
    "status" character varying(32) NOT NULL,
    "owner_user_id" uuid,
    "organization_id" uuid,
    "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    CONSTRAINT "PK_workspaces" PRIMARY KEY ("id"),
    CONSTRAINT "CHK_workspaces_status" CHECK ("status" IN ('active', 'suspended', 'archived')),
    CONSTRAINT "CHK_workspaces_kind_owner" CHECK (
      ("kind" = 'personal' AND "owner_user_id" IS NOT NULL AND "organization_id" IS NULL)
      OR ("kind" = 'organization' AND "owner_user_id" IS NULL AND "organization_id" IS NOT NULL)
    ),
    CONSTRAINT "CHK_workspaces_personal_active" CHECK ("kind" <> 'personal' OR "status" = 'active'),
    CONSTRAINT "FK_workspaces_owner_user" FOREIGN KEY ("owner_user_id") REFERENCES "users"("id") ON DELETE RESTRICT,
    CONSTRAINT "FK_workspaces_organization" FOREIGN KEY ("organization_id") REFERENCES "organizations"("id") ON DELETE RESTRICT
  )`,
  `COMMENT ON TABLE "workspaces" IS 'Tenant boundary. Later modules scope tenant-owned rows by workspace_id. Personal and organization workspaces are different rows, not a nullable organization id.'`,
  `CREATE UNIQUE INDEX "UQ_workspaces_personal_owner" ON "workspaces" ("owner_user_id") WHERE "kind" = 'personal'`,
  `CREATE TABLE "memberships" (
    "id" uuid NOT NULL,
    "workspace_id" uuid NOT NULL,
    "user_id" uuid NOT NULL,
    "role_id" uuid NOT NULL,
    "status" character varying(32) NOT NULL,
    "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    CONSTRAINT "PK_memberships" PRIMARY KEY ("id"),
    CONSTRAINT "UQ_memberships_workspace_user" UNIQUE ("workspace_id", "user_id"),
    CONSTRAINT "CHK_memberships_status" CHECK ("status" IN ('active', 'suspended', 'removed')),
    CONSTRAINT "FK_memberships_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT,
    CONSTRAINT "FK_memberships_user" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE RESTRICT,
    CONSTRAINT "FK_memberships_role" FOREIGN KEY ("role_id") REFERENCES "roles"("id") ON DELETE RESTRICT
  )`,
  `CREATE INDEX "IDX_memberships_user_id" ON "memberships" ("user_id")`,
  `CREATE TABLE "invitations" (
    "id" uuid NOT NULL,
    "workspace_id" uuid NOT NULL,
    "email" character varying(320) NOT NULL,
    "role_id" uuid NOT NULL,
    "token_hash" character varying(128) NOT NULL,
    "status" character varying(32) NOT NULL,
    "invited_by" uuid NOT NULL,
    "expires_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    CONSTRAINT "PK_invitations" PRIMARY KEY ("id"),
    CONSTRAINT "CHK_invitations_status" CHECK ("status" IN ('invited', 'accepted', 'revoked', 'expired')),
    CONSTRAINT "FK_invitations_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT,
    CONSTRAINT "FK_invitations_role" FOREIGN KEY ("role_id") REFERENCES "roles"("id") ON DELETE RESTRICT,
    CONSTRAINT "FK_invitations_invited_by" FOREIGN KEY ("invited_by") REFERENCES "users"("id") ON DELETE RESTRICT
  )`,
  `CREATE UNIQUE INDEX "UQ_invitations_token_hash" ON "invitations" ("token_hash")`,
  `CREATE TABLE "sessions" (
    "id" uuid NOT NULL,
    "user_id" uuid NOT NULL,
    "secret_hash" character varying(128) NOT NULL,
    "active_workspace_id" uuid NOT NULL,
    "user_agent" character varying(256),
    "expires_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    "revoked_at" TIMESTAMP WITH TIME ZONE,
    "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    "rotated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    CONSTRAINT "PK_sessions" PRIMARY KEY ("id"),
    CONSTRAINT "FK_sessions_user" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE RESTRICT,
    CONSTRAINT "FK_sessions_workspace" FOREIGN KEY ("active_workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT
  )`,
  `CREATE INDEX "IDX_sessions_user_id" ON "sessions" ("user_id")`,
  `CREATE TABLE "verification_challenges" (
    "id" uuid NOT NULL,
    "user_id" uuid NOT NULL,
    "purpose" character varying(32) NOT NULL,
    "code_hash" character varying(255) NOT NULL,
    "attempts" integer NOT NULL DEFAULT 0,
    "expires_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    "consumed_at" TIMESTAMP WITH TIME ZONE,
    "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    CONSTRAINT "PK_verification_challenges" PRIMARY KEY ("id"),
    CONSTRAINT "CHK_verification_challenges_purpose" CHECK ("purpose" IN ('email_verification', 'password_reset')),
    CONSTRAINT "FK_verification_challenges_user" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE RESTRICT
  )`,
  `CREATE INDEX "IDX_verification_challenges_user_purpose" ON "verification_challenges" ("user_id", "purpose")`,
  `ALTER TABLE "users" ADD "last_workspace_id" uuid`,
  `ALTER TABLE "users" ADD CONSTRAINT "FK_users_last_workspace" FOREIGN KEY ("last_workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT`,
];

const DOWN_STATEMENTS = [
  `ALTER TABLE "users" DROP CONSTRAINT "FK_users_last_workspace"`,
  `ALTER TABLE "users" DROP COLUMN "last_workspace_id"`,
  `DROP TABLE "verification_challenges"`,
  `DROP TABLE "sessions"`,
  `DROP TABLE "invitations"`,
  `DROP TABLE "memberships"`,
  `DROP TABLE "workspaces"`,
  `DROP TABLE "organizations"`,
  `DROP TABLE "role_permissions"`,
  `DROP TABLE "permissions"`,
  `DROP TABLE "roles"`,
  `DROP TABLE "users"`,
];

export class IdentityWorkspace1763125200000 implements MigrationInterface {
  name = 'IdentityWorkspace1763125200000';

  public async up(queryRunner: QueryRunner): Promise<void> {
    for (const statement of UP_STATEMENTS) {
      await queryRunner.query(statement);
    }
    for (const role of ROLES) {
      await queryRunner.query(`INSERT INTO "roles" ("id", "code", "kind") VALUES ($1, $2, $3)`, [
        role.id,
        role.code,
        role.kind,
      ]);
    }
    for (const permission of PERMISSIONS) {
      await queryRunner.query(`INSERT INTO "permissions" ("id", "code") VALUES ($1, $2)`, [
        permission.id,
        permission.code,
      ]);
    }
    const permissionId = new Map(PERMISSIONS.map((item) => [item.code, item.id]));
    for (const role of ROLES) {
      for (const code of ROLE_PERMISSIONS[role.code]) {
        await queryRunner.query(
          `INSERT INTO "role_permissions" ("role_id", "permission_id") VALUES ($1, $2)`,
          [role.id, permissionId.get(code)],
        );
      }
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    for (const statement of DOWN_STATEMENTS) {
      await queryRunner.query(statement);
    }
  }
}
