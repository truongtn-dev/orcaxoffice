import { MigrationInterface, QueryRunner } from 'typeorm';
import { PERMISSIONS, ROLE_PERMISSIONS, ROLES } from '../../modules/authorization/catalog';
import { SYSTEM_PROFILES, SYSTEM_TEMPLATES } from '../../modules/profile/profile-catalog';

const PROFILE_PERMISSION_CODES = [
  'profile.read',
  'profile.write',
  'profile.review',
  'profile.publish',
  'template.read',
  'template.write',
  'template.publish',
] as const;

const SYSTEM_ACTOR = '00000000-0000-4000-8000-000000000001';

export class E05Profiles1763136000000 implements MigrationInterface {
  name = 'E05Profiles1763136000000';

  public async up(queryRunner: QueryRunner): Promise<void> {
    for (const permission of PERMISSIONS.filter((item) =>
      (PROFILE_PERMISSION_CODES as readonly string[]).includes(item.code),
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
        if (!(PROFILE_PERMISSION_CODES as readonly string[]).includes(code)) {
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

    await queryRunner.query(`CREATE TABLE "profiles" (
      "id" uuid NOT NULL,
      "code" character varying(64) NOT NULL,
      "name" character varying(160) NOT NULL,
      "host" character varying(16) NOT NULL,
      "scope" character varying(32) NOT NULL,
      "workspace_id" uuid,
      "owner_user_id" uuid,
      "status" character varying(32) NOT NULL,
      "current_version" integer NOT NULL,
      "created_by" uuid NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_profiles" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_profiles_host" CHECK ("host" IN ('word', 'excel')),
      CONSTRAINT "CHK_profiles_scope" CHECK ("scope" IN ('system', 'personal', 'organization')),
      CONSTRAINT "CHK_profiles_status" CHECK ("status" IN ('draft', 'in_review', 'approved', 'deprecated')),
      CONSTRAINT "FK_profiles_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE UNIQUE INDEX "UQ_profiles_system_code" ON "profiles" ("code") WHERE "workspace_id" IS NULL`,
    );
    await queryRunner.query(
      `CREATE UNIQUE INDEX "UQ_profiles_workspace_code" ON "profiles" ("workspace_id", "code") WHERE "workspace_id" IS NOT NULL`,
    );
    await queryRunner.query(
      `CREATE INDEX "IDX_profiles_workspace_id" ON "profiles" ("workspace_id")`,
    );

    await queryRunner.query(`CREATE TABLE "profile_versions" (
      "id" uuid NOT NULL,
      "profile_id" uuid NOT NULL,
      "version" integer NOT NULL,
      "status" character varying(32) NOT NULL,
      "source" text NOT NULL,
      "effective_at" TIMESTAMP WITH TIME ZONE,
      "author_id" uuid NOT NULL,
      "reviewer_id" uuid,
      "change_summary" text,
      "immutable" boolean NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "published_at" TIMESTAMP WITH TIME ZONE,
      CONSTRAINT "PK_profile_versions" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_profile_versions_status" CHECK ("status" IN ('draft', 'in_review', 'approved', 'deprecated')),
      CONSTRAINT "FK_profile_versions_profile" FOREIGN KEY ("profile_id") REFERENCES "profiles"("id") ON DELETE CASCADE
    )`);
    await queryRunner.query(
      `CREATE UNIQUE INDEX "UQ_profile_versions_profile_version" ON "profile_versions" ("profile_id", "version")`,
    );
    await queryRunner.query(
      `CREATE INDEX "IDX_profile_versions_profile_id" ON "profile_versions" ("profile_id")`,
    );

    await queryRunner.query(`CREATE TABLE "profile_rules" (
      "id" uuid NOT NULL,
      "profile_version_id" uuid NOT NULL,
      "rule_code" character varying(64) NOT NULL,
      "rule_version" character varying(32) NOT NULL,
      "default_severity" character varying(16) NOT NULL,
      "title" character varying(200) NOT NULL,
      "description" text NOT NULL,
      "check_kind" character varying(64) NOT NULL,
      "check_config" jsonb NOT NULL,
      "sort_order" integer NOT NULL,
      CONSTRAINT "PK_profile_rules" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_profile_rules_severity" CHECK ("default_severity" IN ('fail', 'warning')),
      CONSTRAINT "FK_profile_rules_version" FOREIGN KEY ("profile_version_id") REFERENCES "profile_versions"("id") ON DELETE CASCADE
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_profile_rules_version_id" ON "profile_rules" ("profile_version_id")`,
    );

    await queryRunner.query(`CREATE TABLE "legal_references" (
      "id" uuid NOT NULL,
      "profile_version_id" uuid NOT NULL,
      "citation" character varying(500) NOT NULL,
      "source_url" character varying(1000),
      "verified" boolean NOT NULL,
      "notes" text,
      CONSTRAINT "PK_legal_references" PRIMARY KEY ("id"),
      CONSTRAINT "FK_legal_references_version" FOREIGN KEY ("profile_version_id") REFERENCES "profile_versions"("id") ON DELETE CASCADE
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_legal_references_version_id" ON "legal_references" ("profile_version_id")`,
    );

    await queryRunner.query(`CREATE TABLE "verification_evidence" (
      "id" uuid NOT NULL,
      "legal_reference_id" uuid NOT NULL,
      "recorded_by" uuid NOT NULL,
      "recorded_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "evidence_type" character varying(32) NOT NULL,
      "summary" text NOT NULL,
      "reference_uri" character varying(1000),
      CONSTRAINT "PK_verification_evidence" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_verification_evidence_type" CHECK ("evidence_type" IN ('document', 'url', 'statement')),
      CONSTRAINT "FK_verification_evidence_legal" FOREIGN KEY ("legal_reference_id") REFERENCES "legal_references"("id") ON DELETE CASCADE
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_verification_evidence_legal_id" ON "verification_evidence" ("legal_reference_id")`,
    );

    await queryRunner.query(`CREATE TABLE "templates" (
      "id" uuid NOT NULL,
      "code" character varying(64) NOT NULL,
      "name" character varying(160) NOT NULL,
      "host" character varying(16) NOT NULL,
      "scope" character varying(32) NOT NULL,
      "workspace_id" uuid,
      "status" character varying(32) NOT NULL,
      "current_version" integer NOT NULL,
      "created_by" uuid NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_templates" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_templates_host" CHECK ("host" IN ('word', 'excel')),
      CONSTRAINT "CHK_templates_scope" CHECK ("scope" IN ('system', 'personal', 'organization')),
      CONSTRAINT "CHK_templates_status" CHECK ("status" IN ('draft', 'review', 'approved', 'published', 'deprecated')),
      CONSTRAINT "FK_templates_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE UNIQUE INDEX "UQ_templates_system_code" ON "templates" ("code") WHERE "workspace_id" IS NULL`,
    );
    await queryRunner.query(
      `CREATE UNIQUE INDEX "UQ_templates_workspace_code" ON "templates" ("workspace_id", "code") WHERE "workspace_id" IS NOT NULL`,
    );
    await queryRunner.query(
      `CREATE INDEX "IDX_templates_workspace_id" ON "templates" ("workspace_id")`,
    );

    await queryRunner.query(`CREATE TABLE "template_versions" (
      "id" uuid NOT NULL,
      "template_id" uuid NOT NULL,
      "version" integer NOT NULL,
      "status" character varying(32) NOT NULL,
      "author_id" uuid NOT NULL,
      "reviewer_id" uuid,
      "change_summary" text,
      "immutable" boolean NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "published_at" TIMESTAMP WITH TIME ZONE,
      CONSTRAINT "PK_template_versions" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_template_versions_status" CHECK ("status" IN ('draft', 'review', 'approved', 'published', 'deprecated')),
      CONSTRAINT "FK_template_versions_template" FOREIGN KEY ("template_id") REFERENCES "templates"("id") ON DELETE CASCADE
    )`);
    await queryRunner.query(
      `CREATE UNIQUE INDEX "UQ_template_versions_template_version" ON "template_versions" ("template_id", "version")`,
    );
    await queryRunner.query(
      `CREATE INDEX "IDX_template_versions_template_id" ON "template_versions" ("template_id")`,
    );

    await queryRunner.query(`CREATE TABLE "field_definitions" (
      "id" uuid NOT NULL,
      "template_version_id" uuid NOT NULL,
      "field_key" character varying(64) NOT NULL,
      "label" character varying(160) NOT NULL,
      "field_type" character varying(32) NOT NULL,
      "required" boolean NOT NULL,
      "default_value" text,
      "validation_json" jsonb NOT NULL,
      "sort_order" integer NOT NULL,
      CONSTRAINT "PK_field_definitions" PRIMARY KEY ("id"),
      CONSTRAINT "FK_field_definitions_version" FOREIGN KEY ("template_version_id") REFERENCES "template_versions"("id") ON DELETE CASCADE
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_field_definitions_version_id" ON "field_definitions" ("template_version_id")`,
    );

    await queryRunner.query(`CREATE TABLE "data_mappings" (
      "id" uuid NOT NULL,
      "template_version_id" uuid NOT NULL,
      "field_key" character varying(64) NOT NULL,
      "source_kind" character varying(32) NOT NULL,
      "source_path" character varying(200) NOT NULL,
      "transform" character varying(32) NOT NULL,
      CONSTRAINT "PK_data_mappings" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_data_mappings_source" CHECK ("source_kind" IN ('form', 'entity', 'constant')),
      CONSTRAINT "CHK_data_mappings_transform" CHECK ("transform" IN ('none', 'trim', 'upper', 'lower', 'date_iso')),
      CONSTRAINT "FK_data_mappings_version" FOREIGN KEY ("template_version_id") REFERENCES "template_versions"("id") ON DELETE CASCADE
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_data_mappings_version_id" ON "data_mappings" ("template_version_id")`,
    );

    const now = new Date().toISOString();
    for (const profile of SYSTEM_PROFILES) {
      await queryRunner.query(
        `INSERT INTO "profiles"
          ("id", "code", "name", "host", "scope", "workspace_id", "owner_user_id", "status", "current_version", "created_by", "created_at", "updated_at")
         VALUES ($1, $2, $3, $4, 'system', NULL, NULL, 'approved', 1, $5, $6, $6)`,
        [profile.id, profile.code, profile.name, profile.host, SYSTEM_ACTOR, now],
      );
      await queryRunner.query(
        `INSERT INTO "profile_versions"
          ("id", "profile_id", "version", "status", "source", "effective_at", "author_id", "reviewer_id", "change_summary", "immutable", "created_at", "published_at")
         VALUES ($1, $2, 1, 'approved', $3, $4, $5, $5, 'Seed hệ thống', true, $4, $4)`,
        [profile.versionId, profile.id, profile.source, now, SYSTEM_ACTOR],
      );
      for (const rule of profile.rules) {
        await queryRunner.query(
          `INSERT INTO "profile_rules"
            ("id", "profile_version_id", "rule_code", "rule_version", "default_severity", "title", "description", "check_kind", "check_config", "sort_order")
           VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9::jsonb, $10)`,
          [
            rule.id,
            profile.versionId,
            rule.code,
            rule.version,
            rule.defaultSeverity,
            rule.title,
            rule.description,
            rule.checkKind,
            JSON.stringify(rule.checkConfig),
            rule.sortOrder,
          ],
        );
      }
      for (const ref of profile.legalReferences) {
        await queryRunner.query(
          `INSERT INTO "legal_references"
            ("id", "profile_version_id", "citation", "source_url", "verified", "notes")
           VALUES ($1, $2, $3, $4, true, NULL)`,
          [ref.id, profile.versionId, ref.citation, ref.sourceUrl],
        );
        await queryRunner.query(
          `INSERT INTO "verification_evidence"
            ("id", "legal_reference_id", "recorded_by", "recorded_at", "evidence_type", "summary", "reference_uri")
           VALUES ($1, $2, $3, $4, 'statement', $5, NULL)`,
          [ref.evidenceId, ref.id, SYSTEM_ACTOR, now, ref.evidenceSummary],
        );
      }
    }

    for (const template of SYSTEM_TEMPLATES) {
      await queryRunner.query(
        `INSERT INTO "templates"
          ("id", "code", "name", "host", "scope", "workspace_id", "status", "current_version", "created_by", "created_at", "updated_at")
         VALUES ($1, $2, $3, $4, 'system', NULL, 'published', 1, $5, $6, $6)`,
        [template.id, template.code, template.name, template.host, SYSTEM_ACTOR, now],
      );
      await queryRunner.query(
        `INSERT INTO "template_versions"
          ("id", "template_id", "version", "status", "author_id", "reviewer_id", "change_summary", "immutable", "created_at", "published_at")
         VALUES ($1, $2, 1, 'published', $3, $3, 'Seed hệ thống', true, $4, $4)`,
        [template.versionId, template.id, SYSTEM_ACTOR, now],
      );
      for (const field of template.fields) {
        await queryRunner.query(
          `INSERT INTO "field_definitions"
            ("id", "template_version_id", "field_key", "label", "field_type", "required", "default_value", "validation_json", "sort_order")
           VALUES ($1, $2, $3, $4, $5, $6, $7, $8::jsonb, $9)`,
          [
            field.id,
            template.versionId,
            field.fieldKey,
            field.label,
            field.fieldType,
            field.required,
            field.defaultValue,
            JSON.stringify(field.validationJson),
            field.sortOrder,
          ],
        );
      }
      for (const mapping of template.mappings) {
        await queryRunner.query(
          `INSERT INTO "data_mappings"
            ("id", "template_version_id", "field_key", "source_kind", "source_path", "transform")
           VALUES ($1, $2, $3, $4, $5, $6)`,
          [
            mapping.id,
            template.versionId,
            mapping.fieldKey,
            mapping.sourceKind,
            mapping.sourcePath,
            mapping.transform,
          ],
        );
      }
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(`DROP TABLE IF EXISTS "data_mappings"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "field_definitions"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "template_versions"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "templates"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "verification_evidence"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "legal_references"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "profile_rules"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "profile_versions"`);
    await queryRunner.query(`DROP TABLE IF EXISTS "profiles"`);

    const permissionId = new Map(PERMISSIONS.map((item) => [item.code, item.id]));
    for (const role of ROLES) {
      for (const code of PROFILE_PERMISSION_CODES) {
        if (!(ROLE_PERMISSIONS[role.code] as readonly string[]).includes(code)) {
          continue;
        }
        await queryRunner.query(
          `DELETE FROM "role_permissions" WHERE "role_id" = $1 AND "permission_id" = $2`,
          [role.id, permissionId.get(code)],
        );
      }
    }
    for (const code of PROFILE_PERMISSION_CODES) {
      await queryRunner.query(`DELETE FROM "permissions" WHERE "id" = $1`, [
        permissionId.get(code),
      ]);
    }
  }
}
