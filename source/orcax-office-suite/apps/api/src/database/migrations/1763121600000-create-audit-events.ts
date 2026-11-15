import { MigrationInterface, QueryRunner } from 'typeorm';

const UP_STATEMENTS = [
  `CREATE TABLE "audit_events" (
    "id" uuid NOT NULL DEFAULT gen_random_uuid(),
    "workspace_id" uuid,
    "actor_id" uuid,
    "action" character varying(128) NOT NULL,
    "target_type" character varying(128) NOT NULL,
    "target_id" character varying(128),
    "occurred_at" TIMESTAMP WITH TIME ZONE NOT NULL,
    "outcome" character varying(32) NOT NULL,
    "correlation_id" uuid NOT NULL,
    "metadata" jsonb,
    CONSTRAINT "PK_audit_events" PRIMARY KEY ("id"),
    CONSTRAINT "CHK_audit_events_outcome" CHECK ("outcome" IN ('success', 'failure'))
  )`,
  `COMMENT ON TABLE "audit_events" IS 'Append-only audit foundation. Do not store passwords, tokens, payment credentials, or Office document bodies.'`,
  `CREATE INDEX "IDX_audit_events_workspace_occurred_at" ON "audit_events" ("workspace_id", "occurred_at")`,
  `CREATE INDEX "IDX_audit_events_correlation_id" ON "audit_events" ("correlation_id")`,
  `CREATE OR REPLACE FUNCTION audit_events_append_only()
    RETURNS trigger
    LANGUAGE plpgsql
    AS $function$
    BEGIN
      RAISE EXCEPTION 'audit_events is append-only';
    END;
    $function$`,
  `CREATE TRIGGER "audit_events_no_mutation"
    BEFORE UPDATE OR DELETE ON "audit_events"
    FOR EACH ROW
    EXECUTE FUNCTION audit_events_append_only()`,
  `CREATE TRIGGER "audit_events_no_truncate"
    BEFORE TRUNCATE ON "audit_events"
    FOR EACH STATEMENT
    EXECUTE FUNCTION audit_events_append_only()`,
];

const DOWN_STATEMENTS = [
  `DROP TRIGGER IF EXISTS "audit_events_no_truncate" ON "audit_events"`,
  `DROP TRIGGER IF EXISTS "audit_events_no_mutation" ON "audit_events"`,
  `DROP TABLE IF EXISTS "audit_events"`,
  `DROP FUNCTION IF EXISTS audit_events_append_only()`,
];

export class CreateAuditEvents1763121600000 implements MigrationInterface {
  name = 'CreateAuditEvents1763121600000';

  public async up(queryRunner: QueryRunner): Promise<void> {
    for (const statement of UP_STATEMENTS) {
      await queryRunner.query(statement);
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    for (const statement of DOWN_STATEMENTS) {
      await queryRunner.query(statement);
    }
  }
}
