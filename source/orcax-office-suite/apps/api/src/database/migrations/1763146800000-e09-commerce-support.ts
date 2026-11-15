import { MigrationInterface, QueryRunner } from 'typeorm';
import { PERMISSIONS, ROLE_PERMISSIONS, ROLES } from '../../modules/authorization/catalog';
import { PLAN_PRICES } from '../../modules/commerce/commerce-catalog';
import { INITIAL_HELP_ARTICLES } from '../../modules/support/support-catalog';

const E09_PERMISSION_CODES = [
  'commerce.order.read',
  'commerce.order.create',
  'commerce.refund',
  'commerce.reconcile',
  'support.ticket.read',
  'support.ticket.create',
  'support.ticket.manage',
  'support.help.read',
  'support.help.manage',
] as const;

export class E09CommerceSupport1763146800000 implements MigrationInterface {
  name = 'E09CommerceSupport1763146800000';

  public async up(queryRunner: QueryRunner): Promise<void> {
    for (const permission of PERMISSIONS.filter((item) =>
      (E09_PERMISSION_CODES as readonly string[]).includes(item.code),
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
        if (!(E09_PERMISSION_CODES as readonly string[]).includes(code)) {
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

    await queryRunner.query(`CREATE TABLE "plan_prices" (
      "id" uuid NOT NULL,
      "plan_id" uuid NOT NULL,
      "currency" character varying(8) NOT NULL,
      "amount_minor" integer NOT NULL,
      "billing_period" character varying(32) NOT NULL,
      "status" character varying(32) NOT NULL DEFAULT 'active',
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_plan_prices" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_plan_prices_billing_period" CHECK ("billing_period" IN ('once', 'month', 'year')),
      CONSTRAINT "CHK_plan_prices_status" CHECK ("status" IN ('active', 'archived')),
      CONSTRAINT "FK_plan_prices_plan" FOREIGN KEY ("plan_id") REFERENCES "plans"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE UNIQUE INDEX "UQ_plan_prices_plan_currency_period" ON "plan_prices" ("plan_id", "currency", "billing_period")`,
    );

    await queryRunner.query(`CREATE TABLE "orders" (
      "id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "status" character varying(32) NOT NULL,
      "currency" character varying(8) NOT NULL,
      "total_minor" integer NOT NULL,
      "created_by" uuid NOT NULL,
      "finalized_at" TIMESTAMP WITH TIME ZONE,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_orders" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_orders_status" CHECK ("status" IN ('draft', 'pending_payment', 'paid', 'cancelled', 'refunded', 'partially_refunded')),
      CONSTRAINT "FK_orders_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_orders_user" FOREIGN KEY ("created_by") REFERENCES "users"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(`CREATE INDEX "IDX_orders_workspace_id" ON "orders" ("workspace_id")`);

    await queryRunner.query(`CREATE TABLE "order_items" (
      "id" uuid NOT NULL,
      "order_id" uuid NOT NULL,
      "plan_id" uuid NOT NULL,
      "quantity" integer NOT NULL,
      "unit_amount_minor" integer NOT NULL,
      "line_total_minor" integer NOT NULL,
      CONSTRAINT "PK_order_items" PRIMARY KEY ("id"),
      CONSTRAINT "FK_order_items_order" FOREIGN KEY ("order_id") REFERENCES "orders"("id") ON DELETE CASCADE,
      CONSTRAINT "FK_order_items_plan" FOREIGN KEY ("plan_id") REFERENCES "plans"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_order_items_order_id" ON "order_items" ("order_id")`,
    );

    await queryRunner.query(`CREATE TABLE "payments" (
      "id" uuid NOT NULL,
      "order_id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "provider" character varying(64) NOT NULL,
      "provider_ref" character varying(160),
      "idempotency_key" character varying(128) NOT NULL,
      "status" character varying(32) NOT NULL,
      "amount_minor" integer NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_payments" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_payments_status" CHECK ("status" IN ('pending', 'succeeded', 'failed', 'refunded')),
      CONSTRAINT "UQ_payments_idempotency_key" UNIQUE ("idempotency_key"),
      CONSTRAINT "FK_payments_order" FOREIGN KEY ("order_id") REFERENCES "orders"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_payments_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_payments_workspace_id" ON "payments" ("workspace_id")`,
    );
    await queryRunner.query(`CREATE INDEX "IDX_payments_order_id" ON "payments" ("order_id")`);

    await queryRunner.query(`CREATE TABLE "refunds" (
      "id" uuid NOT NULL,
      "payment_id" uuid NOT NULL,
      "order_id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "amount_minor" integer NOT NULL,
      "status" character varying(32) NOT NULL,
      "provider_ref" character varying(160),
      "reason" character varying(255) NOT NULL,
      "created_by" uuid NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_refunds" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_refunds_status" CHECK ("status" IN ('pending', 'succeeded', 'failed')),
      CONSTRAINT "FK_refunds_payment" FOREIGN KEY ("payment_id") REFERENCES "payments"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_refunds_order" FOREIGN KEY ("order_id") REFERENCES "orders"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_refunds_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_refunds_user" FOREIGN KEY ("created_by") REFERENCES "users"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_refunds_workspace_id" ON "refunds" ("workspace_id")`,
    );

    await queryRunner.query(`CREATE TABLE "invoices" (
      "id" uuid NOT NULL,
      "order_id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "invoice_number" character varying(64) NOT NULL,
      "status" character varying(32) NOT NULL,
      "amount_minor" integer NOT NULL,
      "issued_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_invoices" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_invoices_status" CHECK ("status" IN ('draft', 'issued', 'void')),
      CONSTRAINT "UQ_invoices_number" UNIQUE ("invoice_number"),
      CONSTRAINT "FK_invoices_order" FOREIGN KEY ("order_id") REFERENCES "orders"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_invoices_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_invoices_workspace_id" ON "invoices" ("workspace_id")`,
    );

    await queryRunner.query(`CREATE TABLE "reconciliation_items" (
      "id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "payment_id" uuid NOT NULL,
      "status" character varying(32) NOT NULL,
      "note" text NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_reconciliation_items" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_reconciliation_items_status" CHECK ("status" IN ('unmatched', 'matched', 'ignored')),
      CONSTRAINT "FK_reconciliation_items_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_reconciliation_items_payment" FOREIGN KEY ("payment_id") REFERENCES "payments"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_reconciliation_items_workspace_id" ON "reconciliation_items" ("workspace_id")`,
    );

    await queryRunner.query(`CREATE TABLE "support_tickets" (
      "id" uuid NOT NULL,
      "workspace_id" uuid NOT NULL,
      "user_id" uuid NOT NULL,
      "title" character varying(255) NOT NULL,
      "category" character varying(64) NOT NULL,
      "priority" character varying(32) NOT NULL,
      "status" character varying(32) NOT NULL,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_support_tickets" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_support_tickets_category" CHECK ("category" IN ('general', 'account', 'billing', 'licensing', 'word_addin', 'excel_addin')),
      CONSTRAINT "CHK_support_tickets_priority" CHECK ("priority" IN ('low', 'medium', 'high', 'urgent')),
      CONSTRAINT "CHK_support_tickets_status" CHECK ("status" IN ('open', 'in_progress', 'resolved', 'closed')),
      CONSTRAINT "FK_support_tickets_workspace" FOREIGN KEY ("workspace_id") REFERENCES "workspaces"("id") ON DELETE RESTRICT,
      CONSTRAINT "FK_support_tickets_user" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_support_tickets_workspace_id" ON "support_tickets" ("workspace_id")`,
    );

    await queryRunner.query(`CREATE TABLE "support_ticket_comments" (
      "id" uuid NOT NULL,
      "ticket_id" uuid NOT NULL,
      "user_id" uuid NOT NULL,
      "message" text NOT NULL,
      "is_staff" boolean NOT NULL DEFAULT false,
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_support_ticket_comments" PRIMARY KEY ("id"),
      CONSTRAINT "FK_support_ticket_comments_ticket" FOREIGN KEY ("ticket_id") REFERENCES "support_tickets"("id") ON DELETE CASCADE,
      CONSTRAINT "FK_support_ticket_comments_user" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE RESTRICT
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_support_ticket_comments_ticket_id" ON "support_ticket_comments" ("ticket_id")`,
    );

    await queryRunner.query(`CREATE TABLE "help_articles" (
      "id" uuid NOT NULL,
      "slug" character varying(128) NOT NULL,
      "title" character varying(255) NOT NULL,
      "category" character varying(64) NOT NULL,
      "content" text NOT NULL,
      "product_version" character varying(32) NOT NULL DEFAULT '1.0',
      "status" character varying(32) NOT NULL DEFAULT 'published',
      "created_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      "updated_at" TIMESTAMP WITH TIME ZONE NOT NULL,
      CONSTRAINT "PK_help_articles" PRIMARY KEY ("id"),
      CONSTRAINT "CHK_help_articles_status" CHECK ("status" IN ('draft', 'published')),
      CONSTRAINT "UQ_help_articles_slug" UNIQUE ("slug")
    )`);
    await queryRunner.query(
      `CREATE INDEX "IDX_help_articles_category" ON "help_articles" ("category")`,
    );

    const now = new Date().toISOString();
    for (const price of PLAN_PRICES) {
      await queryRunner.query(
        `INSERT INTO "plan_prices" ("id", "plan_id", "currency", "amount_minor", "billing_period", "status", "created_at", "updated_at")
         SELECT $1, $2, $3, $4, $5, 'active', $6, $6
         WHERE NOT EXISTS (SELECT 1 FROM "plan_prices" WHERE "id" = $1)`,
        [price.id, price.planId, price.currency, price.amountMinor, price.billingPeriod, now],
      );
    }

    for (const article of INITIAL_HELP_ARTICLES) {
      await queryRunner.query(
        `INSERT INTO "help_articles" ("id", "slug", "title", "category", "content", "product_version", "status", "created_at", "updated_at")
         SELECT $1, $2, $3, $4, $5, $6, $7, $8, $8
         WHERE NOT EXISTS (SELECT 1 FROM "help_articles" WHERE "id" = $1 OR "slug" = $2)`,
        [
          article.id,
          article.slug,
          article.title,
          article.category,
          article.content,
          article.productVersion,
          article.status,
          now,
        ],
      );
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(`DROP TABLE IF EXISTS "help_articles" CASCADE`);
    await queryRunner.query(`DROP TABLE IF EXISTS "support_ticket_comments" CASCADE`);
    await queryRunner.query(`DROP TABLE IF EXISTS "support_tickets" CASCADE`);
    await queryRunner.query(`DROP TABLE IF EXISTS "reconciliation_items" CASCADE`);
    await queryRunner.query(`DROP TABLE IF EXISTS "invoices" CASCADE`);
    await queryRunner.query(`DROP TABLE IF EXISTS "refunds" CASCADE`);
    await queryRunner.query(`DROP TABLE IF EXISTS "payments" CASCADE`);
    await queryRunner.query(`DROP TABLE IF EXISTS "order_items" CASCADE`);
    await queryRunner.query(`DROP TABLE IF EXISTS "orders" CASCADE`);
    await queryRunner.query(`DROP TABLE IF EXISTS "plan_prices" CASCADE`);

    for (const code of E09_PERMISSION_CODES) {
      await queryRunner.query(
        `DELETE FROM "role_permissions" WHERE "permission_id" IN (SELECT "id" FROM "permissions" WHERE "code" = $1)`,
        [code],
      );
      await queryRunner.query(`DELETE FROM "permissions" WHERE "code" = $1`, [code]);
    }
  }
}
