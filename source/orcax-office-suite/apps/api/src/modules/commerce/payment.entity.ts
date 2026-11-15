import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'payments' })
@Index('IDX_payments_workspace_id', ['workspaceId'])
@Index('UQ_payments_idempotency_key', ['idempotencyKey'], { unique: true })
@Check(`"status" IN ('pending', 'succeeded', 'failed')`)
export class Payment {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'order_id', type: 'uuid' })
  orderId!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ type: 'varchar', length: 32 })
  provider!: string;

  @Column({ name: 'provider_ref', type: 'varchar', length: 160, nullable: true })
  providerRef: string | null = null;

  @Column({ name: 'idempotency_key', type: 'varchar', length: 128 })
  idempotencyKey!: string;

  @Column({ type: 'varchar', length: 32 })
  status!: 'pending' | 'succeeded' | 'failed';

  @Column({ name: 'amount_minor', type: 'int' })
  amountMinor!: number;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
