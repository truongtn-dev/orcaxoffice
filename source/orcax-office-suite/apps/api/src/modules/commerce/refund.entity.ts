import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'refunds' })
@Index('IDX_refunds_workspace_id', ['workspaceId'])
@Check(`"status" IN ('pending', 'succeeded', 'failed')`)
export class Refund {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'payment_id', type: 'uuid' })
  paymentId!: string;

  @Column({ name: 'order_id', type: 'uuid' })
  orderId!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ name: 'amount_minor', type: 'int' })
  amountMinor!: number;

  @Column({ type: 'varchar', length: 32 })
  status!: 'pending' | 'succeeded' | 'failed';

  @Column({ name: 'provider_ref', type: 'varchar', length: 160, nullable: true })
  providerRef: string | null = null;

  @Column({ type: 'varchar', length: 500 })
  reason!: string;

  @Column({ name: 'created_by', type: 'uuid' })
  createdBy!: string;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;
}
