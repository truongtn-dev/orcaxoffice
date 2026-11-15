import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'orders' })
@Index('IDX_orders_workspace_id', ['workspaceId'])
@Check(
  `"status" IN ('draft', 'pending_payment', 'paid', 'cancelled', 'refunded', 'partially_refunded')`,
)
export class Order {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ type: 'varchar', length: 32 })
  status!: 'draft' | 'pending_payment' | 'paid' | 'cancelled' | 'refunded' | 'partially_refunded';

  @Column({ type: 'varchar', length: 8 })
  currency!: string;

  @Column({ name: 'total_minor', type: 'int' })
  totalMinor!: number;

  @Column({ name: 'created_by', type: 'uuid' })
  createdBy!: string;

  @Column({ name: 'finalized_at', type: 'timestamptz', nullable: true })
  finalizedAt: Date | null = null;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
