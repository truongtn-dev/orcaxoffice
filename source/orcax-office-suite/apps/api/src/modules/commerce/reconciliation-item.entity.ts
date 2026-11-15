import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'reconciliation_items' })
@Index('IDX_reconciliation_items_workspace_id', ['workspaceId'])
@Check(`"status" IN ('unmatched', 'matched', 'ignored')`)
export class ReconciliationItem {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ name: 'payment_id', type: 'uuid', nullable: true })
  paymentId: string | null = null;

  @Column({ type: 'varchar', length: 32 })
  status!: 'unmatched' | 'matched' | 'ignored';

  @Column({ type: 'text' })
  note!: string;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
