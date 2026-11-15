import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'invoices' })
@Index('IDX_invoices_workspace_id', ['workspaceId'])
@Index('UQ_invoices_number', ['invoiceNumber'], { unique: true })
@Check(`"status" IN ('issued', 'void')`)
export class Invoice {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'order_id', type: 'uuid' })
  orderId!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ name: 'invoice_number', type: 'varchar', length: 64 })
  invoiceNumber!: string;

  @Column({ type: 'varchar', length: 32 })
  status!: 'issued' | 'void';

  @Column({ name: 'amount_minor', type: 'int' })
  amountMinor!: number;

  @Column({ name: 'issued_at', type: 'timestamptz' })
  issuedAt!: Date;
}
