import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

export type TicketCategory =
  'general' | 'account' | 'billing' | 'licensing' | 'word_addin' | 'excel_addin';

export type TicketPriority = 'low' | 'medium' | 'high' | 'urgent';

export type TicketStatus = 'open' | 'in_progress' | 'resolved' | 'closed';

@Entity({ name: 'support_tickets' })
@Index('IDX_support_tickets_workspace_id', ['workspaceId'])
@Check(`"category" IN ('general', 'account', 'billing', 'licensing', 'word_addin', 'excel_addin')`)
@Check(`"priority" IN ('low', 'medium', 'high', 'urgent')`)
@Check(`"status" IN ('open', 'in_progress', 'resolved', 'closed')`)
export class SupportTicket {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ name: 'user_id', type: 'uuid' })
  userId!: string;

  @Column({ type: 'varchar', length: 255 })
  title!: string;

  @Column({ type: 'varchar', length: 64 })
  category!: TicketCategory;

  @Column({ type: 'varchar', length: 32 })
  priority!: TicketPriority;

  @Column({ type: 'varchar', length: 32 })
  status!: TicketStatus;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
