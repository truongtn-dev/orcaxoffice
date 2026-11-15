import { Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'support_ticket_comments' })
@Index('IDX_support_ticket_comments_ticket_id', ['ticketId'])
export class SupportComment {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'ticket_id', type: 'uuid' })
  ticketId!: string;

  @Column({ name: 'user_id', type: 'uuid' })
  userId!: string;

  @Column({ type: 'text' })
  message!: string;

  @Column({ name: 'is_staff', type: 'boolean', default: false })
  isStaff: boolean = false;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;
}
