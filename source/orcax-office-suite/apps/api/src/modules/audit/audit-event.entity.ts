import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'audit_events' })
@Index('IDX_audit_events_workspace_occurred_at', ['workspaceId', 'occurredAt'])
@Index('IDX_audit_events_correlation_id', ['correlationId'])
@Check('CHK_audit_events_outcome', `"outcome" IN ('success', 'failure')`)
export class AuditEvent {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'workspace_id', type: 'uuid', nullable: true })
  workspaceId: string | null = null;

  @Column({ name: 'actor_id', type: 'uuid', nullable: true })
  actorId: string | null = null;

  @Column({ type: 'varchar', length: 128 })
  action!: string;

  @Column({ name: 'target_type', type: 'varchar', length: 128 })
  targetType!: string;

  @Column({ name: 'target_id', type: 'varchar', length: 128, nullable: true })
  targetId: string | null = null;

  @Column({ name: 'occurred_at', type: 'timestamptz' })
  occurredAt!: Date;

  @Column({ type: 'varchar', length: 32 })
  outcome!: 'success' | 'failure';

  @Column({ name: 'correlation_id', type: 'uuid' })
  correlationId!: string;

  @Column({ type: 'jsonb', nullable: true })
  metadata: Record<string, unknown> | null = null;
}
