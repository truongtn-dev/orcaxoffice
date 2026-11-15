import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'licenses' })
@Index('IDX_licenses_workspace_id', ['workspaceId'])
@Check('CHK_licenses_status', `"status" IN ('active', 'revoked')`)
export class License {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ name: 'plan_id', type: 'uuid' })
  planId!: string;

  @Column({ type: 'varchar', length: 32 })
  status!: 'active' | 'revoked';

  @Column({ name: 'issued_by', type: 'uuid' })
  issuedBy!: string;

  @Column({ name: 'starts_at', type: 'timestamptz' })
  startsAt!: Date;

  @Column({ name: 'ends_at', type: 'timestamptz', nullable: true })
  endsAt: Date | null = null;

  @Column({ name: 'revoked_at', type: 'timestamptz', nullable: true })
  revokedAt: Date | null = null;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
