import { Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'sessions' })
@Index('IDX_sessions_user_id', ['userId'])
export class Session {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'user_id', type: 'uuid' })
  userId!: string;

  @Column({ name: 'secret_hash', type: 'varchar', length: 128 })
  secretHash!: string;

  @Column({ name: 'active_workspace_id', type: 'uuid' })
  activeWorkspaceId!: string;

  @Column({ name: 'user_agent', type: 'varchar', length: 256, nullable: true })
  userAgent: string | null = null;

  @Column({ name: 'expires_at', type: 'timestamptz' })
  expiresAt!: Date;

  @Column({ name: 'revoked_at', type: 'timestamptz', nullable: true })
  revokedAt: Date | null = null;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'rotated_at', type: 'timestamptz' })
  rotatedAt!: Date;
}
