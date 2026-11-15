import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

export type InvitationStatus = 'invited' | 'accepted' | 'revoked' | 'expired';

@Entity({ name: 'invitations' })
@Index('UQ_invitations_token_hash', ['tokenHash'], { unique: true })
@Check('CHK_invitations_status', `"status" IN ('invited', 'accepted', 'revoked', 'expired')`)
export class Invitation {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ type: 'varchar', length: 320 })
  email!: string;

  @Column({ name: 'role_id', type: 'uuid' })
  roleId!: string;

  @Column({ name: 'token_hash', type: 'varchar', length: 128 })
  tokenHash!: string;

  @Column({ type: 'varchar', length: 32 })
  status!: InvitationStatus;

  @Column({ name: 'invited_by', type: 'uuid' })
  invitedBy!: string;

  @Column({ name: 'expires_at', type: 'timestamptz' })
  expiresAt!: Date;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;
}
