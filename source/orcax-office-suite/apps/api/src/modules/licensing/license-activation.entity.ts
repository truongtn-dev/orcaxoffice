import { Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'license_activations' })
@Index('IDX_license_activations_workspace_id', ['workspaceId'])
export class LicenseActivation {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'license_id', type: 'uuid' })
  licenseId!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ name: 'user_id', type: 'uuid' })
  userId!: string;

  @Column({ name: 'device_key_hash', type: 'varchar', length: 128 })
  deviceKeyHash!: string;

  @Column({ name: 'device_label', type: 'varchar', length: 160 })
  deviceLabel!: string;

  @Column({ name: 'activated_at', type: 'timestamptz' })
  activatedAt!: Date;

  @Column({ name: 'last_seen_at', type: 'timestamptz' })
  lastSeenAt!: Date;

  @Column({ name: 'revoked_at', type: 'timestamptz', nullable: true })
  revokedAt: Date | null = null;
}
