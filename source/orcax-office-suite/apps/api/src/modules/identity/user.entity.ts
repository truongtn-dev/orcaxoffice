import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

export type UserStatus = 'pending' | 'active' | 'suspended';

@Entity({ name: 'users' })
@Index('UQ_users_email', ['email'], { unique: true })
@Check('CHK_users_status', `"status" IN ('pending', 'active', 'suspended')`)
export class User {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ type: 'varchar', length: 320 })
  email!: string;

  @Column({ name: 'password_hash', type: 'varchar', length: 255 })
  passwordHash!: string;

  @Column({ name: 'display_name', type: 'varchar', length: 120 })
  displayName!: string;

  @Column({ type: 'varchar', length: 32 })
  status!: UserStatus;

  @Column({ name: 'email_verified_at', type: 'timestamptz', nullable: true })
  emailVerifiedAt: Date | null = null;

  @Column({ name: 'last_workspace_id', type: 'uuid', nullable: true })
  lastWorkspaceId: string | null = null;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
