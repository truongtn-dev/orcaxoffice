import { Check, Column, Entity, Index, PrimaryColumn, Unique } from 'typeorm';

export type MembershipStatus = 'active' | 'suspended' | 'removed';

@Entity({ name: 'memberships' })
@Unique('UQ_memberships_workspace_user', ['workspaceId', 'userId'])
@Index('IDX_memberships_user_id', ['userId'])
@Check('CHK_memberships_status', `"status" IN ('active', 'suspended', 'removed')`)
export class Membership {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ name: 'user_id', type: 'uuid' })
  userId!: string;

  @Column({ name: 'role_id', type: 'uuid' })
  roleId!: string;

  @Column({ type: 'varchar', length: 32 })
  status!: MembershipStatus;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
