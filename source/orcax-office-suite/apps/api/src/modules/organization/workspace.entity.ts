import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

export type WorkspaceKind = 'personal' | 'organization';
export type WorkspaceStatus = 'active' | 'suspended' | 'archived';

@Entity({ name: 'workspaces' })
@Index('UQ_workspaces_personal_owner', ['ownerUserId'], {
  unique: true,
  where: `"kind" = 'personal'`,
})
@Check('CHK_workspaces_status', `"status" IN ('active', 'suspended', 'archived')`)
@Check(
  'CHK_workspaces_kind_owner',
  `("kind" = 'personal' AND "owner_user_id" IS NOT NULL AND "organization_id" IS NULL) OR ("kind" = 'organization' AND "owner_user_id" IS NULL AND "organization_id" IS NOT NULL)`,
)
@Check('CHK_workspaces_personal_active', `"kind" <> 'personal' OR "status" = 'active'`)
export class Workspace {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ type: 'varchar', length: 32 })
  kind!: WorkspaceKind;

  @Column({ type: 'varchar', length: 160 })
  name!: string;

  @Column({ type: 'varchar', length: 32 })
  status!: WorkspaceStatus;

  @Column({ name: 'owner_user_id', type: 'uuid', nullable: true })
  ownerUserId: string | null = null;

  @Column({ name: 'organization_id', type: 'uuid', nullable: true })
  organizationId: string | null = null;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
