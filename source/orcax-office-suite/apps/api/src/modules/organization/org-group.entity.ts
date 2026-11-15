import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'org_groups' })
@Index('IDX_org_groups_workspace_id', ['workspaceId'])
@Check(`"status" IN ('active', 'archived')`)
export class OrgGroup {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ type: 'varchar', length: 64 })
  code!: string;

  @Column({ type: 'varchar', length: 160 })
  name!: string;

  @Column({ type: 'varchar', length: 32 })
  status!: 'active' | 'archived';

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
