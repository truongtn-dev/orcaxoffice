import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'profiles' })
@Index('IDX_profiles_workspace_id', ['workspaceId'])
@Check(`"host" IN ('word', 'excel')`)
@Check(`"scope" IN ('system', 'personal', 'organization')`)
@Check(`"status" IN ('draft', 'in_review', 'approved', 'deprecated')`)
export class Profile {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ type: 'varchar', length: 64 })
  code!: string;

  @Column({ type: 'varchar', length: 160 })
  name!: string;

  @Column({ type: 'varchar', length: 16 })
  host!: 'word' | 'excel';

  @Column({ type: 'varchar', length: 32 })
  scope!: 'system' | 'personal' | 'organization';

  @Column({ name: 'workspace_id', type: 'uuid', nullable: true })
  workspaceId: string | null = null;

  @Column({ name: 'owner_user_id', type: 'uuid', nullable: true })
  ownerUserId: string | null = null;

  @Column({ type: 'varchar', length: 32 })
  status!: 'draft' | 'in_review' | 'approved' | 'deprecated';

  @Column({ name: 'current_version', type: 'int' })
  currentVersion!: number;

  @Column({ name: 'created_by', type: 'uuid' })
  createdBy!: string;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
