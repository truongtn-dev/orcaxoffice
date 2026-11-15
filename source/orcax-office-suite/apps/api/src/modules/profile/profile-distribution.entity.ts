import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'profile_distributions' })
@Index('IDX_profile_distributions_profile_id', ['profileId'])
@Index('IDX_profile_distributions_workspace_id', ['workspaceId'])
@Check(`"target_type" IN ('department', 'group')`)
export class ProfileDistribution {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'profile_id', type: 'uuid' })
  profileId!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ name: 'target_type', type: 'varchar', length: 32 })
  targetType!: 'department' | 'group';

  @Column({ name: 'target_id', type: 'uuid' })
  targetId!: string;

  @Column({ name: 'distributed_by', type: 'uuid' })
  distributedBy!: string;

  @Column({ name: 'distributed_at', type: 'timestamptz' })
  distributedAt!: Date;
}
