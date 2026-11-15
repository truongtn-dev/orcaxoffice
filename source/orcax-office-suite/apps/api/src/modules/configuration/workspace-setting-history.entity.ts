import { Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'workspace_setting_history' })
@Index('IDX_workspace_setting_history_workspace_id', ['workspaceId'])
export class WorkspaceSettingHistory {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ type: 'varchar', length: 64 })
  key!: string;

  @Column({ type: 'jsonb' })
  value!: object;

  @Column({ type: 'int' })
  version!: number;

  @Column({ name: 'changed_by', type: 'uuid' })
  changedBy!: string;

  @Column({ name: 'changed_at', type: 'timestamptz' })
  changedAt!: Date;
}
