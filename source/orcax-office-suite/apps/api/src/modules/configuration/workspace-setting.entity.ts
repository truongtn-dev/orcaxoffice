import { Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'workspace_settings' })
@Index('UQ_workspace_settings_workspace_key', ['workspaceId', 'key'], { unique: true })
export class WorkspaceSetting {
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

  @Column({ name: 'updated_by', type: 'uuid' })
  updatedBy!: string;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
