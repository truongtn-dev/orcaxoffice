import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'templates' })
@Index('IDX_templates_workspace_id', ['workspaceId'])
@Check(`"host" IN ('word', 'excel')`)
@Check(`"scope" IN ('system', 'personal', 'organization')`)
@Check(`"status" IN ('draft', 'review', 'approved', 'published', 'deprecated')`)
export class Template {
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

  @Column({ type: 'varchar', length: 32 })
  status!: 'draft' | 'review' | 'approved' | 'published' | 'deprecated';

  @Column({ name: 'current_version', type: 'int' })
  currentVersion!: number;

  @Column({ name: 'created_by', type: 'uuid' })
  createdBy!: string;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
