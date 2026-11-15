import { Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'print_runs' })
@Index('IDX_print_runs_workspace_id', ['workspaceId'])
export class PrintRun {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ name: 'user_id', type: 'uuid' })
  userId!: string;

  @Column({ name: 'profile_id', type: 'uuid' })
  profileId!: string;

  @Column({ name: 'profile_version_id', type: 'uuid' })
  profileVersionId!: string;

  @Column({ name: 'profile_version', type: 'int' })
  profileVersion!: number;

  @Column({ name: 'profile_code', type: 'varchar', length: 64 })
  profileCode!: string;

  @Column({ name: 'paper_size', type: 'varchar', length: 32 })
  paperSize!: string;

  @Column({ type: 'varchar', length: 32 })
  orientation!: string;

  @Column({ name: 'applied', type: 'boolean' })
  applied!: boolean;

  @Column({ name: 'pdf_exported', type: 'boolean' })
  pdfExported!: boolean;

  @Column({ name: 'adapter', type: 'varchar', length: 32 })
  adapter!: string;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;
}
