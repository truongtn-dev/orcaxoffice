import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'validation_runs' })
@Index('IDX_validation_runs_workspace_id', ['workspaceId'])
@Check(`"host" IN ('word', 'excel')`)
export class ValidationRun {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'workspace_id', type: 'uuid' })
  workspaceId!: string;

  @Column({ name: 'user_id', type: 'uuid' })
  userId!: string;

  @Column({ type: 'varchar', length: 16 })
  host!: 'word' | 'excel';

  @Column({ name: 'profile_id', type: 'uuid' })
  profileId!: string;

  @Column({ name: 'profile_version_id', type: 'uuid' })
  profileVersionId!: string;

  @Column({ name: 'profile_version', type: 'int' })
  profileVersion!: number;

  @Column({ name: 'profile_code', type: 'varchar', length: 64 })
  profileCode!: string;

  @Column({ name: 'summary_pass', type: 'int' })
  summaryPass!: number;

  @Column({ name: 'summary_fail', type: 'int' })
  summaryFail!: number;

  @Column({ name: 'summary_warning', type: 'int' })
  summaryWarning!: number;

  @Column({ name: 'summary_corrected', type: 'int' })
  summaryCorrected!: number;

  @Column({ name: 'adapter', type: 'varchar', length: 32 })
  adapter!: string;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;
}
