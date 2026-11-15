import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'profile_rules' })
@Index('IDX_profile_rules_version_id', ['profileVersionId'])
@Check(`"default_severity" IN ('fail', 'warning')`)
export class ProfileRule {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'profile_version_id', type: 'uuid' })
  profileVersionId!: string;

  @Column({ name: 'rule_code', type: 'varchar', length: 64 })
  ruleCode!: string;

  @Column({ name: 'rule_version', type: 'varchar', length: 32 })
  ruleVersion!: string;

  @Column({ name: 'default_severity', type: 'varchar', length: 16 })
  defaultSeverity!: 'fail' | 'warning';

  @Column({ type: 'varchar', length: 200 })
  title!: string;

  @Column({ type: 'text' })
  description!: string;

  @Column({ name: 'check_kind', type: 'varchar', length: 64 })
  checkKind!: string;

  @Column({ name: 'check_config', type: 'jsonb' })
  checkConfig!: object;

  @Column({ name: 'sort_order', type: 'int' })
  sortOrder!: number;
}
