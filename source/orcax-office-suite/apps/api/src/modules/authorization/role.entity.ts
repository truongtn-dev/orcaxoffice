import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'roles' })
@Index('UQ_roles_code', ['code'], { unique: true })
@Check('CHK_roles_kind', `"kind" IN ('personal', 'organization')`)
export class Role {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ type: 'varchar', length: 64 })
  code!: string;

  @Column({ type: 'varchar', length: 32 })
  kind!: 'personal' | 'organization';
}
