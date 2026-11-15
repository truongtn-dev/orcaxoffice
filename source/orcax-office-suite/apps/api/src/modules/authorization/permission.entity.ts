import { Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'permissions' })
@Index('UQ_permissions_code', ['code'], { unique: true })
export class Permission {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ type: 'varchar', length: 64 })
  code!: string;
}
