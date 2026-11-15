import { Check, Column, Entity, PrimaryColumn } from 'typeorm';

export type OrganizationStatus = 'active' | 'suspended' | 'archived';

@Entity({ name: 'organizations' })
@Check('CHK_organizations_status', `"status" IN ('active', 'suspended', 'archived')`)
export class Organization {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ type: 'varchar', length: 160 })
  name!: string;

  @Column({ type: 'varchar', length: 32 })
  status!: OrganizationStatus;

  @Column({ name: 'created_by', type: 'uuid' })
  createdBy!: string;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
