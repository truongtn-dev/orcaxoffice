import { Check, Column, Entity, PrimaryColumn } from 'typeorm';

@Entity({ name: 'plans' })
@Check('CHK_plans_license_type', `"license_type" IN ('subscription', 'perpetual')`)
@Check('CHK_plans_workspace_kind', `"workspace_kind" IN ('personal', 'organization')`)
@Check('CHK_plans_status', `"status" IN ('active', 'retired')`)
export class Plan {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'product_id', type: 'uuid' })
  productId!: string;

  @Column({ type: 'varchar', length: 64 })
  code!: string;

  @Column({ type: 'varchar', length: 160 })
  name!: string;

  @Column({ name: 'license_type', type: 'varchar', length: 32 })
  licenseType!: 'subscription' | 'perpetual';

  @Column({ name: 'workspace_kind', type: 'varchar', length: 32 })
  workspaceKind!: 'personal' | 'organization';

  @Column({ name: 'device_limit', type: 'int' })
  deviceLimit!: number;

  @Column({ name: 'seat_limit', type: 'int' })
  seatLimit!: number;

  @Column({ name: 'term_days', type: 'int', nullable: true })
  termDays: number | null = null;

  @Column({ name: 'grace_days', type: 'int' })
  graceDays!: number;

  @Column({ type: 'varchar', length: 32 })
  status!: 'active' | 'retired';

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;
}
