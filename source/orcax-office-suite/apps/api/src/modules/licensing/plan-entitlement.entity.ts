import { Entity, PrimaryColumn } from 'typeorm';

@Entity({ name: 'plan_entitlements' })
export class PlanEntitlement {
  @PrimaryColumn({ name: 'plan_id', type: 'uuid' })
  planId!: string;

  @PrimaryColumn({ name: 'entitlement_code', type: 'varchar', length: 64 })
  entitlementCode!: string;
}
