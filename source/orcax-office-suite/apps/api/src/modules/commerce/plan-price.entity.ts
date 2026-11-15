import { Check, Column, Entity, PrimaryColumn } from 'typeorm';

@Entity({ name: 'plan_prices' })
@Check(`"billing_period" IN ('once', 'year', 'month')`)
@Check(`"status" IN ('active', 'retired')`)
export class PlanPrice {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'plan_id', type: 'uuid' })
  planId!: string;

  @Column({ type: 'varchar', length: 8 })
  currency!: string;

  @Column({ name: 'amount_minor', type: 'int' })
  amountMinor!: number;

  @Column({ name: 'billing_period', type: 'varchar', length: 16 })
  billingPeriod!: 'once' | 'year' | 'month';

  @Column({ type: 'varchar', length: 32 })
  status!: 'active' | 'retired';
}
