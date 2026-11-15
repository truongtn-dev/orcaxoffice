import { Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'order_items' })
@Index('IDX_order_items_order_id', ['orderId'])
export class OrderItem {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'order_id', type: 'uuid' })
  orderId!: string;

  @Column({ name: 'plan_id', type: 'uuid' })
  planId!: string;

  @Column({ type: 'int' })
  quantity!: number;

  @Column({ name: 'unit_amount_minor', type: 'int' })
  unitAmountMinor!: number;

  @Column({ name: 'line_total_minor', type: 'int' })
  lineTotalMinor!: number;
}
