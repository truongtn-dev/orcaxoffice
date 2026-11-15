import { Check, Column, Entity, PrimaryColumn } from 'typeorm';

@Entity({ name: 'products' })
@Check('CHK_products_status', `"status" IN ('active', 'retired')`)
export class Product {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ type: 'varchar', length: 64 })
  code!: string;

  @Column({ type: 'varchar', length: 160 })
  name!: string;

  @Column({ type: 'varchar', length: 32 })
  status!: 'active' | 'retired';

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;
}
