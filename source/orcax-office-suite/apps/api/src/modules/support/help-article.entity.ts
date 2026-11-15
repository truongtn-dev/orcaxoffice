import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

export type HelpArticleStatus = 'draft' | 'published';

@Entity({ name: 'help_articles' })
@Index('IDX_help_articles_category', ['category'])
@Check(`"status" IN ('draft', 'published')`)
export class HelpArticle {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ type: 'varchar', length: 128, unique: true })
  slug!: string;

  @Column({ type: 'varchar', length: 255 })
  title!: string;

  @Column({ type: 'varchar', length: 64 })
  category!: string;

  @Column({ type: 'text' })
  content!: string;

  @Column({ name: 'product_version', type: 'varchar', length: 32, default: '1.0' })
  productVersion: string = '1.0';

  @Column({ type: 'varchar', length: 32, default: 'published' })
  status: HelpArticleStatus = 'published';

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'updated_at', type: 'timestamptz' })
  updatedAt!: Date;
}
