import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'template_versions' })
@Index('IDX_template_versions_template_id', ['templateId'])
@Check(`"status" IN ('draft', 'review', 'approved', 'published', 'deprecated')`)
export class TemplateVersion {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'template_id', type: 'uuid' })
  templateId!: string;

  @Column({ type: 'int' })
  version!: number;

  @Column({ type: 'varchar', length: 32 })
  status!: 'draft' | 'review' | 'approved' | 'published' | 'deprecated';

  @Column({ name: 'author_id', type: 'uuid' })
  authorId!: string;

  @Column({ name: 'reviewer_id', type: 'uuid', nullable: true })
  reviewerId: string | null = null;

  @Column({ name: 'change_summary', type: 'text', nullable: true })
  changeSummary: string | null = null;

  @Column({ type: 'boolean' })
  immutable!: boolean;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;

  @Column({ name: 'published_at', type: 'timestamptz', nullable: true })
  publishedAt: Date | null = null;
}
