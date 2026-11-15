import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'profile_versions' })
@Index('IDX_profile_versions_profile_id', ['profileId'])
@Check(`"status" IN ('draft', 'in_review', 'approved', 'deprecated')`)
export class ProfileVersion {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'profile_id', type: 'uuid' })
  profileId!: string;

  @Column({ type: 'int' })
  version!: number;

  @Column({ type: 'varchar', length: 32 })
  status!: 'draft' | 'in_review' | 'approved' | 'deprecated';

  @Column({ type: 'text' })
  source!: string;

  @Column({ name: 'effective_at', type: 'timestamptz', nullable: true })
  effectiveAt: Date | null = null;

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
