import { Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'legal_references' })
@Index('IDX_legal_references_version_id', ['profileVersionId'])
export class LegalReference {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'profile_version_id', type: 'uuid' })
  profileVersionId!: string;

  @Column({ type: 'varchar', length: 500 })
  citation!: string;

  @Column({ name: 'source_url', type: 'varchar', length: 1000, nullable: true })
  sourceUrl: string | null = null;

  @Column({ type: 'boolean' })
  verified!: boolean;

  @Column({ type: 'text', nullable: true })
  notes: string | null = null;
}
