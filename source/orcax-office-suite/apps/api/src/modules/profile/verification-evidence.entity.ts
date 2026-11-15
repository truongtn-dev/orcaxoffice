import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'verification_evidence' })
@Index('IDX_verification_evidence_legal_id', ['legalReferenceId'])
@Check(`"evidence_type" IN ('document', 'url', 'statement')`)
export class VerificationEvidence {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'legal_reference_id', type: 'uuid' })
  legalReferenceId!: string;

  @Column({ name: 'recorded_by', type: 'uuid' })
  recordedBy!: string;

  @Column({ name: 'recorded_at', type: 'timestamptz' })
  recordedAt!: Date;

  @Column({ name: 'evidence_type', type: 'varchar', length: 32 })
  evidenceType!: 'document' | 'url' | 'statement';

  @Column({ type: 'text' })
  summary!: string;

  @Column({ name: 'reference_uri', type: 'varchar', length: 1000, nullable: true })
  referenceUri: string | null = null;
}
