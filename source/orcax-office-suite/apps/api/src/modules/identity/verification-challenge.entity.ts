import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

export type ChallengePurpose = 'email_verification' | 'password_reset';

@Entity({ name: 'verification_challenges' })
@Index('IDX_verification_challenges_user_purpose', ['userId', 'purpose'])
@Check(
  'CHK_verification_challenges_purpose',
  `"purpose" IN ('email_verification', 'password_reset')`,
)
export class VerificationChallenge {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'user_id', type: 'uuid' })
  userId!: string;

  @Column({ type: 'varchar', length: 32 })
  purpose!: ChallengePurpose;

  @Column({ name: 'code_hash', type: 'varchar', length: 255 })
  codeHash!: string;

  @Column({ type: 'integer', default: 0 })
  attempts!: number;

  @Column({ name: 'expires_at', type: 'timestamptz' })
  expiresAt!: Date;

  @Column({ name: 'consumed_at', type: 'timestamptz', nullable: true })
  consumedAt: Date | null = null;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;
}
