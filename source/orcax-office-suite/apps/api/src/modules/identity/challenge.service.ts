import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { randomInt, randomUUID } from 'crypto';
import { IsNull, Repository } from 'typeorm';
import { AuditRecorder } from '../audit/audit-recorder.service';
import { ChallengePurpose, VerificationChallenge } from './verification-challenge.entity';
import { EmailSender } from './email-sender';
import { dummyPasswordHash, hashSecret, verifySecret } from './password-hasher';
import { User } from './user.entity';

const CODE_TTL_MS = 15 * 60 * 1000;
const MAX_ATTEMPTS = 5;

@Injectable()
export class ChallengeService {
  constructor(
    @InjectRepository(VerificationChallenge)
    private readonly challenges: Repository<VerificationChallenge>,
    private readonly email: EmailSender,
    private readonly recorder: AuditRecorder,
  ) {}

  async issue(input: {
    user: User;
    purpose: ChallengePurpose;
    workspaceId: string | null;
    correlationId: string;
  }): Promise<void> {
    const code = randomInt(0, 1_000_000).toString().padStart(6, '0');
    const now = new Date();
    await this.challenges.update(
      { userId: input.user.id, purpose: input.purpose, consumedAt: IsNull() },
      { consumedAt: now },
    );
    await this.challenges.insert({
      id: randomUUID(),
      userId: input.user.id,
      purpose: input.purpose,
      codeHash: await hashSecret(code),
      attempts: 0,
      expiresAt: new Date(now.getTime() + CODE_TTL_MS),
      consumedAt: null,
      createdAt: now,
    });
    const message = this.message(input.purpose, code);
    await this.email.send({ to: input.user.email, subject: message.subject, text: message.text });
    await this.recorder.append({
      workspaceId: input.workspaceId,
      actorId: input.user.id,
      action:
        input.purpose === 'email_verification'
          ? 'identity.email.verification.sent'
          : 'identity.password.reset.requested',
      targetType: 'user',
      targetId: input.user.id,
      outcome: 'success',
      correlationId: input.correlationId,
      metadata: { purpose: input.purpose },
    });
  }

  async consume(userId: string, purpose: ChallengePurpose, code: string): Promise<boolean> {
    const challenge = await this.challenges.findOne({
      where: { userId, purpose, consumedAt: IsNull() },
      order: { createdAt: 'DESC' },
    });
    if (!challenge || challenge.attempts >= MAX_ATTEMPTS || challenge.expiresAt <= new Date()) {
      await verifySecret(code, await dummyPasswordHash());
      return false;
    }
    const matches = await verifySecret(code, challenge.codeHash);
    if (!matches) {
      await this.challenges.increment({ id: challenge.id }, 'attempts', 1);
      return false;
    }
    await this.challenges.update({ id: challenge.id }, { consumedAt: new Date() });
    return true;
  }

  private message(purpose: ChallengePurpose, code: string): { subject: string; text: string } {
    if (purpose === 'email_verification') {
      return {
        subject: 'Mã xác minh email OrcaX Office',
        text: `Mã xác minh của bạn là ${code}. Mã hết hạn sau 15 phút.`,
      };
    }
    return {
      subject: 'Mã đặt lại mật khẩu OrcaX Office',
      text: `Mã đặt lại mật khẩu của bạn là ${code}. Mã hết hạn sau 15 phút.`,
    };
  }
}
