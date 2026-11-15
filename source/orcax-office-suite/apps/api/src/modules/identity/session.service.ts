import { HttpStatus, Inject, Injectable, forwardRef } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { InjectRepository } from '@nestjs/typeorm';
import { ErrorCodes } from '@orcax/api-contracts';
import { randomBytes, randomUUID } from 'crypto';
import { IsNull, Repository } from 'typeorm';
import { ApiError } from '../../common/http/api-error';
import { isUuid } from '../../common/ids/uuid';
import { AppConfig } from '../../config/validate-env';
import { MembershipService } from '../organization/membership.service';
import { AuditRecorder } from '../audit/audit-recorder.service';
import { Session } from './session.entity';
import { User } from './user.entity';
import { packSession, sha256, sha256Equal } from './session-cookie';

const INVALID_SESSION = 'Phiên đăng nhập không còn hiệu lực.';

export interface IssuedSession {
  sessionId: string;
  cookieValue: string;
  expiresAt: Date;
}

@Injectable()
export class SessionService {
  constructor(
    @InjectRepository(Session)
    private readonly sessions: Repository<Session>,
    @InjectRepository(User)
    private readonly users: Repository<User>,
    private readonly config: ConfigService,
    private readonly recorder: AuditRecorder,
    @Inject(forwardRef(() => MembershipService))
    private readonly memberships: MembershipService,
  ) {}

  async issue(input: {
    userId: string;
    workspaceId: string;
    userAgent: string | null;
  }): Promise<IssuedSession> {
    const now = new Date();
    const secret = randomBytes(32).toString('base64url');
    const session = this.sessions.create({
      id: randomUUID(),
      userId: input.userId,
      secretHash: sha256(secret),
      activeWorkspaceId: input.workspaceId,
      userAgent: input.userAgent,
      expiresAt: this.expiry(now),
      revokedAt: null,
      createdAt: now,
      rotatedAt: now,
    });
    await this.sessions.insert(session);
    return {
      sessionId: session.id,
      cookieValue: packSession(session.id, secret),
      expiresAt: session.expiresAt,
    };
  }

  async requireLive(sessionId: string, userId: string): Promise<{ session: Session; user: User }> {
    const session = await this.sessions.findOne({ where: { id: sessionId, userId } });
    if (!session || session.revokedAt || session.expiresAt <= new Date()) {
      throw new ApiError(HttpStatus.UNAUTHORIZED, ErrorCodes.UNAUTHORIZED, INVALID_SESSION);
    }
    const user = await this.users.findOne({ where: { id: userId } });
    if (!user || user.status === 'pending') {
      throw new ApiError(HttpStatus.UNAUTHORIZED, ErrorCodes.UNAUTHORIZED, INVALID_SESSION);
    }
    return { session, user };
  }

  async rotate(
    cookieValue: string,
    correlationId: string,
  ): Promise<IssuedSession & { user: User; workspaceId: string }> {
    const current = await this.matchCookie(cookieValue);
    if (!current) {
      throw new ApiError(HttpStatus.UNAUTHORIZED, ErrorCodes.UNAUTHORIZED, INVALID_SESSION);
    }
    const secret = randomBytes(32).toString('base64url');
    const now = new Date();
    const expiresAt = this.expiry(now);
    await this.sessions.update(
      { id: current.session.id },
      { secretHash: sha256(secret), rotatedAt: now, expiresAt },
    );
    await this.recorder.append({
      workspaceId: current.session.activeWorkspaceId,
      actorId: current.user.id,
      action: 'session.refreshed',
      targetType: 'session',
      targetId: current.session.id,
      outcome: 'success',
      correlationId,
    });
    return {
      sessionId: current.session.id,
      cookieValue: packSession(current.session.id, secret),
      expiresAt,
      user: current.user,
      workspaceId: current.session.activeWorkspaceId,
    };
  }

  async list(userId: string, currentSessionId: string) {
    const rows = await this.sessions.find({
      where: { userId, revokedAt: IsNull() },
      order: { createdAt: 'DESC' },
    });
    const now = new Date();
    return rows
      .filter((row) => row.expiresAt > now)
      .map((row) => ({
        id: row.id,
        createdAt: row.createdAt.toISOString(),
        expiresAt: row.expiresAt.toISOString(),
        userAgent: row.userAgent,
        current: row.id === currentSessionId,
      }));
  }

  async revoke(
    userId: string,
    sessionId: string,
    actorSessionId: string,
    correlationId: string,
  ): Promise<void> {
    if (!isUuid(sessionId)) {
      throw new ApiError(
        HttpStatus.NOT_FOUND,
        ErrorCodes.NOT_FOUND,
        'Không tìm thấy phiên đăng nhập.',
      );
    }
    const session = await this.sessions.findOne({ where: { id: sessionId, userId } });
    if (!session || session.revokedAt) {
      throw new ApiError(
        HttpStatus.NOT_FOUND,
        ErrorCodes.NOT_FOUND,
        'Không tìm thấy phiên đăng nhập.',
      );
    }
    await this.sessions.update({ id: session.id }, { revokedAt: new Date() });
    await this.recorder.append({
      workspaceId: session.activeWorkspaceId,
      actorId: userId,
      action: 'session.revoked',
      targetType: 'session',
      targetId: session.id,
      outcome: 'success',
      correlationId,
      metadata: { current: session.id === actorSessionId },
    });
  }

  async revokeAll(userId: string): Promise<void> {
    await this.sessions.update({ userId, revokedAt: IsNull() }, { revokedAt: new Date() });
  }

  async switchWorkspace(input: {
    userId: string;
    sessionId: string;
    workspaceId: string;
    confirmed: boolean;
    correlationId: string;
  }): Promise<{ workspaceId: string }> {
    if (input.confirmed !== true) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.WORKSPACE_CONFIRMATION_REQUIRED,
        'Cần xác nhận trước khi chuyển không gian làm việc.',
      );
    }
    const target = await this.memberships.requireSwitchable(input.userId, input.workspaceId);
    const now = new Date();
    await this.sessions.update(
      { id: input.sessionId, userId: input.userId },
      { activeWorkspaceId: target.workspaceId },
    );
    await this.users.update(
      { id: input.userId },
      { lastWorkspaceId: target.workspaceId, updatedAt: now },
    );
    await this.recorder.append({
      workspaceId: target.workspaceId,
      actorId: input.userId,
      action: 'workspace.switched',
      targetType: 'workspace',
      targetId: target.workspaceId,
      outcome: 'success',
      correlationId: input.correlationId,
      metadata: { kind: target.workspaceKind },
    });
    return { workspaceId: target.workspaceId };
  }

  cookieMaxAgeSeconds(): number {
    return this.config.getOrThrow<number>('sessionTtlDays') * 24 * 60 * 60;
  }

  secureCookie(): boolean {
    return this.config.getOrThrow<AppConfig['nodeEnv']>('nodeEnv') === 'production';
  }

  private expiry(from: Date): Date {
    const days = this.config.getOrThrow<number>('sessionTtlDays');
    return new Date(from.getTime() + days * 24 * 60 * 60 * 1000);
  }

  private async matchCookie(cookieValue: string): Promise<{ session: Session; user: User } | null> {
    const separator = cookieValue.indexOf('.');
    if (separator <= 0) {
      return null;
    }
    const sessionId = cookieValue.slice(0, separator);
    const secret = cookieValue.slice(separator + 1);
    if (!isUuid(sessionId) || !secret) {
      return null;
    }
    const session = await this.sessions.findOne({ where: { id: sessionId, revokedAt: IsNull() } });
    if (!session || session.expiresAt <= new Date() || !sha256Equal(secret, session.secretHash)) {
      return null;
    }
    const user = await this.users.findOne({ where: { id: session.userId } });
    if (!user || user.status !== 'active') {
      return null;
    }
    return { session, user };
  }
}
