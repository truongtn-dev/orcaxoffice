import { HttpStatus, Inject, Injectable, forwardRef } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { ErrorCodes } from '@orcax/api-contracts';
import { randomUUID } from 'crypto';
import { DataSource, IsNull, Repository } from 'typeorm';
import { ApiError } from '../../common/http/api-error';
import { isUuid } from '../../common/ids/uuid';
import { loginWorkspace } from '../../tenancy/workspace-scope';
import { AuditRecorder } from '../audit/audit-recorder.service';
import { decideLogin, INVALID_CREDENTIALS_MESSAGE, permissionsFor } from '../authorization/catalog';
import { MembershipService } from '../organization/membership.service';
import { WorkspaceAccess } from '../organization/workspace-access';
import { WorkspaceService } from '../organization/workspace.service';
import { ChallengeService } from './challenge.service';
import { dummyPasswordHash, hashSecret, verifySecret } from './password-hasher';
import { Session } from './session.entity';
import { SessionService } from './session.service';
import { TokenService } from './token.service';
import { User } from './user.entity';

const GENERIC_ACCOUNT_MESSAGE = 'Nếu tài khoản đủ điều kiện, hướng dẫn đã được gửi tới email.';
const INVALID_CODE = 'Mã không đúng hoặc đã hết hạn.';

export interface SessionView {
  user: { id: string; email: string; displayName: string; status: string };
  workspace: {
    id: string;
    name: string;
    kind: string;
    role: string;
    status: string;
  };
  suggestedWorkspace: { id: string; name: string; kind: string } | null;
  permissions: string[];
}

@Injectable()
export class AuthService {
  constructor(
    @InjectRepository(User)
    private readonly users: Repository<User>,
    private readonly dataSource: DataSource,
    private readonly challenges: ChallengeService,
    private readonly sessions: SessionService,
    private readonly tokens: TokenService,
    private readonly recorder: AuditRecorder,
    @Inject(forwardRef(() => WorkspaceService))
    private readonly workspaces: WorkspaceService,
    @Inject(forwardRef(() => MembershipService))
    private readonly memberships: MembershipService,
  ) {}

  async register(input: {
    email: string;
    password: string;
    displayName: string;
    correlationId: string;
  }): Promise<{ message: string }> {
    this.assertCorrelation(input.correlationId);
    const email = normalizeEmail(input.email);
    const displayName = input.displayName.trim();
    const passwordHash = await hashSecret(input.password);
    const existing = await this.users.findOne({ where: { email } });
    if (existing) {
      if (existing.status === 'pending') {
        const personal = await this.workspaces.findPersonal(existing.id);
        await this.challenges.issue({
          user: existing,
          purpose: 'email_verification',
          workspaceId: personal.id,
          correlationId: input.correlationId,
        });
      }
      return { message: GENERIC_ACCOUNT_MESSAGE };
    }

    const userId = randomUUID();
    const now = new Date();
    let workspaceId = '';
    await this.dataSource.transaction(async (manager) => {
      await manager.insert(User, {
        id: userId,
        email,
        passwordHash,
        displayName,
        status: 'pending',
        emailVerifiedAt: null,
        lastWorkspaceId: null,
        createdAt: now,
        updatedAt: now,
      });
      const created = await this.workspaces.createPersonal(manager, { userId, displayName });
      workspaceId = created.workspaceId;
      await manager.update(User, { id: userId }, { lastWorkspaceId: workspaceId });
      await this.recorder.append(
        {
          workspaceId,
          actorId: userId,
          action: 'identity.registered',
          targetType: 'user',
          targetId: userId,
          outcome: 'success',
          correlationId: input.correlationId,
        },
        manager,
      );
    });
    const user = await this.users.findOneByOrFail({ id: userId });
    await this.challenges.issue({
      user,
      purpose: 'email_verification',
      workspaceId,
      correlationId: input.correlationId,
    });
    return { message: GENERIC_ACCOUNT_MESSAGE };
  }

  async login(input: {
    email: string;
    password: string;
    userAgent: string | null;
    correlationId: string;
  }): Promise<SessionView & { accessToken: string; expiresIn: number; cookieValue: string }> {
    this.assertCorrelation(input.correlationId);
    const email = normalizeEmail(input.email);
    const user = await this.users.findOne({ where: { email } });
    const hash = user?.passwordHash ?? (await dummyPasswordHash());
    const passwordMatches = await verifySecret(input.password, hash);
    const decision = decideLogin({
      userFound: Boolean(user),
      passwordMatches,
      status: user?.status,
    });
    if (!user || decision === 'invalid') {
      await this.recorder.append({
        actorId: user?.id ?? null,
        action: 'identity.login.failed',
        targetType: 'user',
        targetId: user?.id ?? null,
        outcome: 'failure',
        correlationId: input.correlationId,
        metadata: user ? { reason: 'credentials' } : null,
      });
      throw new ApiError(
        HttpStatus.UNAUTHORIZED,
        ErrorCodes.UNAUTHORIZED,
        INVALID_CREDENTIALS_MESSAGE,
      );
    }
    if (decision === 'unverified') {
      await this.recordLoginFailure(user, input.correlationId, 'unverified');
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.EMAIL_NOT_VERIFIED,
        'Email chưa được xác minh.',
      );
    }
    if (decision === 'suspended') {
      await this.recordLoginFailure(user, input.correlationId, 'suspended');
      throw new ApiError(HttpStatus.FORBIDDEN, ErrorCodes.FORBIDDEN, 'Tài khoản đã bị khóa.');
    }

    const personal = await this.workspaces.findPersonal(user.id);
    const choice = loginWorkspace({
      personalWorkspaceId: personal.id,
      lastWorkspaceId: user.lastWorkspaceId,
    });
    const issued = await this.sessions.issue({
      userId: user.id,
      workspaceId: choice.activeWorkspaceId,
      userAgent: input.userAgent,
    });
    await this.recorder.append({
      workspaceId: choice.activeWorkspaceId,
      actorId: user.id,
      action: 'identity.login.succeeded',
      targetType: 'session',
      targetId: issued.sessionId,
      outcome: 'success',
      correlationId: input.correlationId,
    });
    const view = await this.present(user, choice.activeWorkspaceId);
    const access = this.tokens.sign(user.id, issued.sessionId, choice.activeWorkspaceId);
    return {
      ...view,
      accessToken: access.token,
      expiresIn: access.expiresIn,
      cookieValue: issued.cookieValue,
    };
  }

  async logout(userId: string, sessionId: string, correlationId: string): Promise<void> {
    await this.sessions.revoke(userId, sessionId, sessionId, correlationId);
    await this.recorder.append({
      actorId: userId,
      action: 'identity.logout',
      targetType: 'session',
      targetId: sessionId,
      outcome: 'success',
      correlationId,
    });
  }

  async verifyEmail(
    email: string,
    code: string,
    correlationId: string,
  ): Promise<{ verified: true }> {
    this.assertCorrelation(correlationId);
    const user = await this.users.findOne({ where: { email: normalizeEmail(email) } });
    const consumed = user
      ? await this.challenges.consume(user.id, 'email_verification', code)
      : await this.rejectCode(code);
    if (!user || !consumed) {
      throw new ApiError(HttpStatus.BAD_REQUEST, ErrorCodes.BAD_REQUEST, INVALID_CODE);
    }
    if (user.status === 'pending') {
      const now = new Date();
      await this.users.update(
        { id: user.id },
        { status: 'active', emailVerifiedAt: now, updatedAt: now },
      );
    }
    await this.recorder.append({
      workspaceId: user.lastWorkspaceId,
      actorId: user.id,
      action: 'identity.email.verified',
      targetType: 'user',
      targetId: user.id,
      outcome: 'success',
      correlationId,
    });
    return { verified: true };
  }

  async resend(email: string, correlationId: string): Promise<{ message: string }> {
    this.assertCorrelation(correlationId);
    const user = await this.users.findOne({ where: { email: normalizeEmail(email) } });
    if (user?.status === 'pending') {
      const personal = await this.workspaces.findPersonal(user.id);
      await this.challenges.issue({
        user,
        purpose: 'email_verification',
        workspaceId: personal.id,
        correlationId,
      });
    }
    return { message: GENERIC_ACCOUNT_MESSAGE };
  }

  async forgot(email: string, correlationId: string): Promise<{ message: string }> {
    this.assertCorrelation(correlationId);
    const user = await this.users.findOne({ where: { email: normalizeEmail(email) } });
    if (user?.status === 'active') {
      await this.challenges.issue({
        user,
        purpose: 'password_reset',
        workspaceId: user.lastWorkspaceId,
        correlationId,
      });
    }
    return { message: GENERIC_ACCOUNT_MESSAGE };
  }

  async resetPassword(input: {
    email: string;
    code: string;
    password: string;
    correlationId: string;
  }): Promise<{ reset: true }> {
    this.assertCorrelation(input.correlationId);
    const user = await this.users.findOne({ where: { email: normalizeEmail(input.email) } });
    const consumed =
      user && user.status === 'active'
        ? await this.challenges.consume(user.id, 'password_reset', input.code)
        : await this.rejectCode(input.code);
    if (!user || user.status !== 'active' || !consumed) {
      throw new ApiError(HttpStatus.BAD_REQUEST, ErrorCodes.BAD_REQUEST, INVALID_CODE);
    }
    const passwordHash = await hashSecret(input.password);
    const now = new Date();
    await this.dataSource.transaction(async (manager) => {
      await manager.update(User, { id: user.id }, { passwordHash, updatedAt: now });
      await manager.update(Session, { userId: user.id, revokedAt: IsNull() }, { revokedAt: now });
      await this.recorder.append(
        {
          workspaceId: user.lastWorkspaceId,
          actorId: user.id,
          action: 'identity.password.reset.completed',
          targetType: 'user',
          targetId: user.id,
          outcome: 'success',
          correlationId: input.correlationId,
        },
        manager,
      );
    });
    return { reset: true };
  }

  async updateProfile(
    userId: string,
    displayName: string,
    activeWorkspaceId: string,
    correlationId: string,
  ): Promise<SessionView> {
    const trimmed = displayName.trim();
    if (trimmed.length < 1 || trimmed.length > 120) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.VALIDATION_FAILED,
        'Tên hiển thị cần từ 1 đến 120 ký tự.',
      );
    }
    const now = new Date();
    await this.users.update({ id: userId }, { displayName: trimmed, updatedAt: now });
    await this.recorder.append({
      workspaceId: activeWorkspaceId,
      actorId: userId,
      action: 'identity.profile.updated',
      targetType: 'user',
      targetId: userId,
      outcome: 'success',
      correlationId,
    });
    const updated = await this.users.findOneByOrFail({ id: userId });
    return this.present(updated, activeWorkspaceId);
  }

  async presentAccess(userId: string, activeWorkspaceId: string): Promise<SessionView> {
    const user = await this.users.findOneByOrFail({ id: userId });
    return this.present(user, activeWorkspaceId);
  }

  async present(user: User, activeWorkspaceId: string): Promise<SessionView> {
    const active = await this.memberships.describe(user.id, activeWorkspaceId);
    if (!active) {
      throw new ApiError(
        HttpStatus.UNAUTHORIZED,
        ErrorCodes.UNAUTHORIZED,
        'Phiên đăng nhập không còn hiệu lực.',
      );
    }
    return {
      user: {
        id: user.id,
        email: user.email,
        displayName: user.displayName,
        status: user.status,
      },
      workspace: workspaceSummary(active),
      suggestedWorkspace: await this.suggestion(user, active.workspaceId),
      permissions: [...permissionsFor(active.roleCode)],
    };
  }

  private async suggestion(user: User, activeWorkspaceId: string) {
    if (!user.lastWorkspaceId || user.lastWorkspaceId === activeWorkspaceId) {
      return null;
    }
    const candidate = await this.memberships.describe(user.id, user.lastWorkspaceId);
    if (
      !candidate ||
      candidate.membershipStatus !== 'active' ||
      candidate.workspaceStatus !== 'active'
    ) {
      return null;
    }
    return {
      id: candidate.workspaceId,
      name: candidate.workspaceName,
      kind: candidate.workspaceKind,
    };
  }

  private async recordLoginFailure(
    user: User,
    correlationId: string,
    reason: string,
  ): Promise<void> {
    await this.recorder.append({
      actorId: user.id,
      action: 'identity.login.failed',
      targetType: 'user',
      targetId: user.id,
      outcome: 'failure',
      correlationId,
      metadata: { reason },
    });
  }

  private async rejectCode(code: string): Promise<false> {
    await verifySecret(code, await dummyPasswordHash());
    return false;
  }

  private assertCorrelation(correlationId: string): void {
    if (!isUuid(correlationId)) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Yêu cầu thiếu mã tương quan.',
      );
    }
  }
}

function normalizeEmail(email: string): string {
  return email.trim().toLowerCase();
}

function workspaceSummary(access: WorkspaceAccess): SessionView['workspace'] {
  return {
    id: access.workspaceId,
    name: access.workspaceName,
    kind: access.workspaceKind,
    role: access.roleCode,
    status: access.workspaceStatus,
  };
}
