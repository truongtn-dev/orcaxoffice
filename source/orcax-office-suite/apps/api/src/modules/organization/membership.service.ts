import { HttpStatus, Inject, Injectable, forwardRef } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { ErrorCodes } from '@orcax/api-contracts';
import { createHash, randomBytes, randomUUID, timingSafeEqual } from 'crypto';
import { Repository } from 'typeorm';
import { ApiError } from '../../common/http/api-error';
import { AccessContext } from '../../tenancy/access-context';
import { AuditRecorder } from '../audit/audit-recorder.service';
import { AuthorizationService } from '../authorization/authorization.service';
import { RoleCatalog } from '../authorization/role-catalog';
import { EmailSender } from '../identity/email-sender';
import { UserDirectory } from '../identity/user-directory';
import { Invitation } from './invitation.entity';
import { Membership } from './membership.entity';
import { Workspace } from './workspace.entity';
import { WorkspaceAccess } from './workspace-access';

const INVALID_INVITATION = 'Lời mời không hợp lệ hoặc đã hết hạn.';
const INVITE_DAYS = 7;

@Injectable()
export class MembershipService {
  constructor(
    @InjectRepository(Membership)
    private readonly memberships: Repository<Membership>,
    @InjectRepository(Workspace)
    private readonly workspaces: Repository<Workspace>,
    @InjectRepository(Invitation)
    private readonly invitations: Repository<Invitation>,
    private readonly roles: RoleCatalog,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
    private readonly email: EmailSender,
    @Inject(forwardRef(() => UserDirectory))
    private readonly users: UserDirectory,
  ) {}

  async describe(userId: string, workspaceId: string): Promise<WorkspaceAccess | null> {
    const membership = await this.memberships.findOne({ where: { userId, workspaceId } });
    if (!membership) {
      return null;
    }
    const workspace = await this.workspaces.findOne({ where: { id: workspaceId } });
    if (!workspace) {
      return null;
    }
    const role = await this.roles.requireById(membership.roleId);
    return this.toAccess(membership, workspace, role.code);
  }

  async requireSwitchable(userId: string, workspaceId: string): Promise<WorkspaceAccess> {
    const access = await this.describe(userId, workspaceId);
    if (!access || access.membershipStatus !== 'active' || access.workspaceStatus !== 'active') {
      throw new ApiError(
        HttpStatus.NOT_FOUND,
        ErrorCodes.NOT_FOUND,
        'Không gian làm việc không khả dụng.',
      );
    }
    return access;
  }

  async listForUser(userId: string) {
    const memberships = await this.memberships.find({ where: { userId } });
    const visible = memberships.filter((item) => item.status !== 'removed');
    const result = [];
    for (const membership of visible) {
      const workspace = await this.workspaces.findOne({ where: { id: membership.workspaceId } });
      if (!workspace || workspace.status === 'archived') {
        continue;
      }
      const role = await this.roles.requireById(membership.roleId);
      result.push({
        id: workspace.id,
        name: workspace.name,
        kind: workspace.kind,
        status: workspace.status,
        membershipStatus: membership.status,
        role: role.code,
        switchable: membership.status === 'active' && workspace.status === 'active',
      });
    }
    return result;
  }

  async listMembers(actor: AccessContext) {
    this.authorization.assert(actor.roleCode, 'workspace.read');
    const memberships = await this.memberships.find({
      where: { workspaceId: actor.workspaceId },
      order: { createdAt: 'ASC' },
    });
    const visible = memberships.filter((item) => item.status !== 'removed');
    const summaries = await this.users.summaries(visible.map((item) => item.userId));
    const members = [];
    for (const membership of visible) {
      const role = await this.roles.requireById(membership.roleId);
      const person = summaries.get(membership.userId);
      members.push({
        id: membership.id,
        userId: membership.userId,
        email: person?.email ?? '',
        displayName: person?.displayName ?? '',
        role: role.code,
        status: membership.status,
      });
    }
    return members;
  }

  async invite(
    actor: AccessContext,
    email: string,
    roleCode: string,
  ): Promise<{ id: string; email: string; role: string; status: 'invited' }> {
    this.authorization.assertPolicy(actor, {
      permission: 'workspace.member.invite',
      workspaceKinds: ['organization'],
    });
    this.authorization.assertAssignable(actor.roleCode, roleCode);
    const normalized = email.trim().toLowerCase();
    const role = await this.roles.requireByCode(roleCode);
    const token = randomBytes(32).toString('base64url');
    const now = new Date();
    const invitationId = randomUUID();
    await this.invitations.insert({
      id: invitationId,
      workspaceId: actor.workspaceId,
      email: normalized,
      roleId: role.id,
      tokenHash: sha256(token),
      status: 'invited',
      invitedBy: actor.userId,
      expiresAt: new Date(now.getTime() + INVITE_DAYS * 24 * 60 * 60 * 1000),
      createdAt: now,
    });
    await this.email.send({
      to: normalized,
      subject: 'Lời mời tham gia không gian làm việc OrcaX Office',
      text: `Bạn được mời vào không gian ${actor.workspaceName} với vai trò ${roleCode}. Mã lời mời: ${token}. Mã hết hạn sau ${INVITE_DAYS} ngày.`,
    });
    await this.recorder.append({
      workspaceId: actor.workspaceId,
      actorId: actor.userId,
      action: 'membership.invited',
      targetType: 'invitation',
      targetId: invitationId,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { role: roleCode },
    });
    return { id: invitationId, email: normalized, role: roleCode, status: 'invited' };
  }

  async accept(
    actor: AccessContext,
    token: string,
  ): Promise<{ workspaceId: string; role: string }> {
    const match = await this.findInvitation(token);
    if (!match || match.email !== actor.email) {
      throw new ApiError(HttpStatus.BAD_REQUEST, ErrorCodes.BAD_REQUEST, INVALID_INVITATION);
    }
    const workspace = await this.workspaces.findOne({ where: { id: match.workspaceId } });
    const role = await this.roles.requireById(match.roleId);
    if (!workspace || workspace.kind !== 'organization' || workspace.status !== 'active') {
      throw new ApiError(HttpStatus.BAD_REQUEST, ErrorCodes.BAD_REQUEST, INVALID_INVITATION);
    }
    const now = new Date();
    const existing = await this.memberships.findOne({
      where: { workspaceId: workspace.id, userId: actor.userId },
    });
    if (existing && existing.status === 'active') {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Bạn đã tham gia không gian này.',
      );
    }
    if (existing) {
      await this.memberships.update(
        { id: existing.id },
        { status: 'active', roleId: role.id, updatedAt: now },
      );
    } else {
      await this.memberships.insert({
        id: randomUUID(),
        workspaceId: workspace.id,
        userId: actor.userId,
        roleId: role.id,
        status: 'active',
        createdAt: now,
        updatedAt: now,
      });
    }
    await this.invitations.update({ id: match.id }, { status: 'accepted' });
    await this.recorder.append({
      workspaceId: workspace.id,
      actorId: actor.userId,
      action: 'membership.accepted',
      targetType: 'invitation',
      targetId: match.id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { role: role.code },
    });
    return { workspaceId: workspace.id, role: role.code };
  }

  async suspend(actor: AccessContext, membershipId: string): Promise<void> {
    await this.changeStatus(
      actor,
      membershipId,
      'suspended',
      'workspace.member.suspend',
      'membership.suspended',
    );
  }

  async remove(actor: AccessContext, membershipId: string): Promise<void> {
    await this.changeStatus(
      actor,
      membershipId,
      'removed',
      'workspace.member.remove',
      'membership.removed',
    );
  }

  async changeRole(actor: AccessContext, membershipId: string, roleCode: string): Promise<void> {
    this.authorization.assertPolicy(actor, {
      permission: 'workspace.member.role',
      workspaceKinds: ['organization'],
    });
    this.authorization.assertAssignable(actor.roleCode, roleCode);
    const membership = await this.requireInWorkspace(actor.workspaceId, membershipId);
    const current = await this.roles.requireById(membership.roleId);
    this.assertCanTouch(actor.roleCode, current.code);
    const next = await this.roles.requireByCode(roleCode);
    await this.memberships.manager.transaction(async (manager) => {
      await this.assertKeepsOwner(manager, membership, roleCode);
      await manager.update(
        Membership,
        { id: membership.id },
        { roleId: next.id, updatedAt: new Date() },
      );
      await this.recorder.append(
        {
          workspaceId: actor.workspaceId,
          actorId: actor.userId,
          action: 'membership.role.changed',
          targetType: 'membership',
          targetId: membership.id,
          outcome: 'success',
          correlationId: actor.correlationId,
          metadata: { from: current.code, to: roleCode },
        },
        manager,
      );
    });
  }

  private async changeStatus(
    actor: AccessContext,
    membershipId: string,
    status: 'suspended' | 'removed',
    permission: string,
    action: string,
  ): Promise<void> {
    this.authorization.assertPolicy(actor, {
      permission: permission as 'workspace.member.suspend' | 'workspace.member.remove',
      workspaceKinds: ['organization'],
    });
    const membership = await this.requireInWorkspace(actor.workspaceId, membershipId);
    const current = await this.roles.requireById(membership.roleId);
    this.assertCanTouch(actor.roleCode, current.code);
    await this.memberships.manager.transaction(async (manager) => {
      await this.assertKeepsOwner(
        manager,
        membership,
        status === 'suspended' ? 'suspended' : 'removed',
      );
      await manager.update(Membership, { id: membership.id }, { status, updatedAt: new Date() });
      await this.recorder.append(
        {
          workspaceId: actor.workspaceId,
          actorId: actor.userId,
          action,
          targetType: 'membership',
          targetId: membership.id,
          outcome: 'success',
          correlationId: actor.correlationId,
          metadata: { role: current.code },
        },
        manager,
      );
    });
  }

  private assertCanTouch(actorRole: string, targetRole: string): void {
    if (targetRole === 'personal_owner') {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Bạn không thể gán vai trò này.',
      );
    }
    if (targetRole === 'organization_owner' && actorRole !== 'organization_owner') {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Chỉ chủ sở hữu mới thay đổi chủ sở hữu khác.',
      );
    }
  }

  private async requireInWorkspace(workspaceId: string, membershipId: string): Promise<Membership> {
    const membership = await this.memberships.findOne({ where: { id: membershipId, workspaceId } });
    if (!membership || membership.status === 'removed') {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy thành viên.');
    }
    return membership;
  }

  private async assertKeepsOwner(
    manager: Repository<Membership>['manager'],
    membership: Membership,
    next: string,
  ): Promise<void> {
    const owner = await this.roles.requireByCode('organization_owner', manager);
    if (membership.roleId !== owner.id || membership.status !== 'active') {
      return;
    }
    if (next === 'organization_owner') {
      return;
    }
    const count = await manager.getRepository(Membership).count({
      where: { workspaceId: membership.workspaceId, roleId: owner.id, status: 'active' },
    });
    if (count <= 1) {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Không gian làm việc cần ít nhất một chủ sở hữu đang hoạt động.',
      );
    }
  }

  private async findInvitation(token: string): Promise<Invitation | null> {
    if (!token || token.length < 20) {
      return null;
    }
    const hash = sha256(token);
    const invitation = await this.invitations.findOne({
      where: { tokenHash: hash, status: 'invited' },
    });
    if (!invitation || invitation.expiresAt <= new Date()) {
      if (invitation) {
        await this.invitations.update({ id: invitation.id }, { status: 'expired' });
      }
      return null;
    }
    const actual = Buffer.from(hash, 'hex');
    const expected = Buffer.from(invitation.tokenHash, 'hex');
    if (actual.length !== expected.length || !timingSafeEqual(actual, expected)) {
      return null;
    }
    return invitation;
  }

  private toAccess(
    membership: Membership,
    workspace: Workspace,
    roleCode: string,
  ): WorkspaceAccess {
    return {
      membershipId: membership.id,
      workspaceId: workspace.id,
      workspaceKind: workspace.kind,
      workspaceName: workspace.name,
      workspaceStatus: workspace.status,
      membershipStatus: membership.status,
      roleCode,
      organizationId: workspace.organizationId,
    };
  }
}

function sha256(value: string): string {
  return createHash('sha256').update(value).digest('hex');
}
