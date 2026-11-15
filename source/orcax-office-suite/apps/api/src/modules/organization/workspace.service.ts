import { HttpStatus, Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { ErrorCodes } from '@orcax/api-contracts';
import { randomUUID } from 'crypto';
import { EntityManager, Repository } from 'typeorm';
import { ApiError } from '../../common/http/api-error';
import { AccessContext } from '../../tenancy/access-context';
import { assertCanCreateOrganization, WorkspaceBoundaryError } from '../../tenancy/workspace-scope';
import { AuditRecorder } from '../audit/audit-recorder.service';
import { AuthorizationService } from '../authorization/authorization.service';
import { RoleCatalog } from '../authorization/role-catalog';
import { Organization } from './organization.entity';
import { Workspace } from './workspace.entity';
import { Membership } from './membership.entity';

@Injectable()
export class WorkspaceService {
  constructor(
    @InjectRepository(Organization)
    private readonly organizations: Repository<Organization>,
    @InjectRepository(Workspace)
    private readonly workspaces: Repository<Workspace>,
    private readonly roles: RoleCatalog,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
  ) {}

  async createPersonal(
    manager: EntityManager,
    input: { userId: string; displayName: string },
  ): Promise<{ workspaceId: string }> {
    const role = await this.roles.requireByCode('personal_owner', manager);
    const now = new Date();
    const workspaceId = randomUUID();
    await manager.insert(Workspace, {
      id: workspaceId,
      kind: 'personal',
      name: 'Cá nhân',
      status: 'active',
      ownerUserId: input.userId,
      organizationId: null,
      createdAt: now,
      updatedAt: now,
    });
    await manager.insert(Membership, {
      id: randomUUID(),
      workspaceId,
      userId: input.userId,
      roleId: role.id,
      status: 'active',
      createdAt: now,
      updatedAt: now,
    });
    return { workspaceId };
  }

  async findPersonal(userId: string): Promise<Workspace> {
    const workspace = await this.workspaces.findOne({
      where: { kind: 'personal', ownerUserId: userId },
    });
    if (!workspace) {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Không gian cá nhân chưa được tạo.',
      );
    }
    return workspace;
  }

  async createOrganization(
    actor: AccessContext,
    name: string,
  ): Promise<{ id: string; name: string; kind: 'organization' }> {
    try {
      assertCanCreateOrganization(actor);
    } catch (error) {
      if (error instanceof WorkspaceBoundaryError) {
        throw new ApiError(HttpStatus.FORBIDDEN, ErrorCodes.FORBIDDEN, error.message);
      }
      throw error;
    }
    const trimmed = name.trim();
    if (trimmed.length < 2 || trimmed.length > 160) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.VALIDATION_FAILED,
        'Tên tổ chức cần từ 2 đến 160 ký tự.',
      );
    }
    const organizationId = randomUUID();
    const workspaceId = randomUUID();
    const now = new Date();
    await this.organizations.manager.transaction(async (manager) => {
      const role = await this.roles.requireByCode('organization_owner', manager);
      await manager.insert(Organization, {
        id: organizationId,
        name: trimmed,
        status: 'active',
        createdBy: actor.userId,
        createdAt: now,
        updatedAt: now,
      });
      await manager.insert(Workspace, {
        id: workspaceId,
        kind: 'organization',
        name: trimmed,
        status: 'active',
        ownerUserId: null,
        organizationId,
        createdAt: now,
        updatedAt: now,
      });
      await manager.insert(Membership, {
        id: randomUUID(),
        workspaceId,
        userId: actor.userId,
        roleId: role.id,
        status: 'active',
        createdAt: now,
        updatedAt: now,
      });
      await this.recorder.append(
        {
          workspaceId,
          actorId: actor.userId,
          action: 'workspace.created',
          targetType: 'workspace',
          targetId: workspaceId,
          outcome: 'success',
          correlationId: actor.correlationId,
          metadata: { kind: 'organization' },
        },
        manager,
      );
    });
    return { id: workspaceId, name: trimmed, kind: 'organization' };
  }

  async renameCurrent(actor: AccessContext, name: string): Promise<{ id: string; name: string }> {
    this.authorization.assert(actor.roleCode, 'workspace.update');
    if (actor.workspaceKind !== 'organization' || !actor.organizationId) {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Không gian cá nhân không đổi tên tại đây.',
      );
    }
    const trimmed = name.trim();
    if (trimmed.length < 2 || trimmed.length > 160) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.VALIDATION_FAILED,
        'Tên tổ chức cần từ 2 đến 160 ký tự.',
      );
    }
    const now = new Date();
    await this.organizations.update(
      { id: actor.organizationId },
      { name: trimmed, updatedAt: now },
    );
    await this.workspaces.update({ id: actor.workspaceId }, { name: trimmed, updatedAt: now });
    await this.recorder.append({
      workspaceId: actor.workspaceId,
      actorId: actor.userId,
      action: 'workspace.updated',
      targetType: 'organization',
      targetId: actor.organizationId,
      outcome: 'success',
      correlationId: actor.correlationId,
    });
    return { id: actor.workspaceId, name: trimmed };
  }

  async currentOrganization(
    actor: AccessContext,
  ): Promise<{ id: string; name: string; status: string } | null> {
    if (actor.workspaceKind !== 'organization' || !actor.organizationId) {
      return null;
    }
    const organization = await this.organizations.findOne({ where: { id: actor.organizationId } });
    if (!organization) {
      return null;
    }
    return { id: organization.id, name: organization.name, status: organization.status };
  }

  async setOrganizationStatus(
    actor: AccessContext,
    status: 'active' | 'suspended' | 'archived',
  ): Promise<{ id: string; name: string; status: string }> {
    this.authorization.assertPolicy(actor, {
      permission: 'workspace.update',
      workspaceKinds: ['organization'],
    });
    if (!actor.organizationId || !actor.workspaceId) {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Chỉ áp dụng cho không gian tổ chức.',
      );
    }
    if (actor.roleCode !== 'organization_owner') {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Chỉ chủ sở hữu tổ chức mới đổi trạng thái tổ chức.',
      );
    }
    const organization = await this.organizations.findOneByOrFail({ id: actor.organizationId });
    if (organization.status === status) {
      return { id: organization.id, name: organization.name, status: organization.status };
    }
    const now = new Date();
    const workspaceStatus = status === 'active' ? 'active' : status;
    await this.organizations.manager.transaction(async (manager) => {
      await manager.update(Organization, { id: organization.id }, { status, updatedAt: now });
      await manager.update(
        Workspace,
        { id: actor.workspaceId },
        { status: workspaceStatus, updatedAt: now },
      );
    });
    await this.recorder.append({
      workspaceId: actor.workspaceId,
      actorId: actor.userId,
      action: `organization.${status}`,
      targetType: 'organization',
      targetId: organization.id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { from: organization.status, to: status },
    });
    return { id: organization.id, name: organization.name, status };
  }
}
