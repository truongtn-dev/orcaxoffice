import { HttpStatus, Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { ErrorCodes } from '@orcax/api-contracts';
import { randomUUID } from 'crypto';
import { Repository } from 'typeorm';
import { ApiError } from '../../common/http/api-error';
import { AccessContext } from '../../tenancy/access-context';
import { requireWorkspaceId } from '../../tenancy/workspace-scope';
import { AuditRecorder } from '../audit/audit-recorder.service';
import { AuthorizationService } from '../authorization/authorization.service';
import { CreateDepartmentDto, CreateGroupDto } from './department.dto';
import { DepartmentMember } from './department-member.entity';
import { Department } from './department.entity';
import { GroupMember } from './group-member.entity';
import { OrgGroup } from './org-group.entity';

@Injectable()
export class DepartmentService {
  constructor(
    @InjectRepository(Department)
    private readonly departments: Repository<Department>,
    @InjectRepository(OrgGroup)
    private readonly groups: Repository<OrgGroup>,
    @InjectRepository(DepartmentMember)
    private readonly departmentMembers: Repository<DepartmentMember>,
    @InjectRepository(GroupMember)
    private readonly groupMembers: Repository<GroupMember>,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
  ) {}

  async listDepartments(actor: AccessContext) {
    this.assertOrgAdmin(actor);
    const workspaceId = requireWorkspaceId(actor);
    const rows = await this.departments.find({
      where: { workspaceId },
      order: { name: 'ASC' },
    });
    const result = [];
    for (const row of rows) {
      const members = await this.departmentMembers.find({ where: { departmentId: row.id } });
      result.push({
        id: row.id,
        code: row.code,
        name: row.name,
        parentId: row.parentId,
        status: row.status,
        memberUserIds: members.map((item) => item.userId),
      });
    }
    return { departments: result };
  }

  async createDepartment(actor: AccessContext, body: CreateDepartmentDto) {
    this.assertOrgAdmin(actor);
    const workspaceId = requireWorkspaceId(actor);
    if (body.parentId) {
      const parent = await this.departments.findOne({
        where: { id: body.parentId, workspaceId },
      });
      if (!parent) {
        throw new ApiError(
          HttpStatus.BAD_REQUEST,
          ErrorCodes.BAD_REQUEST,
          'Phòng ban cha không hợp lệ.',
        );
      }
    }
    const existing = await this.departments.findOne({ where: { workspaceId, code: body.code } });
    if (existing) {
      throw new ApiError(HttpStatus.CONFLICT, ErrorCodes.CONFLICT, 'Mã phòng ban đã tồn tại.');
    }
    const now = new Date();
    const id = randomUUID();
    await this.departments.insert({
      id,
      workspaceId,
      parentId: body.parentId ?? null,
      code: body.code,
      name: body.name,
      status: 'active',
      createdAt: now,
      updatedAt: now,
    });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'department.created',
      targetType: 'department',
      targetId: id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { code: body.code },
    });
    return { id, code: body.code, name: body.name, parentId: body.parentId ?? null };
  }

  async assignDepartmentMember(actor: AccessContext, departmentId: string, userId: string) {
    this.assertOrgAdmin(actor);
    const workspaceId = requireWorkspaceId(actor);
    const department = await this.departments.findOne({ where: { id: departmentId, workspaceId } });
    if (!department) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy phòng ban.');
    }
    const existing = await this.departmentMembers.findOne({
      where: { departmentId, userId },
    });
    if (existing) {
      return { assigned: true };
    }
    await this.departmentMembers.insert({
      departmentId,
      userId,
      createdAt: new Date(),
    });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'department.member.assigned',
      targetType: 'department',
      targetId: departmentId,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { userId },
    });
    return { assigned: true };
  }

  async removeDepartmentMember(actor: AccessContext, departmentId: string, userId: string) {
    this.assertOrgAdmin(actor);
    const workspaceId = requireWorkspaceId(actor);
    const department = await this.departments.findOne({ where: { id: departmentId, workspaceId } });
    if (!department) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy phòng ban.');
    }
    await this.departmentMembers.delete({ departmentId, userId });
    return { removed: true };
  }

  async listGroups(actor: AccessContext) {
    this.assertOrgAdmin(actor);
    const workspaceId = requireWorkspaceId(actor);
    const rows = await this.groups.find({ where: { workspaceId }, order: { name: 'ASC' } });
    const result = [];
    for (const row of rows) {
      const members = await this.groupMembers.find({ where: { groupId: row.id } });
      result.push({
        id: row.id,
        code: row.code,
        name: row.name,
        status: row.status,
        memberUserIds: members.map((item) => item.userId),
      });
    }
    return { groups: result };
  }

  async createGroup(actor: AccessContext, body: CreateGroupDto) {
    this.assertOrgAdmin(actor);
    const workspaceId = requireWorkspaceId(actor);
    const existing = await this.groups.findOne({ where: { workspaceId, code: body.code } });
    if (existing) {
      throw new ApiError(HttpStatus.CONFLICT, ErrorCodes.CONFLICT, 'Mã nhóm đã tồn tại.');
    }
    const now = new Date();
    const id = randomUUID();
    await this.groups.insert({
      id,
      workspaceId,
      code: body.code,
      name: body.name,
      status: 'active',
      createdAt: now,
      updatedAt: now,
    });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'group.created',
      targetType: 'group',
      targetId: id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { code: body.code },
    });
    return { id, code: body.code, name: body.name };
  }

  async assignGroupMember(actor: AccessContext, groupId: string, userId: string) {
    this.assertOrgAdmin(actor);
    const workspaceId = requireWorkspaceId(actor);
    const group = await this.groups.findOne({ where: { id: groupId, workspaceId } });
    if (!group) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy nhóm.');
    }
    const existing = await this.groupMembers.findOne({ where: { groupId, userId } });
    if (existing) {
      return { assigned: true };
    }
    await this.groupMembers.insert({ groupId, userId, createdAt: new Date() });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'group.member.assigned',
      targetType: 'group',
      targetId: groupId,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { userId },
    });
    return { assigned: true };
  }

  async removeGroupMember(actor: AccessContext, groupId: string, userId: string) {
    this.assertOrgAdmin(actor);
    const workspaceId = requireWorkspaceId(actor);
    const group = await this.groups.findOne({ where: { id: groupId, workspaceId } });
    if (!group) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy nhóm.');
    }
    await this.groupMembers.delete({ groupId, userId });
    return { removed: true };
  }

  async userTargets(workspaceId: string, userId: string) {
    const departmentRows = await this.departmentMembers.find({ where: { userId } });
    const groupRows = await this.groupMembers.find({ where: { userId } });
    const departmentIds: string[] = [];
    for (const row of departmentRows) {
      const department = await this.departments.findOne({
        where: { id: row.departmentId, workspaceId },
      });
      if (department) {
        departmentIds.push(department.id);
      }
    }
    const groupIds: string[] = [];
    for (const row of groupRows) {
      const group = await this.groups.findOne({ where: { id: row.groupId, workspaceId } });
      if (group) {
        groupIds.push(group.id);
      }
    }
    return { departmentIds, groupIds };
  }

  private assertOrgAdmin(actor: AccessContext) {
    this.authorization.assertPolicy(actor, {
      permission: 'workspace.department.manage',
      workspaceKinds: ['organization'],
    });
  }
}
