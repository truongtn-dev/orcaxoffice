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
import { can } from '../authorization/catalog';
import { WorkspaceSettingHistory } from './workspace-setting-history.entity';
import { WorkspaceSetting } from './workspace-setting.entity';

const ALLOWED_KEYS = new Set([
  'portal.locale',
  'portal.timezone',
  'notifications.email.enabled',
  'security.session.hint',
]);

@Injectable()
export class ConfigurationService {
  constructor(
    @InjectRepository(WorkspaceSetting)
    private readonly settings: Repository<WorkspaceSetting>,
    @InjectRepository(WorkspaceSettingHistory)
    private readonly history: Repository<WorkspaceSettingHistory>,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
  ) {}

  async list(actor: AccessContext) {
    this.authorization.assertPolicy(actor, { permission: 'workspace.config.read' });
    const workspaceId = requireWorkspaceId(actor);
    const rows = await this.settings.find({
      where: { workspaceId },
      order: { key: 'ASC' },
    });
    return {
      settings: rows.map((row) => ({
        key: row.key,
        value: row.value,
        version: row.version,
        updatedAt: row.updatedAt.toISOString(),
        updatedBy: row.updatedBy,
      })),
    };
  }

  async upsert(actor: AccessContext, key: string, value: Record<string, unknown>) {
    this.authorization.assertPolicy(actor, { permission: 'workspace.config.manage' });
    if (!ALLOWED_KEYS.has(key)) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Khóa cấu hình không được hỗ trợ.',
      );
    }
    const workspaceId = requireWorkspaceId(actor);
    const now = new Date();
    const existing = await this.settings.findOne({ where: { workspaceId, key } });
    const nextVersion = existing ? existing.version + 1 : 1;
    const id = existing?.id ?? randomUUID();
    await this.settings.manager.transaction(async (manager) => {
      if (existing) {
        await manager.update(
          WorkspaceSetting,
          { id },
          { value: { ...value }, version: nextVersion, updatedBy: actor.userId, updatedAt: now },
        );
      } else {
        await manager.insert(WorkspaceSetting, {
          id,
          workspaceId,
          key,
          value: { ...value },
          version: nextVersion,
          updatedBy: actor.userId,
          updatedAt: now,
        });
      }
      await manager.insert(WorkspaceSettingHistory, {
        id: randomUUID(),
        workspaceId,
        key,
        value: { ...value },
        version: nextVersion,
        changedBy: actor.userId,
        changedAt: now,
      });
    });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'workspace.config.updated',
      targetType: 'workspace_setting',
      targetId: id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { key, version: nextVersion },
    });
    return { key, value, version: nextVersion };
  }

  async historyFor(actor: AccessContext, key: string) {
    this.authorization.assertPolicy(actor, { permission: 'workspace.config.read' });
    const workspaceId = requireWorkspaceId(actor);
    const rows = await this.history.find({
      where: { workspaceId, key },
      order: { version: 'DESC' },
      take: 20,
    });
    return {
      history: rows.map((row) => ({
        version: row.version,
        value: row.value,
        changedBy: row.changedBy,
        changedAt: row.changedAt.toISOString(),
      })),
    };
  }

  canManage(roleCode: string) {
    return can(roleCode, 'workspace.config.manage');
  }
}
