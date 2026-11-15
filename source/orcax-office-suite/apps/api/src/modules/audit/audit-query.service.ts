import { HttpStatus, Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { ErrorCodes } from '@orcax/api-contracts';
import { Repository } from 'typeorm';
import { ApiError } from '../../common/http/api-error';
import { AccessContext } from '../../tenancy/access-context';
import { requireWorkspaceId } from '../../tenancy/workspace-scope';
import { AuthorizationService } from '../authorization/authorization.service';
import { AuditEvent } from './audit-event.entity';
import { AuditRecorder } from './audit-recorder.service';
import { AuditQueryDto } from './audit-query.dto';

@Injectable()
export class AuditQueryService {
  constructor(
    @InjectRepository(AuditEvent)
    private readonly events: Repository<AuditEvent>,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
  ) {}

  async search(actor: AccessContext, query: AuditQueryDto) {
    this.authorization.assertPolicy(actor, {
      permission: 'workspace.audit.read',
      requireActiveMembership: true,
    });
    const workspaceId = requireWorkspaceId(actor);
    const limit = query.limit ?? 50;
    const qb = this.events
      .createQueryBuilder('event')
      .where('event.workspace_id = :workspaceId', { workspaceId })
      .orderBy('event.occurred_at', 'DESC')
      .take(limit);

    if (query.action) {
      qb.andWhere('event.action = :action', { action: query.action });
    }
    if (query.actorId) {
      qb.andWhere('event.actor_id = :actorId', { actorId: query.actorId });
    }
    if (query.targetType) {
      qb.andWhere('event.target_type = :targetType', { targetType: query.targetType });
    }
    if (query.outcome) {
      qb.andWhere('event.outcome = :outcome', { outcome: query.outcome });
    }
    if (query.from) {
      const from = parseInstant(query.from, 'from');
      qb.andWhere('event.occurred_at >= :from', { from });
    }
    if (query.to) {
      const to = parseInstant(query.to, 'to');
      qb.andWhere('event.occurred_at <= :to', { to });
    }

    const rows = await qb.getMany();
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'audit.searched',
      targetType: 'audit_event',
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: {
        count: rows.length,
        filters: {
          action: query.action ?? null,
          actorId: query.actorId ?? null,
          targetType: query.targetType ?? null,
          outcome: query.outcome ?? null,
        },
      },
    });

    return {
      events: rows.map((row) => ({
        id: row.id,
        workspaceId: row.workspaceId,
        actorId: row.actorId,
        action: row.action,
        targetType: row.targetType,
        targetId: row.targetId,
        occurredAt: row.occurredAt.toISOString(),
        outcome: row.outcome,
        correlationId: row.correlationId,
        metadata: row.metadata,
      })),
    };
  }
}

function parseInstant(value: string, field: string): Date {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    throw new ApiError(
      HttpStatus.BAD_REQUEST,
      ErrorCodes.VALIDATION_FAILED,
      `Thời điểm ${field} không hợp lệ.`,
    );
  }
  return parsed;
}
