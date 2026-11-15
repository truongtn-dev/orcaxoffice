import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { AuditOutcomes } from '@orcax/api-contracts';
import { randomUUID } from 'crypto';
import { EntityManager, QueryDeepPartialEntity, Repository } from 'typeorm';
import { isUuid } from '../../common/ids/uuid';
import { AuditEvent } from './audit-event.entity';
import { assertSafeAuditMetadata, AuditRecordRejectedError } from './audit-metadata';

export interface AppendAuditEventInput {
  workspaceId?: string | null;
  actorId?: string | null;
  action: string;
  targetType: string;
  targetId?: string | null;
  outcome: 'success' | 'failure';
  correlationId: string;
  metadata?: Record<string, unknown> | null;
  occurredAt?: Date;
}

@Injectable()
export class AuditRecorder {
  constructor(
    @InjectRepository(AuditEvent)
    private readonly events: Repository<AuditEvent>,
  ) {}

  async append(input: AppendAuditEventInput, manager?: EntityManager): Promise<string> {
    assertLabel(input.action, 'action');
    assertLabel(input.targetType, 'targetType');
    assertOptionalLabel(input.targetId, 'targetId');
    assertOptionalUuid(input.workspaceId, 'workspaceId');
    assertOptionalUuid(input.actorId, 'actorId');
    if (!isUuid(input.correlationId)) {
      throw new AuditRecordRejectedError('correlationId must be a UUID');
    }
    if (input.outcome !== AuditOutcomes.SUCCESS && input.outcome !== AuditOutcomes.FAILURE) {
      throw new AuditRecordRejectedError('outcome is not supported');
    }
    assertSafeAuditMetadata(input.metadata ?? null);

    const id = randomUUID();
    const row: QueryDeepPartialEntity<AuditEvent> = {
      id,
      workspaceId: input.workspaceId ?? null,
      actorId: input.actorId ?? null,
      action: input.action,
      targetType: input.targetType,
      targetId: input.targetId ?? null,
      outcome: input.outcome,
      correlationId: input.correlationId,
      occurredAt: input.occurredAt ?? new Date(),
    };
    if (input.metadata) {
      row.metadata = input.metadata as QueryDeepPartialEntity<AuditEvent>['metadata'];
    }
    const events = manager ? manager.getRepository(AuditEvent) : this.events;
    await events.insert(row);
    return id;
  }
}

function assertLabel(value: string, field: string): void {
  if (typeof value !== 'string' || value.trim() === '' || value.length > 128) {
    throw new AuditRecordRejectedError(`${field} must be 1 to 128 characters`);
  }
}

function assertOptionalLabel(value: string | null | undefined, field: string): void {
  if (value === undefined || value === null) {
    return;
  }
  assertLabel(value, field);
}

function assertOptionalUuid(value: string | null | undefined, field: string): void {
  if (value === undefined || value === null) {
    return;
  }
  if (!isUuid(value)) {
    throw new AuditRecordRejectedError(`${field} must be a UUID`);
  }
}
