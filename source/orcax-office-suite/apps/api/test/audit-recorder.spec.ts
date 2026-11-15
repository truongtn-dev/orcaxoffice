import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { randomUUID } from 'crypto';
import { Repository } from 'typeorm';
import { AuditEvent } from '../src/modules/audit/audit-event.entity';
import { AuditRecordRejectedError } from '../src/modules/audit/audit-metadata';
import { AuditRecorder } from '../src/modules/audit/audit-recorder.service';

describe('AuditRecorder', () => {
  const correlationId = '22222222-2222-4222-8222-222222222222';

  function recorder() {
    const calls: Array<Record<string, unknown>> = [];
    const insert = async (row: Record<string, unknown>) => {
      calls.push(row);
    };
    const events = { insert } as unknown as Repository<AuditEvent>;
    return { calls, recorder: new AuditRecorder(events) };
  }

  it('inserts one append-only row and returns its id', async () => {
    const { calls, recorder: audit } = recorder();
    const id = await audit.append({
      action: 'foundation.probe',
      targetType: 'audit_event',
      outcome: 'success',
      correlationId,
      metadata: { before: { state: 'draft' }, after: { state: 'recorded' } },
    });
    assert.match(id, /^[0-9a-f-]{36}$/i);
    assert.equal(calls.length, 1);
    assert.equal(calls[0].id, id);
    assert.equal(calls[0].action, 'foundation.probe');
    assert.equal(calls[0].outcome, 'success');
    assert.equal(calls[0].correlationId, correlationId);
  });

  it('rejects passwords, tokens, and document bodies before insert', async () => {
    const { calls, recorder: audit } = recorder();
    await assert.rejects(
      () =>
        audit.append({
          action: 'foundation.probe',
          targetType: 'audit_event',
          outcome: 'failure',
          correlationId,
          metadata: { before: { password: 'hidden' } },
        }),
      AuditRecordRejectedError,
    );
    await assert.rejects(
      () =>
        audit.append({
          action: 'foundation.probe',
          targetType: 'audit_event',
          outcome: 'failure',
          correlationId,
          metadata: { documentBody: 'full word document' },
        }),
      /documentBody/,
    );
    assert.equal(calls.length, 0);
  });

  it('rejects an oversized metadata document', async () => {
    const { calls, recorder: audit } = recorder();
    await assert.rejects(
      () =>
        audit.append({
          action: 'foundation.probe',
          targetType: 'audit_event',
          outcome: 'success',
          correlationId,
          metadata: { note: 'x'.repeat(20_000) },
        }),
      /size limit/,
    );
    assert.equal(calls.length, 0);
  });

  it('rejects a correlation id that is not a UUID', async () => {
    const { recorder: audit } = recorder();
    await assert.rejects(
      () =>
        audit.append({
          action: 'foundation.probe',
          targetType: 'audit_event',
          outcome: 'success',
          correlationId: randomUUID().slice(0, 8),
        }),
      /correlationId/,
    );
  });
});
