const test = require('node:test');
const assert = require('node:assert/strict');
const {
  API_VERSION,
  AuditEventFields,
  ErrorCodes,
  RuleSeverities,
  createRuleResult,
  failure,
  success,
} = require('./index');

test('success envelope keeps data and an empty error', () => {
  const body = success({ ok: true }, { correlationId: 'abc', apiVersion: API_VERSION });
  assert.equal(body.error, null);
  assert.deepEqual(body.data, { ok: true });
  assert.equal(body.meta.apiVersion, 'v1');
});

test('failure envelope always includes a details field', () => {
  const body = failure({ code: ErrorCodes.NOT_FOUND, message: 'Không tìm thấy bản ghi.' });
  assert.equal(body.data, null);
  assert.equal(body.error.code, 'NOT_FOUND');
  assert.equal(body.error.details, null);
});

test('audit event fields cover the approved record shape', () => {
  assert.deepEqual(AuditEventFields, [
    'id',
    'workspaceId',
    'actorId',
    'action',
    'targetType',
    'targetId',
    'occurredAt',
    'outcome',
    'correlationId',
    'metadata',
  ]);
});

test('rule result accepts the shared severities and rejects an unknown one', () => {
  const result = createRuleResult({
    ruleId: 'sample.rule',
    severity: RuleSeverities.WARNING,
    explanation: 'Cần xem lại khoảng cách dòng.',
  });
  assert.equal(result.severity, 'warning');
  assert.equal(result.location, null);
  assert.throws(() => createRuleResult({ ruleId: 'sample.rule', severity: 'info' }), TypeError);
});
