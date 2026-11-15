const API_VERSION = 'v1';
const CORRELATION_ID_HEADER = 'x-correlation-id';

const ErrorCodes = Object.freeze({
  VALIDATION_FAILED: 'VALIDATION_FAILED',
  BAD_REQUEST: 'BAD_REQUEST',
  UNAUTHORIZED: 'UNAUTHORIZED',
  FORBIDDEN: 'FORBIDDEN',
  NOT_FOUND: 'NOT_FOUND',
  CONFLICT: 'CONFLICT',
  INTERNAL_ERROR: 'INTERNAL_ERROR',
  SERVICE_UNAVAILABLE: 'SERVICE_UNAVAILABLE',
  EMAIL_NOT_VERIFIED: 'EMAIL_NOT_VERIFIED',
  WORKSPACE_CONFIRMATION_REQUIRED: 'WORKSPACE_CONFIRMATION_REQUIRED',
});

const AuditOutcomes = Object.freeze({
  SUCCESS: 'success',
  FAILURE: 'failure',
});

const AuditEventFields = Object.freeze([
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

const RuleSeverities = Object.freeze({
  PASS: 'pass',
  FAIL: 'fail',
  WARNING: 'warning',
});

function success(data, meta = {}) {
  return {
    data: data === undefined ? null : data,
    error: null,
    meta,
  };
}

function failure(error, meta = {}) {
  return {
    data: null,
    error: {
      code: error.code,
      message: error.message,
      details: error.details === undefined ? null : error.details,
    },
    meta,
  };
}

function optionalText(value, field) {
  if (value === undefined || value === null) {
    return null;
  }
  if (typeof value !== 'string') {
    throw new TypeError(`${field} must be a string`);
  }
  return value;
}

function createRuleResult(input) {
  if (!input || typeof input.ruleId !== 'string' || input.ruleId.trim() === '') {
    throw new TypeError('ruleId is required');
  }
  const allowed = Object.values(RuleSeverities);
  if (!allowed.includes(input.severity)) {
    throw new TypeError('severity is not supported');
  }
  return {
    ruleId: input.ruleId,
    severity: input.severity,
    location: optionalText(input.location, 'location'),
    explanation: optionalText(input.explanation, 'explanation'),
    correction: optionalText(input.correction, 'correction'),
  };
}

module.exports = {
  API_VERSION,
  CORRELATION_ID_HEADER,
  ErrorCodes,
  AuditOutcomes,
  AuditEventFields,
  RuleSeverities,
  success,
  failure,
  createRuleResult,
};
