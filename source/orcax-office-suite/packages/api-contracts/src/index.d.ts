export const API_VERSION: 'v1';
export const CORRELATION_ID_HEADER: 'x-correlation-id';

export const ErrorCodes: {
  readonly VALIDATION_FAILED: 'VALIDATION_FAILED';
  readonly BAD_REQUEST: 'BAD_REQUEST';
  readonly UNAUTHORIZED: 'UNAUTHORIZED';
  readonly FORBIDDEN: 'FORBIDDEN';
  readonly NOT_FOUND: 'NOT_FOUND';
  readonly CONFLICT: 'CONFLICT';
  readonly INTERNAL_ERROR: 'INTERNAL_ERROR';
  readonly SERVICE_UNAVAILABLE: 'SERVICE_UNAVAILABLE';
  readonly EMAIL_NOT_VERIFIED: 'EMAIL_NOT_VERIFIED';
  readonly WORKSPACE_CONFIRMATION_REQUIRED: 'WORKSPACE_CONFIRMATION_REQUIRED';
};

export const AuditOutcomes: {
  readonly SUCCESS: 'success';
  readonly FAILURE: 'failure';
};

export const AuditEventFields: readonly string[];

export const RuleSeverities: {
  readonly PASS: 'pass';
  readonly FAIL: 'fail';
  readonly WARNING: 'warning';
};

export interface ApiErrorBody {
  code: string;
  message: string;
  details: unknown;
}

export interface ApiEnvelope<T> {
  data: T | null;
  error: ApiErrorBody | null;
  meta: Record<string, unknown>;
}

export function success<T>(data: T, meta?: Record<string, unknown>): ApiEnvelope<T>;

export function failure(
  error: { code: string; message: string; details?: unknown },
  meta?: Record<string, unknown>,
): ApiEnvelope<null>;

export interface RuleResult {
  ruleId: string;
  severity: 'pass' | 'fail' | 'warning';
  location: string | null;
  explanation: string | null;
  correction: string | null;
}

export function createRuleResult(input: {
  ruleId: string;
  severity: string;
  location?: string | null;
  explanation?: string | null;
  correction?: string | null;
}): RuleResult;
