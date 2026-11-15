const FORBIDDEN_METADATA_KEYS = new Set([
  'password',
  'passwd',
  'secret',
  'token',
  'accesstoken',
  'refreshtoken',
  'authorization',
  'cookie',
  'documentbody',
  'documentcontent',
  'workbookcontent',
  'filecontent',
  'privatekey',
  'cardnumber',
  'cvv',
]);

const MAX_METADATA_BYTES = 16 * 1024;

export class AuditRecordRejectedError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'AuditRecordRejectedError';
  }
}

function normalizeKey(key: string): string {
  return key.toLowerCase().replace(/[^a-z0-9]/g, '');
}

function walk(value: unknown, seen: WeakSet<object>): void {
  if (typeof value !== 'object' || value === null) {
    return;
  }
  if (seen.has(value)) {
    throw new AuditRecordRejectedError('Audit metadata must not contain circular references');
  }
  seen.add(value);
  if (Array.isArray(value)) {
    for (const item of value) {
      walk(item, seen);
    }
    return;
  }
  for (const [key, nested] of Object.entries(value)) {
    if (key === '__proto__' || key === 'constructor' || key === 'prototype') {
      throw new AuditRecordRejectedError('Audit metadata contains an unsupported key');
    }
    if (FORBIDDEN_METADATA_KEYS.has(normalizeKey(key))) {
      throw new AuditRecordRejectedError(`Audit metadata must not contain ${key}`);
    }
    walk(nested, seen);
  }
}

export function assertSafeAuditMetadata(metadata: Record<string, unknown> | null): void {
  if (metadata === null) {
    return;
  }
  if (typeof metadata !== 'object' || Array.isArray(metadata)) {
    throw new AuditRecordRejectedError('Audit metadata must be an object');
  }
  walk(metadata, new WeakSet());
  let encoded: string;
  try {
    encoded = JSON.stringify(metadata);
  } catch {
    throw new AuditRecordRejectedError('Audit metadata must be JSON-serializable');
  }
  if (new TextEncoder().encode(encoded).length > MAX_METADATA_BYTES) {
    throw new AuditRecordRejectedError('Audit metadata exceeds the size limit');
  }
}
