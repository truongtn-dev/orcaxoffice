function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

export function isApiEnvelope(value: unknown): boolean {
  return isRecord(value) && 'data' in value && 'error' in value && 'meta' in value;
}
