import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { resolveCorrelationId } from '../src/observability/correlation';

describe('correlation id', () => {
  const existing = '33333333-3333-4333-8333-333333333333';

  it('keeps a valid incoming id and an id already assigned to the request', () => {
    assert.equal(resolveCorrelationId(existing), existing);
    assert.equal(resolveCorrelationId('44444444-4444-4444-8444-444444444444', existing), existing);
  });

  it('replaces a header that could break the log line', () => {
    const resolved = resolveCorrelationId('not-a-uuid\ninjected');
    assert.equal(resolved.includes('\n'), false);
    assert.match(
      resolved,
      /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i,
    );
  });
});
