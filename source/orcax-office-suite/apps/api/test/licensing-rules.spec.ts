import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  deriveLicenseStatus,
  licenseAllowsUse,
  PLANS,
} from '../src/modules/licensing/license-catalog';

describe('licensing rules', () => {
  it('derives active, grace, expired, revoked, and pending states', () => {
    const now = new Date('2026-06-01T00:00:00.000Z');
    assert.equal(
      deriveLicenseStatus({
        now,
        startsAt: new Date('2026-01-01T00:00:00.000Z'),
        endsAt: new Date('2026-12-01T00:00:00.000Z'),
        revokedAt: null,
        graceDays: 14,
      }),
      'active',
    );
    assert.equal(
      deriveLicenseStatus({
        now,
        startsAt: new Date('2026-01-01T00:00:00.000Z'),
        endsAt: new Date('2026-05-25T00:00:00.000Z'),
        revokedAt: null,
        graceDays: 14,
      }),
      'grace',
    );
    assert.equal(
      deriveLicenseStatus({
        now,
        startsAt: new Date('2026-01-01T00:00:00.000Z'),
        endsAt: new Date('2026-05-01T00:00:00.000Z'),
        revokedAt: null,
        graceDays: 14,
      }),
      'expired',
    );
    assert.equal(
      deriveLicenseStatus({
        now,
        startsAt: new Date('2026-01-01T00:00:00.000Z'),
        endsAt: null,
        revokedAt: new Date('2026-05-01T00:00:00.000Z'),
        graceDays: 14,
      }),
      'revoked',
    );
    assert.equal(
      deriveLicenseStatus({
        now,
        startsAt: new Date('2026-07-01T00:00:00.000Z'),
        endsAt: null,
        revokedAt: null,
        graceDays: 0,
      }),
      'pending',
    );
  });

  it('allows use only while active or in grace', () => {
    assert.equal(licenseAllowsUse('active'), true);
    assert.equal(licenseAllowsUse('grace'), true);
    assert.equal(licenseAllowsUse('expired'), false);
    assert.equal(licenseAllowsUse('revoked'), false);
    assert.equal(licenseAllowsUse('pending'), false);
  });

  it('keeps personal and organization plans separate', () => {
    assert.ok(PLANS.some((plan) => plan.workspaceKind === 'personal'));
    assert.ok(PLANS.some((plan) => plan.workspaceKind === 'organization'));
    assert.ok(PLANS.every((plan) => plan.deviceLimit >= 1));
  });
});
