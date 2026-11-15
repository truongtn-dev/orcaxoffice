import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  approvalBlockedByLegalReferences,
  canTransitionProfile,
  canTransitionTemplate,
  nextProfileVersion,
  templateAvailableForGovernedUse,
} from '../src/modules/profile/profile-lifecycle';
import { SYSTEM_PROFILES, SYSTEM_TEMPLATES } from '../src/modules/profile/profile-catalog';

describe('profile lifecycle', () => {
  it('allows draft to review to approved to deprecated', () => {
    assert.equal(canTransitionProfile('draft', 'in_review'), true);
    assert.equal(canTransitionProfile('in_review', 'approved'), true);
    assert.equal(canTransitionProfile('approved', 'deprecated'), true);
    assert.equal(canTransitionProfile('draft', 'approved'), false);
    assert.equal(canTransitionProfile('deprecated', 'draft'), false);
  });

  it('blocks approval when any legal reference is unverified', () => {
    assert.equal(approvalBlockedByLegalReferences([]), false);
    assert.equal(approvalBlockedByLegalReferences([{ verified: true }]), false);
    assert.equal(approvalBlockedByLegalReferences([{ verified: true }, { verified: false }]), true);
  });

  it('increments immutable version numbers', () => {
    assert.equal(nextProfileVersion(1), 2);
  });

  it('seeds curated system profiles for word and excel', () => {
    assert.ok(SYSTEM_PROFILES.some((item) => item.host === 'word'));
    assert.ok(SYSTEM_PROFILES.some((item) => item.host === 'excel'));
    assert.ok(SYSTEM_PROFILES.every((item) => item.rules.length > 0));
    assert.ok(SYSTEM_PROFILES.every((item) => item.legalReferences.length > 0));
  });
});

describe('template lifecycle', () => {
  it('requires publish after approve for governed distribution', () => {
    assert.equal(canTransitionTemplate('draft', 'review'), true);
    assert.equal(canTransitionTemplate('review', 'approved'), true);
    assert.equal(canTransitionTemplate('approved', 'published'), true);
    assert.equal(templateAvailableForGovernedUse('approved'), true);
    assert.equal(templateAvailableForGovernedUse('published'), true);
    assert.equal(templateAvailableForGovernedUse('draft'), false);
    assert.ok(SYSTEM_TEMPLATES[0].fields.length >= 1);
    assert.ok(SYSTEM_TEMPLATES[0].mappings.length >= 1);
  });
});
