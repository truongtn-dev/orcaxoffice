import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { AccessContext } from '../src/tenancy/access-context';
import { evaluatePolicy } from '../src/modules/authorization/policy';

function access(overrides: Partial<AccessContext> = {}): AccessContext {
  return {
    userId: '11111111-1111-4111-8111-111111111111',
    sessionId: '22222222-2222-4222-8222-222222222222',
    workspaceId: '33333333-3333-4333-8333-333333333333',
    workspaceKind: 'organization',
    workspaceName: 'Công ty',
    workspaceStatus: 'active',
    membershipStatus: 'active',
    roleCode: 'member',
    organizationId: '44444444-4444-4444-8444-444444444444',
    email: 'member@example.com',
    correlationId: '55555555-5555-4555-8555-555555555555',
    ...overrides,
  };
}

describe('authorization policy', () => {
  it('allows an RBAC match and denies a missing permission', () => {
    assert.equal(
      evaluatePolicy(access({ roleCode: 'auditor' }), { permission: 'workspace.audit.read' }),
      'allow',
    );
    assert.equal(
      evaluatePolicy(access({ roleCode: 'member' }), { permission: 'workspace.audit.read' }),
      'deny_permission',
    );
  });

  it('denies when workspace kind or membership is wrong', () => {
    assert.equal(
      evaluatePolicy(access({ workspaceKind: 'personal', roleCode: 'organization_owner' }), {
        permission: 'workspace.member.invite',
        workspaceKinds: ['organization'],
      }),
      'deny_attribute',
    );
    assert.equal(
      evaluatePolicy(access({ membershipStatus: 'suspended', roleCode: 'organization_owner' }), {
        permission: 'workspace.member.invite',
        workspaceKinds: ['organization'],
      }),
      'deny_attribute',
    );
  });

  it('enforces ownership when requested', () => {
    const actor = access({ userId: 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa' });
    assert.equal(
      evaluatePolicy(actor, {
        permission: 'session.self',
        ownership: 'self',
        resourceOwnerId: actor.userId,
      }),
      'allow',
    );
    assert.equal(
      evaluatePolicy(actor, {
        permission: 'session.self',
        ownership: 'self',
        resourceOwnerId: 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb',
      }),
      'deny_attribute',
    );
  });
});
