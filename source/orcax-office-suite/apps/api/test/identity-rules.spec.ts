import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  can,
  canAssignRole,
  decideLogin,
  permissionsFor,
  ROLE_PERMISSIONS,
  ROLES,
} from '../src/modules/authorization/catalog';
import { hashSecret, verifySecret } from '../src/modules/identity/password-hasher';
import {
  assertCanCreateOrganization,
  assertOrganizationMembershipActions,
  assertRecordInWorkspace,
  assertSwitchConfirmed,
  filterByWorkspace,
  loginWorkspace,
  requireWorkspaceId,
  WorkspaceBoundaryError,
} from '../src/tenancy/workspace-scope';

describe('identity and workspace rules', () => {
  it('does not reveal which part of a login failed', () => {
    assert.equal(decideLogin({ userFound: false, passwordMatches: false }), 'invalid');
    assert.equal(decideLogin({ userFound: false, passwordMatches: true }), 'invalid');
    assert.equal(
      decideLogin({ userFound: true, passwordMatches: false, status: 'active' }),
      'invalid',
    );
    assert.equal(
      decideLogin({ userFound: true, passwordMatches: true, status: 'pending' }),
      'unverified',
    );
    assert.equal(
      decideLogin({ userFound: true, passwordMatches: true, status: 'suspended' }),
      'suspended',
    );
    assert.equal(
      decideLogin({ userFound: true, passwordMatches: true, status: 'active' }),
      'allow',
    );
  });

  it('activates the personal workspace and only suggests another one', () => {
    const personal = '10000000-0000-4000-8000-000000000010';
    const organization = '10000000-0000-4000-8000-000000000011';
    assert.deepEqual(
      loginWorkspace({ personalWorkspaceId: personal, lastWorkspaceId: organization }),
      {
        activeWorkspaceId: personal,
        suggestedWorkspaceId: organization,
      },
    );
    assert.deepEqual(loginWorkspace({ personalWorkspaceId: personal, lastWorkspaceId: personal }), {
      activeWorkspaceId: personal,
      suggestedWorkspaceId: null,
    });
    assert.deepEqual(loginWorkspace({ personalWorkspaceId: personal, lastWorkspaceId: null }), {
      activeWorkspaceId: personal,
      suggestedWorkspaceId: null,
    });
  });

  it('keeps rows inside the active workspace', () => {
    const rows = [
      { workspaceId: 'a', name: 'one' },
      { workspaceId: 'b', name: 'two' },
    ];
    assert.deepEqual(filterByWorkspace('a', rows), [{ workspaceId: 'a', name: 'one' }]);
    assert.doesNotThrow(() => assertRecordInWorkspace('a', 'a'));
    assert.throws(() => assertRecordInWorkspace('a', 'b'), WorkspaceBoundaryError);
    assert.throws(() => assertSwitchConfirmed(false), WorkspaceBoundaryError);
    assert.doesNotThrow(() => assertSwitchConfirmed(true));
    assert.equal(requireWorkspaceId({ workspaceId: 'a' }), 'a');
    assert.throws(() => requireWorkspaceId({ workspaceId: '' }), WorkspaceBoundaryError);
  });

  it('creates organizations only from an active personal workspace', () => {
    assert.doesNotThrow(() =>
      assertCanCreateOrganization({
        workspaceKind: 'personal',
        membershipStatus: 'active',
        workspaceStatus: 'active',
      }),
    );
    assert.throws(
      () =>
        assertCanCreateOrganization({
          workspaceKind: 'organization',
          membershipStatus: 'active',
          workspaceStatus: 'active',
        }),
      WorkspaceBoundaryError,
    );
    assert.throws(() => assertOrganizationMembershipActions('personal'), WorkspaceBoundaryError);
    assert.doesNotThrow(() => assertOrganizationMembershipActions('organization'));
  });

  it('assigns organization roles only within the catalog', () => {
    assert.equal(canAssignRole('organization_owner', 'member'), true);
    assert.equal(canAssignRole('organization_administrator', 'member'), true);
    assert.equal(canAssignRole('organization_administrator', 'organization_owner'), false);
    assert.equal(canAssignRole('personal_owner', 'member'), false);
    assert.equal(canAssignRole('organization_owner', 'personal_owner'), false);
    assert.equal(can('personal_owner', 'workspace.member.invite'), false);
    assert.equal(can('personal_owner', 'workspace.audit.read'), true);
    assert.equal(can('auditor', 'workspace.audit.read'), true);
    assert.equal(can('member', 'workspace.audit.read'), false);
    assert.equal(can('member', 'workspace.update'), false);
    for (const role of ROLES) {
      assert.ok(permissionsFor(role.code).size === ROLE_PERMISSIONS[role.code].length);
    }
  });

  it('verifies a password hash and rejects a different password', async () => {
    const encoded = await hashSecret('correct-horse');
    assert.equal(await verifySecret('correct-horse', encoded), true);
    assert.equal(await verifySecret('other-horse', encoded), false);
    assert.equal(encoded.includes('correct-horse'), false);
  });
});
