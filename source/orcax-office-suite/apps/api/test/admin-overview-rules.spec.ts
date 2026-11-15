import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { can, ROLE_PERMISSIONS } from '../src/modules/authorization/catalog';

describe('admin portal permissions', () => {
  it('gives organization administrators department and config management', () => {
    assert.equal(can('organization_administrator', 'workspace.department.manage'), true);
    assert.equal(can('organization_administrator', 'workspace.config.manage'), true);
    assert.equal(can('organization_administrator', 'workspace.update'), true);
    assert.equal(can('organization_administrator', 'workspace.member.role'), true);
    assert.equal(can('member', 'workspace.department.manage'), false);
    assert.equal(can('auditor', 'workspace.config.read'), true);
  });

  it('keeps personal owner without department management', () => {
    assert.equal(can('personal_owner', 'workspace.department.manage'), false);
    assert.equal(can('personal_owner', 'workspace.config.manage'), true);
    assert.ok(ROLE_PERMISSIONS.organization_owner.includes('workspace.department.manage'));
  });
});
