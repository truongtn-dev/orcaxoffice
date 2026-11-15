import { AccessContext } from '../../tenancy/access-context';
import { WorkspaceKind } from '../../tenancy/workspace-scope';
import { can, PermissionCode } from './catalog';

export interface PolicyRequirement {
  permission: PermissionCode;
  workspaceKinds?: readonly WorkspaceKind[];
  /** When set to self, the actor must equal resourceOwnerId. */
  ownership?: 'self' | 'any';
  resourceOwnerId?: string | null;
  requireActiveMembership?: boolean;
}

export type PolicyDecision = 'allow' | 'deny_permission' | 'deny_attribute';

export function evaluatePolicy(
  access: AccessContext,
  requirement: PolicyRequirement,
): PolicyDecision {
  if (requirement.requireActiveMembership !== false) {
    if (access.membershipStatus !== 'active' || access.workspaceStatus !== 'active') {
      return 'deny_attribute';
    }
  }
  if (requirement.workspaceKinds && requirement.workspaceKinds.length > 0) {
    if (!requirement.workspaceKinds.includes(access.workspaceKind)) {
      return 'deny_attribute';
    }
  }
  if (requirement.ownership === 'self') {
    if (!requirement.resourceOwnerId || requirement.resourceOwnerId !== access.userId) {
      return 'deny_attribute';
    }
  }
  if (!can(access.roleCode, requirement.permission)) {
    return 'deny_permission';
  }
  return 'allow';
}

export function isPolicyAllowed(access: AccessContext, requirement: PolicyRequirement): boolean {
  return evaluatePolicy(access, requirement) === 'allow';
}
