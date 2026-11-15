export type { AccessContext } from './access-context';
export {
  assertCanCreateOrganization,
  assertOrganizationMembershipActions,
  assertRecordInWorkspace,
  assertSwitchConfirmed,
  filterByWorkspace,
  loginWorkspace,
  requireWorkspaceId,
  WorkspaceBoundaryError,
  type WorkspaceKind,
} from './workspace-scope';
