import { MembershipStatus } from './membership.entity';
import { WorkspaceKind, WorkspaceStatus } from './workspace.entity';

export interface WorkspaceAccess {
  membershipId: string;
  workspaceId: string;
  workspaceKind: WorkspaceKind;
  workspaceName: string;
  workspaceStatus: WorkspaceStatus;
  membershipStatus: MembershipStatus;
  roleCode: string;
  organizationId: string | null;
}
