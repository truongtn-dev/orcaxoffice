import { WorkspaceKind } from './workspace-scope';

export interface AccessContext {
  userId: string;
  sessionId: string;
  workspaceId: string;
  workspaceKind: WorkspaceKind;
  workspaceName: string;
  workspaceStatus: 'active' | 'suspended' | 'archived';
  membershipStatus: 'active' | 'suspended' | 'removed';
  roleCode: string;
  organizationId: string | null;
  email: string;
  correlationId: string;
}
