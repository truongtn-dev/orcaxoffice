export type WorkspaceKind = 'personal' | 'organization';

export class WorkspaceBoundaryError extends Error {
  constructor(message = 'Bản ghi không thuộc không gian làm việc đang hoạt động.') {
    super(message);
    this.name = 'WorkspaceBoundaryError';
  }
}

export function requireWorkspaceId(access: { workspaceId: string }): string {
  if (typeof access.workspaceId !== 'string' || access.workspaceId.length === 0) {
    throw new WorkspaceBoundaryError('Thiếu không gian làm việc đang hoạt động.');
  }
  return access.workspaceId;
}

export function assertRecordInWorkspace(
  activeWorkspaceId: string,
  recordWorkspaceId: string,
): void {
  if (activeWorkspaceId !== recordWorkspaceId) {
    throw new WorkspaceBoundaryError();
  }
}

export function filterByWorkspace<T extends { workspaceId: string }>(
  workspaceId: string,
  rows: readonly T[],
): T[] {
  return rows.filter((row) => row.workspaceId === workspaceId);
}

export function loginWorkspace(input: {
  personalWorkspaceId: string;
  lastWorkspaceId: string | null;
}): { activeWorkspaceId: string; suggestedWorkspaceId: string | null } {
  const suggested =
    input.lastWorkspaceId && input.lastWorkspaceId !== input.personalWorkspaceId
      ? input.lastWorkspaceId
      : null;
  return {
    activeWorkspaceId: input.personalWorkspaceId,
    suggestedWorkspaceId: suggested,
  };
}

export function assertSwitchConfirmed(confirmed: boolean): void {
  if (confirmed !== true) {
    throw new WorkspaceBoundaryError('Cần xác nhận trước khi chuyển không gian làm việc.');
  }
}

export function assertCanCreateOrganization(access: {
  workspaceKind: WorkspaceKind;
  membershipStatus: string;
  workspaceStatus: string;
}): void {
  if (
    access.workspaceKind !== 'personal' ||
    access.membershipStatus !== 'active' ||
    access.workspaceStatus !== 'active'
  ) {
    throw new WorkspaceBoundaryError(
      'Chỉ tạo tổ chức khi đang ở không gian cá nhân đang hoạt động.',
    );
  }
}

export function assertOrganizationMembershipActions(workspaceKind: WorkspaceKind): void {
  if (workspaceKind !== 'organization') {
    throw new WorkspaceBoundaryError('Không gian cá nhân không có thành viên khác.');
  }
}
