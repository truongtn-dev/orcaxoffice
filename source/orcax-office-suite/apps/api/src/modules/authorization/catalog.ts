export const PERMISSIONS = [
  { id: '20000000-0000-4000-8000-000000000001', code: 'workspace.read' },
  { id: '20000000-0000-4000-8000-000000000002', code: 'workspace.update' },
  { id: '20000000-0000-4000-8000-000000000003', code: 'workspace.member.invite' },
  { id: '20000000-0000-4000-8000-000000000004', code: 'workspace.member.suspend' },
  { id: '20000000-0000-4000-8000-000000000005', code: 'workspace.member.remove' },
  { id: '20000000-0000-4000-8000-000000000006', code: 'workspace.member.role' },
  { id: '20000000-0000-4000-8000-000000000007', code: 'workspace.audit.read' },
  { id: '20000000-0000-4000-8000-000000000008', code: 'session.self' },
  { id: '20000000-0000-4000-8000-000000000009', code: 'license.read' },
  { id: '20000000-0000-4000-8000-00000000000a', code: 'license.issue' },
  { id: '20000000-0000-4000-8000-00000000000b', code: 'license.manage' },
  { id: '20000000-0000-4000-8000-00000000000c', code: 'license.activate' },
  { id: '20000000-0000-4000-8000-00000000000d', code: 'profile.read' },
  { id: '20000000-0000-4000-8000-00000000000e', code: 'profile.write' },
  { id: '20000000-0000-4000-8000-00000000000f', code: 'profile.review' },
  { id: '20000000-0000-4000-8000-000000000010', code: 'profile.publish' },
  { id: '20000000-0000-4000-8000-000000000011', code: 'template.read' },
  { id: '20000000-0000-4000-8000-000000000012', code: 'template.write' },
  { id: '20000000-0000-4000-8000-000000000013', code: 'template.publish' },
  { id: '20000000-0000-4000-8000-000000000014', code: 'addin.word.use' },
  { id: '20000000-0000-4000-8000-000000000015', code: 'addin.excel.use' },
  { id: '20000000-0000-4000-8000-000000000016', code: 'workspace.department.manage' },
  { id: '20000000-0000-4000-8000-000000000017', code: 'workspace.config.read' },
  { id: '20000000-0000-4000-8000-000000000018', code: 'workspace.config.manage' },
  { id: '20000000-0000-4000-8000-000000000019', code: 'commerce.order.read' },
  { id: '20000000-0000-4000-8000-00000000001a', code: 'commerce.order.create' },
  { id: '20000000-0000-4000-8000-00000000001b', code: 'commerce.refund' },
  { id: '20000000-0000-4000-8000-00000000001c', code: 'commerce.reconcile' },
  { id: '20000000-0000-4000-8000-00000000001d', code: 'support.ticket.read' },
  { id: '20000000-0000-4000-8000-00000000001e', code: 'support.ticket.create' },
  { id: '20000000-0000-4000-8000-00000000001f', code: 'support.ticket.manage' },
  { id: '20000000-0000-4000-8000-000000000020', code: 'support.help.read' },
  { id: '20000000-0000-4000-8000-000000000021', code: 'support.help.manage' },
] as const;

export type PermissionCode = (typeof PERMISSIONS)[number]['code'];

export const ROLES = [
  { id: '10000000-0000-4000-8000-000000000001', code: 'personal_owner', kind: 'personal' },
  { id: '10000000-0000-4000-8000-000000000002', code: 'organization_owner', kind: 'organization' },
  {
    id: '10000000-0000-4000-8000-000000000003',
    code: 'organization_administrator',
    kind: 'organization',
  },
  { id: '10000000-0000-4000-8000-000000000004', code: 'member', kind: 'organization' },
  { id: '10000000-0000-4000-8000-000000000005', code: 'auditor', kind: 'organization' },
] as const;

export type RoleCode = (typeof ROLES)[number]['code'];

const ALL_ORGANIZATION = PERMISSIONS.map((item) => item.code);

export const ROLE_PERMISSIONS: Record<RoleCode, readonly PermissionCode[]> = {
  personal_owner: [
    'workspace.read',
    'workspace.audit.read',
    'session.self',
    'license.read',
    'license.issue',
    'license.manage',
    'license.activate',
    'profile.read',
    'profile.write',
    'profile.review',
    'profile.publish',
    'template.read',
    'template.write',
    'template.publish',
    'addin.word.use',
    'addin.excel.use',
    'workspace.config.read',
    'workspace.config.manage',
    'commerce.order.read',
    'commerce.order.create',
    'commerce.refund',
    'commerce.reconcile',
    'support.ticket.read',
    'support.ticket.create',
    'support.help.read',
    'support.help.manage',
  ],
  organization_owner: ALL_ORGANIZATION,
  organization_administrator: [
    'workspace.read',
    'workspace.update',
    'workspace.member.invite',
    'workspace.member.suspend',
    'workspace.member.remove',
    'workspace.member.role',
    'workspace.audit.read',
    'session.self',
    'license.read',
    'license.manage',
    'license.activate',
    'profile.read',
    'profile.write',
    'profile.review',
    'profile.publish',
    'template.read',
    'template.write',
    'template.publish',
    'addin.word.use',
    'addin.excel.use',
    'workspace.department.manage',
    'workspace.config.read',
    'workspace.config.manage',
    'commerce.order.read',
    'commerce.order.create',
    'commerce.refund',
    'commerce.reconcile',
    'support.ticket.read',
    'support.ticket.create',
    'support.ticket.manage',
    'support.help.read',
    'support.help.manage',
  ],
  member: [
    'workspace.read',
    'session.self',
    'license.read',
    'license.activate',
    'profile.read',
    'template.read',
    'addin.word.use',
    'addin.excel.use',
    'workspace.config.read',
    'commerce.order.read',
    'support.ticket.read',
    'support.ticket.create',
    'support.help.read',
  ],
  auditor: [
    'workspace.read',
    'workspace.audit.read',
    'session.self',
    'license.read',
    'profile.read',
    'template.read',
    'workspace.config.read',
    'commerce.order.read',
    'commerce.reconcile',
    'support.ticket.read',
    'support.help.read',
  ],
};

export function permissionsFor(roleCode: string): ReadonlySet<string> {
  const permissions = ROLE_PERMISSIONS[roleCode as RoleCode];
  return new Set(permissions ?? []);
}

export function can(roleCode: string, permission: string): boolean {
  return permissionsFor(roleCode).has(permission);
}

export function roleKind(roleCode: string): 'personal' | 'organization' | null {
  return ROLES.find((role) => role.code === roleCode)?.kind ?? null;
}

export function canAssignRole(actorRole: string, targetRole: string): boolean {
  if (roleKind(targetRole) !== 'organization') {
    return false;
  }
  if (actorRole === 'organization_owner') {
    return (
      targetRole === 'organization_owner' ||
      targetRole === 'organization_administrator' ||
      targetRole === 'member' ||
      targetRole === 'auditor'
    );
  }
  if (actorRole === 'organization_administrator') {
    return targetRole === 'member' || targetRole === 'auditor';
  }
  return false;
}

export type LoginDecision = 'allow' | 'invalid' | 'unverified' | 'suspended';

export function decideLogin(input: {
  userFound: boolean;
  passwordMatches: boolean;
  status?: 'pending' | 'active' | 'suspended';
}): LoginDecision {
  if (!input.userFound || !input.passwordMatches) {
    return 'invalid';
  }
  if (input.status === 'pending') {
    return 'unverified';
  }
  if (input.status === 'suspended') {
    return 'suspended';
  }
  return 'allow';
}

export const INVALID_CREDENTIALS_MESSAGE = 'Email hoặc mật khẩu không đúng.';
