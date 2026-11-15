export const ENTITLEMENTS = [
  { code: 'word.core', name: 'Tự động hóa Word cơ bản' },
  { code: 'excel.core', name: 'Tối ưu in Excel cơ bản' },
  { code: 'portal.workspace', name: 'Cổng thông tin workspace' },
] as const;

export type EntitlementCode = (typeof ENTITLEMENTS)[number]['code'];

export const PRODUCTS = [
  {
    id: '30000000-0000-4000-8000-000000000001',
    code: 'orcax-office',
    name: 'OrcaX Office Suite',
  },
] as const;

export const PLANS = [
  {
    id: '31000000-0000-4000-8000-000000000001',
    productId: PRODUCTS[0].id,
    code: 'personal-free',
    name: 'Cá nhân miễn phí',
    licenseType: 'subscription' as const,
    workspaceKind: 'personal' as const,
    deviceLimit: 2,
    seatLimit: 1,
    termDays: 3650,
    graceDays: 14,
    entitlements: ['word.core', 'excel.core', 'portal.workspace'] as const,
  },
  {
    id: '31000000-0000-4000-8000-000000000002',
    productId: PRODUCTS[0].id,
    code: 'organization-starter',
    name: 'Tổ chức khởi đầu',
    licenseType: 'subscription' as const,
    workspaceKind: 'organization' as const,
    deviceLimit: 10,
    seatLimit: 25,
    termDays: 365,
    graceDays: 14,
    entitlements: ['word.core', 'excel.core', 'portal.workspace'] as const,
  },
  {
    id: '31000000-0000-4000-8000-000000000003',
    productId: PRODUCTS[0].id,
    code: 'organization-perpetual',
    name: 'Tổ chức vĩnh viễn',
    licenseType: 'perpetual' as const,
    workspaceKind: 'organization' as const,
    deviceLimit: 25,
    seatLimit: 100,
    termDays: null,
    graceDays: 0,
    entitlements: ['word.core', 'excel.core', 'portal.workspace'] as const,
  },
] as const;

export type LicenseRuntimeStatus = 'pending' | 'active' | 'grace' | 'expired' | 'revoked';

export function deriveLicenseStatus(input: {
  now: Date;
  startsAt: Date;
  endsAt: Date | null;
  revokedAt: Date | null;
  graceDays: number;
}): LicenseRuntimeStatus {
  if (input.revokedAt) {
    return 'revoked';
  }
  if (input.startsAt > input.now) {
    return 'pending';
  }
  if (!input.endsAt) {
    return 'active';
  }
  if (input.endsAt >= input.now) {
    return 'active';
  }
  const graceEnds = new Date(
    input.endsAt.getTime() + Math.max(0, input.graceDays) * 24 * 60 * 60 * 1000,
  );
  if (graceEnds >= input.now) {
    return 'grace';
  }
  return 'expired';
}

export function licenseAllowsUse(status: LicenseRuntimeStatus): boolean {
  return status === 'active' || status === 'grace';
}
