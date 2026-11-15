export const PROFILE_PACKAGE_FORMAT: 'orcax.profile.v1';
export const PROFILE_STATUSES: readonly ['draft', 'in_review', 'approved', 'deprecated'];
export const PROFILE_HOSTS: readonly ['word', 'excel'];
export const PROFILE_SCOPES: readonly ['system', 'personal', 'organization'];

export interface ProfilePackageRule {
  code: string;
  version: string;
  title?: string;
  description?: string;
  checkKind: string;
  defaultSeverity: 'fail' | 'warning';
  checkConfig?: Record<string, unknown>;
  sortOrder?: number;
}

export interface ProfilePackageLegalReference {
  citation: string;
  sourceUrl?: string | null;
  verified?: boolean;
  notes?: string | null;
}

export interface ProfilePackage {
  format: typeof PROFILE_PACKAGE_FORMAT;
  profile: {
    code: string;
    name: string;
    host: 'word' | 'excel';
    scope?: 'system' | 'personal' | 'organization';
    source?: string | null;
    effectiveAt?: string | null;
  };
  rules: ProfilePackageRule[];
  legalReferences?: ProfilePackageLegalReference[];
}

export function validateProfilePackage(input: unknown): { ok: boolean; errors: string[] };

export function createProfilePackage(input: {
  profile: ProfilePackage['profile'];
  rules: ProfilePackageRule[];
  legalReferences?: ProfilePackageLegalReference[];
}): ProfilePackage;
