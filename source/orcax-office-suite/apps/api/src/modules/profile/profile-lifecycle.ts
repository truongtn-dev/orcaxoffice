export type ProfileStatus = 'draft' | 'in_review' | 'approved' | 'deprecated';

export type ProfileHost = 'word' | 'excel';

export type ProfileScope = 'system' | 'personal' | 'organization';

const TRANSITIONS: Record<ProfileStatus, readonly ProfileStatus[]> = {
  draft: ['in_review'],
  in_review: ['draft', 'approved'],
  approved: ['deprecated'],
  deprecated: [],
};

export function canTransitionProfile(from: ProfileStatus, to: ProfileStatus): boolean {
  return TRANSITIONS[from].includes(to);
}

export function approvalBlockedByLegalReferences(
  references: readonly { verified: boolean }[],
): boolean {
  return references.some((item) => !item.verified);
}

export function assertProfileMutable(status: ProfileStatus): boolean {
  return status === 'draft' || status === 'in_review';
}

export function nextProfileVersion(current: number): number {
  return current + 1;
}

export type TemplateStatus = 'draft' | 'review' | 'approved' | 'published' | 'deprecated';

const TEMPLATE_TRANSITIONS: Record<TemplateStatus, readonly TemplateStatus[]> = {
  draft: ['review'],
  review: ['draft', 'approved'],
  approved: ['published', 'deprecated'],
  published: ['deprecated'],
  deprecated: [],
};

export function canTransitionTemplate(from: TemplateStatus, to: TemplateStatus): boolean {
  return TEMPLATE_TRANSITIONS[from].includes(to);
}

export function templateAvailableForGovernedUse(status: TemplateStatus): boolean {
  return status === 'approved' || status === 'published';
}
