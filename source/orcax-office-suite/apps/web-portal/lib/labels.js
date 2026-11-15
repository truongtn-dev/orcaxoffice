export const ROLE_LABELS = {
  personal_owner: 'Chủ không gian cá nhân',
  organization_owner: 'Chủ sở hữu',
  organization_administrator: 'Quản trị',
  member: 'Thành viên',
  auditor: 'Kiểm toán',
};

export const KIND_LABELS = {
  personal: 'Cá nhân',
  organization: 'Tổ chức',
};

export function roleLabel(code) {
  return ROLE_LABELS[code] ?? code;
}

export function kindLabel(kind) {
  return KIND_LABELS[kind] ?? kind;
}
