import { IsEmail, IsIn, IsString, MaxLength, MinLength } from 'class-validator';

const ASSIGNABLE_ROLES = [
  'organization_owner',
  'organization_administrator',
  'member',
  'auditor',
] as const;

export class CreateOrganizationDto {
  @IsString({ message: 'Tên tổ chức không hợp lệ.' })
  @MinLength(2, { message: 'Tên tổ chức cần từ 2 đến 160 ký tự.' })
  @MaxLength(160, { message: 'Tên tổ chức cần từ 2 đến 160 ký tự.' })
  name!: string;
}

export class InviteDto {
  @IsEmail({}, { message: 'Email không hợp lệ.' })
  email!: string;

  @IsIn(ASSIGNABLE_ROLES, { message: 'Vai trò không hợp lệ.' })
  role!: (typeof ASSIGNABLE_ROLES)[number];
}

export class AcceptInvitationDto {
  @IsString({ message: 'Mã lời mời không hợp lệ.' })
  @MinLength(20, { message: 'Mã lời mời không hợp lệ.' })
  @MaxLength(200, { message: 'Mã lời mời không hợp lệ.' })
  token!: string;
}

export class RoleChangeDto {
  @IsIn(ASSIGNABLE_ROLES, { message: 'Vai trò không hợp lệ.' })
  role!: (typeof ASSIGNABLE_ROLES)[number];
}
