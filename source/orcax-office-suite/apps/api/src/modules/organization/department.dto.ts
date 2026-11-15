import { IsIn, IsOptional, IsString, IsUUID, MaxLength, MinLength } from 'class-validator';

export class CreateDepartmentDto {
  @IsString()
  @MinLength(2)
  @MaxLength(64)
  code!: string;

  @IsString()
  @MinLength(2)
  @MaxLength(160)
  name!: string;

  @IsOptional()
  @IsUUID('4')
  parentId?: string;
}

export class CreateGroupDto {
  @IsString()
  @MinLength(2)
  @MaxLength(64)
  code!: string;

  @IsString()
  @MinLength(2)
  @MaxLength(160)
  name!: string;
}

export class AssignMemberDto {
  @IsUUID('4', { message: 'Thành viên không hợp lệ.' })
  userId!: string;
}

export class OrganizationStatusDto {
  @IsIn(['active', 'suspended', 'archived'])
  status!: 'active' | 'suspended' | 'archived';
}

export class DistributeProfileDto {
  @IsIn(['department', 'group'])
  targetType!: 'department' | 'group';

  @IsUUID('4')
  targetId!: string;
}
