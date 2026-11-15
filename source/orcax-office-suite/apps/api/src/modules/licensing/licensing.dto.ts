import { IsString, IsUUID, MaxLength, MinLength } from 'class-validator';

export class IssueLicenseDto {
  @IsUUID('4', { message: 'Gói không hợp lệ.' })
  planId!: string;
}

export class ActivateLicenseDto {
  @IsUUID('4', { message: 'Giấy phép không hợp lệ.' })
  licenseId!: string;

  @IsString({ message: 'Khóa thiết bị không hợp lệ.' })
  @MinLength(8, { message: 'Khóa thiết bị cần ít nhất 8 ký tự.' })
  @MaxLength(128, { message: 'Khóa thiết bị tối đa 128 ký tự.' })
  deviceKey!: string;

  @IsString({ message: 'Tên thiết bị không hợp lệ.' })
  @MinLength(1, { message: 'Tên thiết bị không được để trống.' })
  @MaxLength(160, { message: 'Tên thiết bị tối đa 160 ký tự.' })
  deviceLabel!: string;
}

export class EntitlementCheckDto {
  @IsString({ message: 'Mã quyền lợi không hợp lệ.' })
  @MinLength(3, { message: 'Mã quyền lợi không hợp lệ.' })
  @MaxLength(64, { message: 'Mã quyền lợi tối đa 64 ký tự.' })
  code!: string;
}
