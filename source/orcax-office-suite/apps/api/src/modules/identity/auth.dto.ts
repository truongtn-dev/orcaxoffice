import { IsBoolean, IsEmail, IsString, IsUUID, MaxLength, MinLength } from 'class-validator';

export class RegisterDto {
  @IsEmail({}, { message: 'Email không hợp lệ.' })
  email!: string;

  @IsString({ message: 'Mật khẩu không hợp lệ.' })
  @MinLength(10, { message: 'Mật khẩu cần ít nhất 10 ký tự.' })
  @MaxLength(72, { message: 'Mật khẩu tối đa 72 ký tự.' })
  password!: string;

  @IsString({ message: 'Tên hiển thị không hợp lệ.' })
  @MinLength(1, { message: 'Tên hiển thị không được để trống.' })
  @MaxLength(120, { message: 'Tên hiển thị tối đa 120 ký tự.' })
  displayName!: string;
}

export class LoginDto {
  @IsEmail({}, { message: 'Email không hợp lệ.' })
  email!: string;

  @IsString({ message: 'Mật khẩu không hợp lệ.' })
  @MinLength(1, { message: 'Mật khẩu không được để trống.' })
  @MaxLength(72, { message: 'Mật khẩu tối đa 72 ký tự.' })
  password!: string;
}

export class EmailCodeDto {
  @IsEmail({}, { message: 'Email không hợp lệ.' })
  email!: string;

  @IsString({ message: 'Mã không hợp lệ.' })
  @MinLength(6, { message: 'Mã gồm 6 chữ số.' })
  @MaxLength(6, { message: 'Mã gồm 6 chữ số.' })
  code!: string;
}

export class EmailDto {
  @IsEmail({}, { message: 'Email không hợp lệ.' })
  email!: string;
}

export class ResetPasswordDto extends EmailCodeDto {
  @IsString({ message: 'Mật khẩu không hợp lệ.' })
  @MinLength(10, { message: 'Mật khẩu cần ít nhất 10 ký tự.' })
  @MaxLength(72, { message: 'Mật khẩu tối đa 72 ký tự.' })
  password!: string;
}

export class ProfileDto {
  @IsString({ message: 'Tên hiển thị không hợp lệ.' })
  @MinLength(1, { message: 'Tên hiển thị không được để trống.' })
  @MaxLength(120, { message: 'Tên hiển thị tối đa 120 ký tự.' })
  displayName!: string;
}

export class SwitchWorkspaceDto {
  @IsUUID('4', { message: 'Không gian làm việc không hợp lệ.' })
  workspaceId!: string;

  @IsBoolean({ message: 'Cần xác nhận trước khi chuyển không gian làm việc.' })
  confirmed!: boolean;
}
