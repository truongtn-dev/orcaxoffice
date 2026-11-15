import { IsIn, IsInt, IsOptional, IsString, IsUUID, Max, MaxLength, Min } from 'class-validator';
import { Type } from 'class-transformer';

export class AuditQueryDto {
  @IsOptional()
  @IsString({ message: 'Hành động không hợp lệ.' })
  @MaxLength(128, { message: 'Hành động tối đa 128 ký tự.' })
  action?: string;

  @IsOptional()
  @IsUUID('4', { message: 'Người thực hiện không hợp lệ.' })
  actorId?: string;

  @IsOptional()
  @IsString({ message: 'Loại đối tượng không hợp lệ.' })
  @MaxLength(128, { message: 'Loại đối tượng tối đa 128 ký tự.' })
  targetType?: string;

  @IsOptional()
  @IsIn(['success', 'failure'], { message: 'Kết quả không hợp lệ.' })
  outcome?: 'success' | 'failure';

  @IsOptional()
  @IsString({ message: 'Thời điểm bắt đầu không hợp lệ.' })
  from?: string;

  @IsOptional()
  @IsString({ message: 'Thời điểm kết thúc không hợp lệ.' })
  to?: string;

  @IsOptional()
  @Type(() => Number)
  @IsInt({ message: 'Giới hạn phải là số nguyên.' })
  @Min(1, { message: 'Giới hạn tối thiểu là 1.' })
  @Max(100, { message: 'Giới hạn tối đa là 100.' })
  limit?: number;
}
