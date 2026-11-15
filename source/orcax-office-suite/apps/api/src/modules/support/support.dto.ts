import { IsIn, IsNotEmpty, IsOptional, IsString, MaxLength, MinLength } from 'class-validator';
import { TicketCategory, TicketPriority, TicketStatus } from './support-ticket.entity';

export class CreateTicketDto {
  @IsString({ message: 'Tiêu đề phải là chuỗi ký tự.' })
  @IsNotEmpty({ message: 'Tiêu đề không được để trống.' })
  @MaxLength(255, { message: 'Tiêu đề tối đa 255 ký tự.' })
  title!: string;

  @IsString({ message: 'Nội dung phải là chuỗi ký tự.' })
  @IsNotEmpty({ message: 'Nội dung không được để trống.' })
  @MinLength(5, { message: 'Nội dung tối thiểu 5 ký tự.' })
  description!: string;

  @IsIn(['general', 'account', 'billing', 'licensing', 'word_addin', 'excel_addin'], {
    message: 'Danh mục không hợp lệ.',
  })
  category!: TicketCategory;

  @IsIn(['low', 'medium', 'high', 'urgent'], {
    message: 'Mức độ ưu tiên không hợp lệ.',
  })
  priority!: TicketPriority;
}

export class AddCommentDto {
  @IsString({ message: 'Nội dung phản hồi phải là chuỗi ký tự.' })
  @IsNotEmpty({ message: 'Nội dung phản hồi không được để trống.' })
  message!: string;
}

export class UpdateTicketStatusDto {
  @IsIn(['open', 'in_progress', 'resolved', 'closed'], {
    message: 'Trạng thái không hợp lệ.',
  })
  status!: TicketStatus;
}

export class CreateHelpArticleDto {
  @IsString({ message: 'Slug phải là chuỗi ký tự.' })
  @IsNotEmpty({ message: 'Slug không được để trống.' })
  @MaxLength(128, { message: 'Slug tối đa 128 ký tự.' })
  slug!: string;

  @IsString({ message: 'Tiêu đề phải là chuỗi ký tự.' })
  @IsNotEmpty({ message: 'Tiêu đề không được để trống.' })
  @MaxLength(255, { message: 'Tiêu đề tối đa 255 ký tự.' })
  title!: string;

  @IsString({ message: 'Danh mục phải là chuỗi ký tự.' })
  @IsNotEmpty({ message: 'Danh mục không được để trống.' })
  @MaxLength(64, { message: 'Danh mục tối đa 64 ký tự.' })
  category!: string;

  @IsString({ message: 'Nội dung phải là chuỗi ký tự.' })
  @IsNotEmpty({ message: 'Nội dung không được để trống.' })
  content!: string;

  @IsOptional()
  @IsString({ message: 'Phiên bản phải là chuỗi ký tự.' })
  @MaxLength(32, { message: 'Phiên bản tối đa 32 ký tự.' })
  productVersion?: string;
}
