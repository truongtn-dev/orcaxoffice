import { Type } from 'class-transformer';
import {
  IsBoolean,
  IsIn,
  IsInt,
  IsOptional,
  IsString,
  IsUUID,
  Max,
  MaxLength,
  Min,
  MinLength,
  ValidateNested,
} from 'class-validator';

export class ValidationSummaryDto {
  @IsInt()
  @Min(0)
  @Max(100000)
  pass!: number;

  @IsInt()
  @Min(0)
  @Max(100000)
  fail!: number;

  @IsInt()
  @Min(0)
  @Max(100000)
  warning!: number;

  @IsInt()
  @Min(0)
  @Max(100000)
  corrected!: number;
}

export class RecordValidationRunDto {
  @IsUUID('4')
  profileId!: string;

  @IsUUID('4')
  profileVersionId!: string;

  @IsInt()
  @Min(1)
  profileVersion!: number;

  @IsString()
  @MinLength(2)
  @MaxLength(64)
  profileCode!: string;

  @ValidateNested()
  @Type(() => ValidationSummaryDto)
  summary!: ValidationSummaryDto;

  @IsIn(['office-js', 'vsto'])
  adapter!: 'office-js' | 'vsto';
}

export class RecordPrintRunDto {
  @IsUUID('4')
  profileId!: string;

  @IsUUID('4')
  profileVersionId!: string;

  @IsInt()
  @Min(1)
  profileVersion!: number;

  @IsString()
  @MinLength(2)
  @MaxLength(64)
  profileCode!: string;

  @IsString()
  @MaxLength(32)
  paperSize!: string;

  @IsString()
  @MaxLength(32)
  orientation!: string;

  @IsBoolean()
  applied!: boolean;

  @IsOptional()
  @IsBoolean()
  pdfExported?: boolean;

  @IsIn(['office-js', 'vsto'])
  adapter!: 'office-js' | 'vsto';
}
