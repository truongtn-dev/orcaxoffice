import { Type } from 'class-transformer';
import {
  ArrayMinSize,
  IsArray,
  IsIn,
  IsObject,
  IsOptional,
  IsString,
  IsUUID,
  MaxLength,
  MinLength,
  ValidateNested,
} from 'class-validator';

export class CreateProfileDto {
  @IsString({ message: 'Mã hồ sơ không hợp lệ.' })
  @MinLength(2, { message: 'Mã hồ sơ quá ngắn.' })
  @MaxLength(64, { message: 'Mã hồ sơ tối đa 64 ký tự.' })
  code!: string;

  @IsString({ message: 'Tên hồ sơ không hợp lệ.' })
  @MinLength(2, { message: 'Tên hồ sơ quá ngắn.' })
  @MaxLength(160, { message: 'Tên hồ sơ tối đa 160 ký tự.' })
  name!: string;

  @IsIn(['word', 'excel'], { message: 'Host phải là word hoặc excel.' })
  host!: 'word' | 'excel';

  @IsString({ message: 'Nguồn không hợp lệ.' })
  @MinLength(2, { message: 'Nguồn không được để trống.' })
  @MaxLength(2000, { message: 'Nguồn quá dài.' })
  source!: string;
}

export class ProfileRuleDto {
  @IsString()
  @MinLength(2)
  @MaxLength(64)
  code!: string;

  @IsString()
  @MinLength(1)
  @MaxLength(32)
  version!: string;

  @IsString()
  @MinLength(1)
  @MaxLength(200)
  title!: string;

  @IsString()
  @MaxLength(2000)
  description!: string;

  @IsString()
  @MinLength(1)
  @MaxLength(64)
  checkKind!: string;

  @IsIn(['fail', 'warning'])
  defaultSeverity!: 'fail' | 'warning';

  @IsOptional()
  @IsObject()
  checkConfig?: Record<string, unknown>;
}

export class ReplaceProfileRulesDto {
  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => ProfileRuleDto)
  rules!: ProfileRuleDto[];
}

export class AddLegalReferenceDto {
  @IsString()
  @MinLength(2)
  @MaxLength(500)
  citation!: string;

  @IsOptional()
  @IsString()
  @MaxLength(1000)
  sourceUrl?: string;

  @IsOptional()
  @IsString()
  @MaxLength(2000)
  notes?: string;
}

export class VerifyLegalReferenceDto {
  @IsIn(['document', 'url', 'statement'])
  evidenceType!: 'document' | 'url' | 'statement';

  @IsString()
  @MinLength(2)
  @MaxLength(2000)
  summary!: string;

  @IsOptional()
  @IsString()
  @MaxLength(1000)
  referenceUri?: string;
}

export class ImportProfileDto {
  @IsObject()
  package!: Record<string, unknown>;
}

export class EvaluateProfileDto {
  @IsArray()
  @ArrayMinSize(0)
  results!: Array<Record<string, unknown>>;
}

export class ProfileQueryDto {
  @IsOptional()
  @IsIn(['word', 'excel'])
  host?: 'word' | 'excel';

  @IsOptional()
  @IsIn(['draft', 'in_review', 'approved', 'deprecated'])
  status?: 'draft' | 'in_review' | 'approved' | 'deprecated';
}

export class TransitionNoteDto {
  @IsOptional()
  @IsString()
  @MaxLength(2000)
  changeSummary?: string;
}

export class CreateTemplateDto {
  @IsString()
  @MinLength(2)
  @MaxLength(64)
  code!: string;

  @IsString()
  @MinLength(2)
  @MaxLength(160)
  name!: string;

  @IsIn(['word', 'excel'])
  host!: 'word' | 'excel';
}

export class FieldDefinitionDto {
  @IsString()
  @MinLength(1)
  @MaxLength(64)
  fieldKey!: string;

  @IsString()
  @MinLength(1)
  @MaxLength(160)
  label!: string;

  @IsString()
  @MinLength(1)
  @MaxLength(32)
  fieldType!: string;

  @IsOptional()
  required?: boolean;

  @IsOptional()
  @IsString()
  @MaxLength(2000)
  defaultValue?: string;

  @IsOptional()
  @IsObject()
  validationJson?: Record<string, unknown>;
}

export class ReplaceFieldsDto {
  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => FieldDefinitionDto)
  fields!: FieldDefinitionDto[];
}

export class DataMappingDto {
  @IsString()
  @MinLength(1)
  @MaxLength(64)
  fieldKey!: string;

  @IsIn(['form', 'entity', 'constant'])
  sourceKind!: 'form' | 'entity' | 'constant';

  @IsString()
  @MinLength(1)
  @MaxLength(200)
  sourcePath!: string;

  @IsIn(['none', 'trim', 'upper', 'lower', 'date_iso'])
  transform!: 'none' | 'trim' | 'upper' | 'lower' | 'date_iso';
}

export class ReplaceMappingsDto {
  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => DataMappingDto)
  mappings!: DataMappingDto[];
}

export class TemplateQueryDto {
  @IsOptional()
  @IsIn(['word', 'excel'])
  host?: 'word' | 'excel';

  @IsOptional()
  @IsUUID('4')
  ignore?: string;
}
