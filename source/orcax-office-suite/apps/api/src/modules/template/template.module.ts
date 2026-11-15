import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuditModule } from '../audit/audit.module';
import { AuthorizationModule } from '../authorization/authorization.module';
import { DataMapping } from './data-mapping.entity';
import { FieldDefinition } from './field-definition.entity';
import { TemplateController } from './template.controller';
import { TemplateVersion } from './template-version.entity';
import { Template } from './template.entity';
import { TemplateService } from './template.service';

/** Template lifecycle and deterministic field mapping. */
@Module({
  imports: [
    TypeOrmModule.forFeature([Template, TemplateVersion, FieldDefinition, DataMapping]),
    AuthorizationModule,
    AuditModule,
  ],
  controllers: [TemplateController],
  providers: [TemplateService],
  exports: [TemplateService],
})
export class TemplateModule {}
