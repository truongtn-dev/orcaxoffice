import {
  Body,
  Controller,
  Get,
  HttpCode,
  Param,
  ParseUUIDPipe,
  Post,
  Put,
  Query,
} from '@nestjs/common';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { RequirePermission } from '../authorization/require-permission.decorator';
import {
  CreateTemplateDto,
  ReplaceFieldsDto,
  ReplaceMappingsDto,
  TemplateQueryDto,
} from '../profile/profile.dto';
import { TemplateService } from './template.service';

@Controller()
export class TemplateController {
  constructor(private readonly templates: TemplateService) {}

  @Get('templates')
  @RequirePermission('template.read')
  list(@CurrentAccess() access: AccessContext, @Query() query: TemplateQueryDto) {
    return this.templates.list(access, query.host);
  }

  @Get('templates/:id')
  @RequirePermission('template.read')
  get(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.templates.get(access, id);
  }

  @Post('templates')
  @HttpCode(201)
  @RequirePermission('template.write')
  create(@CurrentAccess() access: AccessContext, @Body() body: CreateTemplateDto) {
    return this.templates.create(access, body);
  }

  @Put('templates/:id/fields')
  @RequirePermission('template.write')
  fields(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: ReplaceFieldsDto,
  ) {
    return this.templates.replaceFields(access, id, body.fields);
  }

  @Put('templates/:id/mappings')
  @RequirePermission('template.write')
  mappings(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: ReplaceMappingsDto,
  ) {
    return this.templates.replaceMappings(access, id, body.mappings);
  }

  @Get('templates/:id/mapping-preview')
  @RequirePermission('template.read')
  preview(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.templates.mappingPreview(access, id);
  }

  @Post('templates/:id/submit-review')
  @HttpCode(200)
  @RequirePermission('template.write')
  submit(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.templates.submitReview(access, id);
  }

  @Post('templates/:id/approve')
  @HttpCode(200)
  @RequirePermission('template.publish')
  approve(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.templates.approve(access, id);
  }

  @Post('templates/:id/publish')
  @HttpCode(200)
  @RequirePermission('template.publish')
  publish(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.templates.publish(access, id);
  }

  @Post('templates/:id/deprecate')
  @HttpCode(200)
  @RequirePermission('template.publish')
  deprecate(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.templates.deprecate(access, id);
  }

  @Post('templates/:id/versions')
  @HttpCode(201)
  @RequirePermission('template.write')
  newVersion(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.templates.newDraftVersion(access, id);
  }
}
