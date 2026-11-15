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
import { DistributeProfileDto } from '../organization/department.dto';
import {
  AddLegalReferenceDto,
  CreateProfileDto,
  EvaluateProfileDto,
  ImportProfileDto,
  ProfileQueryDto,
  ReplaceProfileRulesDto,
  TransitionNoteDto,
  VerifyLegalReferenceDto,
} from './profile.dto';
import { ProfileService } from './profile.service';

@Controller()
export class ProfileController {
  constructor(private readonly profiles: ProfileService) {}

  @Get('profiles')
  @RequirePermission('profile.read')
  list(@CurrentAccess() access: AccessContext, @Query() query: ProfileQueryDto) {
    return this.profiles.list(access, query);
  }

  @Get('profiles/:id')
  @RequirePermission('profile.read')
  get(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.profiles.get(access, id);
  }

  @Post('profiles')
  @HttpCode(201)
  @RequirePermission('profile.write')
  create(@CurrentAccess() access: AccessContext, @Body() body: CreateProfileDto) {
    return this.profiles.create(access, body);
  }

  @Post('profiles/import')
  @HttpCode(201)
  @RequirePermission('profile.write')
  importPackage(@CurrentAccess() access: AccessContext, @Body() body: ImportProfileDto) {
    return this.profiles.importPackage(access, body.package);
  }

  @Put('profiles/:id/rules')
  @RequirePermission('profile.write')
  replaceRules(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: ReplaceProfileRulesDto,
  ) {
    return this.profiles.replaceRules(access, id, body.rules);
  }

  @Post('profiles/:id/legal-references')
  @HttpCode(201)
  @RequirePermission('profile.write')
  addLegal(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: AddLegalReferenceDto,
  ) {
    return this.profiles.addLegalReference(access, id, body);
  }

  @Post('legal-references/:id/verify')
  @HttpCode(200)
  @RequirePermission('profile.review')
  verify(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: VerifyLegalReferenceDto,
  ) {
    return this.profiles.verifyLegalReference(access, id, body);
  }

  @Post('profiles/:id/submit-review')
  @HttpCode(200)
  @RequirePermission('profile.write')
  submit(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: TransitionNoteDto,
  ) {
    return this.profiles.submitReview(access, id, body.changeSummary);
  }

  @Post('profiles/:id/approve')
  @HttpCode(200)
  @RequirePermission('profile.publish')
  approve(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: TransitionNoteDto,
  ) {
    return this.profiles.approve(access, id, body.changeSummary);
  }

  @Post('profiles/:id/deprecate')
  @HttpCode(200)
  @RequirePermission('profile.publish')
  deprecate(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: TransitionNoteDto,
  ) {
    return this.profiles.deprecate(access, id, body.changeSummary);
  }

  @Post('profiles/:id/versions')
  @HttpCode(201)
  @RequirePermission('profile.write')
  newVersion(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: TransitionNoteDto,
  ) {
    return this.profiles.newDraftVersion(access, id, body.changeSummary);
  }

  @Get('profiles/:id/export')
  @RequirePermission('profile.read')
  export(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.profiles.exportPackage(access, id);
  }

  @Post('profiles/:id/evaluate')
  @HttpCode(200)
  @RequirePermission('profile.read')
  evaluate(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: EvaluateProfileDto,
  ) {
    return this.profiles.evaluate(access, id, body);
  }

  @Post('profiles/:id/distribute')
  @HttpCode(201)
  @RequirePermission('profile.publish')
  distribute(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: DistributeProfileDto,
  ) {
    return this.profiles.distribute(access, id, body);
  }

  @Get('profiles/:id/distributions')
  @RequirePermission('profile.read')
  distributions(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.profiles.listDistributions(access, id);
  }
}
