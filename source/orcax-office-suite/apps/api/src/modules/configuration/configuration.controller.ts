import { Body, Controller, Get, HttpCode, Param, Post } from '@nestjs/common';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { RequirePermission } from '../authorization/require-permission.decorator';
import { AdminOverviewService } from './admin-overview.service';
import { UpsertSettingDto } from './configuration.dto';
import { ConfigurationService } from './configuration.service';

@Controller()
export class ConfigurationController {
  constructor(
    private readonly configuration: ConfigurationService,
    private readonly overview: AdminOverviewService,
  ) {}

  @Get('admin/overview')
  @RequirePermission('workspace.read')
  adminOverview(@CurrentAccess() access: AccessContext) {
    return this.overview.overview(access);
  }

  @Get('workspace-settings')
  @RequirePermission('workspace.config.read')
  list(@CurrentAccess() access: AccessContext) {
    return this.configuration.list(access);
  }

  @Post('workspace-settings')
  @HttpCode(200)
  @RequirePermission('workspace.config.manage')
  upsert(@CurrentAccess() access: AccessContext, @Body() body: UpsertSettingDto) {
    return this.configuration.upsert(access, body.key, body.value);
  }

  @Get('workspace-settings/:key/history')
  @RequirePermission('workspace.config.read')
  history(@CurrentAccess() access: AccessContext, @Param('key') key: string) {
    return this.configuration.historyFor(access, key);
  }
}
