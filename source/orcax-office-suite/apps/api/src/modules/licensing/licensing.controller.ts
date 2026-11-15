import {
  Body,
  Controller,
  Delete,
  Get,
  HttpCode,
  Param,
  ParseUUIDPipe,
  Post,
  Query,
} from '@nestjs/common';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { RequirePermission } from '../authorization/require-permission.decorator';
import { ActivateLicenseDto, EntitlementCheckDto, IssueLicenseDto } from './licensing.dto';
import { LicensingService } from './licensing.service';

@Controller()
export class LicensingController {
  constructor(private readonly licensing: LicensingService) {}

  @Get('plans')
  @RequirePermission('license.read')
  listPlans(@CurrentAccess() access: AccessContext) {
    return this.licensing.listPlans(access);
  }

  @Get('licenses')
  @RequirePermission('license.read')
  listLicenses(@CurrentAccess() access: AccessContext) {
    return this.licensing.listLicenses(access);
  }

  @Post('licenses')
  @HttpCode(201)
  @RequirePermission('license.issue')
  issue(@CurrentAccess() access: AccessContext, @Body() body: IssueLicenseDto) {
    return this.licensing.issue(access, body.planId);
  }

  @Post('licenses/:id/revoke')
  @HttpCode(200)
  @RequirePermission('license.manage')
  revoke(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.licensing.revoke(access, id);
  }

  @Post('licenses/:id/renew')
  @HttpCode(200)
  @RequirePermission('license.manage')
  renew(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.licensing.renew(access, id);
  }

  @Post('licenses/activate')
  @HttpCode(200)
  @RequirePermission('license.activate')
  activate(@CurrentAccess() access: AccessContext, @Body() body: ActivateLicenseDto) {
    return this.licensing.activate(access, body);
  }

  @Get('license-activations')
  @RequirePermission('license.read')
  listActivations(@CurrentAccess() access: AccessContext) {
    return this.licensing.listActivations(access);
  }

  @Delete('license-activations/:id')
  revokeActivation(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.licensing.revokeActivation(access, id);
  }

  @Get('entitlements/check')
  @RequirePermission('license.read')
  check(@CurrentAccess() access: AccessContext, @Query() query: EntitlementCheckDto) {
    return this.licensing.checkEntitlement(access, query.code);
  }
}
