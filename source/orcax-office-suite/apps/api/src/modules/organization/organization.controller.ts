import { Body, Controller, Get, HttpCode, HttpStatus, Patch, Post } from '@nestjs/common';
import { ErrorCodes } from '@orcax/api-contracts';
import { ApiError } from '../../common/http/api-error';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { RequirePermission } from '../authorization/require-permission.decorator';
import { OrganizationStatusDto } from './department.dto';
import { CreateOrganizationDto } from './membership.dto';
import { WorkspaceService } from './workspace.service';

@Controller('organizations')
export class OrganizationController {
  constructor(private readonly workspaces: WorkspaceService) {}

  @Post()
  @HttpCode(201)
  create(@CurrentAccess() access: AccessContext, @Body() body: CreateOrganizationDto) {
    return this.workspaces.createOrganization(access, body.name);
  }

  @Get('current')
  async current(@CurrentAccess() access: AccessContext) {
    const organization = await this.workspaces.currentOrganization(access);
    if (!organization) {
      throw new ApiError(
        HttpStatus.NOT_FOUND,
        ErrorCodes.NOT_FOUND,
        'Không gian hiện tại không phải tổ chức.',
      );
    }
    return organization;
  }

  @Patch('current')
  @RequirePermission('workspace.update')
  rename(@CurrentAccess() access: AccessContext, @Body() body: CreateOrganizationDto) {
    return this.workspaces.renameCurrent(access, body.name);
  }

  @Post('current/status')
  @HttpCode(200)
  @RequirePermission('workspace.update')
  setStatus(@CurrentAccess() access: AccessContext, @Body() body: OrganizationStatusDto) {
    return this.workspaces.setOrganizationStatus(access, body.status);
  }
}
