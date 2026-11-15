import { Body, Controller, Get, HttpCode, Param, ParseUUIDPipe, Post } from '@nestjs/common';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { AcceptInvitationDto, InviteDto, RoleChangeDto } from './membership.dto';
import { MembershipService } from './membership.service';

@Controller()
export class MembershipController {
  constructor(private readonly memberships: MembershipService) {}

  @Get('members')
  async list(@CurrentAccess() access: AccessContext) {
    return { members: await this.memberships.listMembers(access) };
  }

  @Post('invitations')
  @HttpCode(201)
  invite(@CurrentAccess() access: AccessContext, @Body() body: InviteDto) {
    return this.memberships.invite(access, body.email, body.role);
  }

  @Post('invitations/accept')
  @HttpCode(200)
  accept(@CurrentAccess() access: AccessContext, @Body() body: AcceptInvitationDto) {
    return this.memberships.accept(access, body.token);
  }

  @Post('memberships/:id/suspend')
  @HttpCode(200)
  async suspend(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    await this.memberships.suspend(access, id);
    return { suspended: true };
  }

  @Post('memberships/:id/remove')
  @HttpCode(200)
  async remove(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    await this.memberships.remove(access, id);
    return { removed: true };
  }

  @Post('memberships/:id/role')
  @HttpCode(200)
  async changeRole(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: RoleChangeDto,
  ) {
    await this.memberships.changeRole(access, id, body.role);
    return { updated: true };
  }
}
