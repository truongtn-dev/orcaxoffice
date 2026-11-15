import { Controller, Get } from '@nestjs/common';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { MembershipService } from './membership.service';

@Controller('workspaces')
export class WorkspaceController {
  constructor(private readonly memberships: MembershipService) {}

  @Get()
  async list(@CurrentAccess() access: AccessContext) {
    return { workspaces: await this.memberships.listForUser(access.userId) };
  }
}
