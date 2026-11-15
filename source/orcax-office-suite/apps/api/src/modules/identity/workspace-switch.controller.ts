import { Body, Controller, HttpCode, Post } from '@nestjs/common';
import { CorrelationId, CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { SwitchWorkspaceDto } from './auth.dto';
import { AuthService } from './auth.service';
import { SessionService } from './session.service';
import { TokenService } from './token.service';

@Controller('workspaces')
export class WorkspaceSwitchController {
  constructor(
    private readonly sessions: SessionService,
    private readonly auth: AuthService,
    private readonly tokens: TokenService,
  ) {}

  @Post('switch')
  @HttpCode(200)
  async switchWorkspace(
    @CurrentAccess() access: AccessContext,
    @Body() body: SwitchWorkspaceDto,
    @CorrelationId() correlationId: string,
  ) {
    const switched = await this.sessions.switchWorkspace({
      userId: access.userId,
      sessionId: access.sessionId,
      workspaceId: body.workspaceId,
      confirmed: body.confirmed,
      correlationId,
    });
    const view = await this.auth.presentAccess(access.userId, switched.workspaceId);
    const accessToken = this.tokens.sign(access.userId, access.sessionId, switched.workspaceId);
    return {
      accessToken: accessToken.token,
      expiresIn: accessToken.expiresIn,
      ...view,
    };
  }
}
