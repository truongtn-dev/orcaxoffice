import { Controller, Delete, Get, Param, ParseUUIDPipe, Res } from '@nestjs/common';
import { Response } from 'express';
import { CorrelationId, CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { clearSessionCookie } from './session-cookie';
import { SessionService } from './session.service';

@Controller('sessions')
export class SessionController {
  constructor(private readonly sessions: SessionService) {}

  @Get()
  async list(@CurrentAccess() access: AccessContext) {
    return { sessions: await this.sessions.list(access.userId, access.sessionId) };
  }

  @Delete(':id')
  async revoke(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @CorrelationId() correlationId: string,
    @Res({ passthrough: true }) response: Response,
  ) {
    await this.sessions.revoke(access.userId, id, access.sessionId, correlationId);
    if (id === access.sessionId) {
      response.append('Set-Cookie', clearSessionCookie(this.sessions.secureCookie()));
    }
    return { revoked: true };
  }
}
