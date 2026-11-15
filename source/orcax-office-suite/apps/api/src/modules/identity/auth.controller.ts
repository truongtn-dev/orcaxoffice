import { Body, Controller, Get, HttpCode, Patch, Post, Req, Res } from '@nestjs/common';
import { Request, Response } from 'express';
import { AccessContext } from '../../tenancy/access-context';
import { CorrelationId, CurrentAccess } from '../../common/http/request-context';
import {
  EmailCodeDto,
  EmailDto,
  LoginDto,
  ProfileDto,
  RegisterDto,
  ResetPasswordDto,
} from './auth.dto';
import { AuthService } from './auth.service';
import { clearSessionCookie, readCookie, sessionCookie, SESSION_COOKIE } from './session-cookie';
import { SessionService } from './session.service';
import { TokenService } from './token.service';

@Controller('auth')
export class AuthController {
  constructor(
    private readonly auth: AuthService,
    private readonly sessions: SessionService,
    private readonly tokens: TokenService,
  ) {}

  @Post('register')
  @HttpCode(200)
  register(@Body() body: RegisterDto, @CorrelationId() correlationId: string) {
    return this.auth.register({ ...body, correlationId });
  }

  @Post('login')
  @HttpCode(200)
  async login(
    @Body() body: LoginDto,
    @CorrelationId() correlationId: string,
    @Req() request: Request,
    @Res({ passthrough: true }) response: Response,
  ) {
    const result = await this.auth.login({
      email: body.email,
      password: body.password,
      userAgent: userAgent(request),
      correlationId,
    });
    this.writeCookie(response, result.cookieValue);
    return {
      accessToken: result.accessToken,
      expiresIn: result.expiresIn,
      user: result.user,
      workspace: result.workspace,
      suggestedWorkspace: result.suggestedWorkspace,
      permissions: result.permissions,
    };
  }

  @Post('refresh')
  @HttpCode(200)
  async refresh(
    @CorrelationId() correlationId: string,
    @Req() request: Request,
    @Res({ passthrough: true }) response: Response,
  ) {
    const cookie = readCookie(request.header('cookie'), SESSION_COOKIE);
    const rotated = await this.sessions.rotate(cookie ?? '', correlationId);
    this.writeCookie(response, rotated.cookieValue);
    const access = this.tokens.sign(rotated.user.id, rotated.sessionId, rotated.workspaceId);
    const view = await this.auth.present(rotated.user, rotated.workspaceId);
    return { accessToken: access.token, expiresIn: access.expiresIn, ...view };
  }

  @Post('logout')
  @HttpCode(200)
  async logout(
    @CurrentAccess() access: AccessContext,
    @CorrelationId() correlationId: string,
    @Res({ passthrough: true }) response: Response,
  ) {
    await this.auth.logout(access.userId, access.sessionId, correlationId);
    response.append('Set-Cookie', clearSessionCookie(this.sessions.secureCookie()));
    return { loggedOut: true };
  }

  @Post('email/verify')
  @HttpCode(200)
  verify(@Body() body: EmailCodeDto, @CorrelationId() correlationId: string) {
    return this.auth.verifyEmail(body.email, body.code, correlationId);
  }

  @Post('email/resend')
  @HttpCode(200)
  resend(@Body() body: EmailDto, @CorrelationId() correlationId: string) {
    return this.auth.resend(body.email, correlationId);
  }

  @Post('password/forgot')
  @HttpCode(200)
  forgot(@Body() body: EmailDto, @CorrelationId() correlationId: string) {
    return this.auth.forgot(body.email, correlationId);
  }

  @Post('password/reset')
  @HttpCode(200)
  reset(@Body() body: ResetPasswordDto, @CorrelationId() correlationId: string) {
    return this.auth.resetPassword({ ...body, correlationId });
  }

  @Get('me')
  me(@CurrentAccess() access: AccessContext) {
    return this.auth.presentAccess(access.userId, access.workspaceId);
  }

  @Patch('me')
  update(
    @CurrentAccess() access: AccessContext,
    @Body() body: ProfileDto,
    @CorrelationId() correlationId: string,
  ) {
    return this.auth.updateProfile(
      access.userId,
      body.displayName,
      access.workspaceId,
      correlationId,
    );
  }

  private writeCookie(response: Response, value: string): void {
    response.append(
      'Set-Cookie',
      sessionCookie(value, {
        maxAgeSeconds: this.sessions.cookieMaxAgeSeconds(),
        secure: this.sessions.secureCookie(),
      }),
    );
  }
}

function userAgent(request: Request): string | null {
  const value = request.header('user-agent');
  if (!value) {
    return null;
  }
  return value.slice(0, 256);
}
