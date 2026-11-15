import {
  CanActivate,
  ExecutionContext,
  HttpStatus,
  Inject,
  Injectable,
  forwardRef,
} from '@nestjs/common';
import { ErrorCodes } from '@orcax/api-contracts';
import { Request } from 'express';
import { ApiError } from '../../common/http/api-error';
import { AccessContext } from '../../tenancy/access-context';
import { MembershipService } from '../organization/membership.service';
import { SessionService } from './session.service';
import { TokenService } from './token.service';

const PUBLIC_ROUTES = new Set([
  'POST /api/v1/auth/register',
  'POST /api/v1/auth/login',
  'POST /api/v1/auth/refresh',
  'POST /api/v1/auth/email/verify',
  'POST /api/v1/auth/email/resend',
  'POST /api/v1/auth/password/forgot',
  'POST /api/v1/auth/password/reset',
  'GET /api/v1/dev/mailbox',
  'POST /api/v1/payments/webhooks/mock',
]);

@Injectable()
export class AccessGuard implements CanActivate {
  constructor(
    private readonly tokens: TokenService,
    private readonly sessions: SessionService,
    @Inject(forwardRef(() => MembershipService))
    private readonly memberships: MembershipService,
  ) {}

  async canActivate(context: ExecutionContext): Promise<boolean> {
    const request = context
      .switchToHttp()
      .getRequest<Request & { id?: string; access?: AccessContext }>();
    if (request.method === 'OPTIONS') {
      return true;
    }
    const path = routePath(request);
    if (
      path === '/health' ||
      path === '/health/ready' ||
      PUBLIC_ROUTES.has(`${request.method} ${path}`)
    ) {
      return true;
    }
    const header = request.header('authorization') ?? '';
    const token = header.startsWith('Bearer ') ? header.slice('Bearer '.length).trim() : '';
    const claims = token ? this.tokens.verify(token) : null;
    if (!claims) {
      throw new ApiError(HttpStatus.UNAUTHORIZED, ErrorCodes.UNAUTHORIZED, 'Bạn cần đăng nhập.');
    }
    const live = await this.sessions.requireLive(claims.sid, claims.sub);
    if (live.session.activeWorkspaceId !== claims.wid) {
      throw new ApiError(
        HttpStatus.UNAUTHORIZED,
        ErrorCodes.UNAUTHORIZED,
        'Phiên đăng nhập không còn hiệu lực.',
      );
    }
    if (live.user.status === 'suspended') {
      if (!isLogout(request.method, path)) {
        throw new ApiError(HttpStatus.FORBIDDEN, ErrorCodes.FORBIDDEN, 'Tài khoản đã bị khóa.');
      }
    }
    const membership = await this.memberships.describe(claims.sub, claims.wid);
    if (!membership) {
      throw new ApiError(
        HttpStatus.UNAUTHORIZED,
        ErrorCodes.UNAUTHORIZED,
        'Phiên đăng nhập không còn hiệu lực.',
      );
    }
    const access: AccessContext = {
      userId: live.user.id,
      sessionId: live.session.id,
      workspaceId: membership.workspaceId,
      workspaceKind: membership.workspaceKind,
      workspaceName: membership.workspaceName,
      workspaceStatus: membership.workspaceStatus,
      membershipStatus: membership.membershipStatus,
      roleCode: membership.roleCode,
      organizationId: membership.organizationId,
      correlationId: request.id ?? claims.sid,
      email: live.user.email,
    };
    request.access = access;
    const blocked =
      live.user.status !== 'suspended' &&
      (membership.membershipStatus !== 'active' || membership.workspaceStatus !== 'active');
    if (blocked && !isRecovery(request.method, path)) {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Không gian làm việc hiện tại không còn hoạt động. Hãy chuyển sang không gian khác hoặc đăng xuất.',
      );
    }
    return true;
  }
}

function routePath(request: Request): string {
  const raw = request.originalUrl || request.url || '';
  return raw.split('?')[0];
}

function isLogout(method: string, path: string): boolean {
  return method === 'POST' && path === '/api/v1/auth/logout';
}

function isRecovery(method: string, path: string): boolean {
  if (
    method === 'GET' &&
    (path === '/api/v1/auth/me' || path === '/api/v1/workspaces' || path === '/api/v1/sessions')
  ) {
    return true;
  }
  if (
    method === 'POST' &&
    (path === '/api/v1/auth/logout' || path === '/api/v1/workspaces/switch')
  ) {
    return true;
  }
  return method === 'DELETE' && /^\/api\/v1\/sessions\/[0-9a-f-]{36}$/i.test(path);
}
