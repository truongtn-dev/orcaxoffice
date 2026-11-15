import { CanActivate, ExecutionContext, HttpStatus, Injectable } from '@nestjs/common';
import { Reflector } from '@nestjs/core';
import { ErrorCodes } from '@orcax/api-contracts';
import { Request } from 'express';
import { ApiError } from '../../common/http/api-error';
import { AccessContext } from '../../tenancy/access-context';
import { AuthorizationService } from './authorization.service';
import { PERMISSION_KEY, RequirePermissionOptions } from './require-permission.decorator';

@Injectable()
export class PermissionsGuard implements CanActivate {
  constructor(
    private readonly reflector: Reflector,
    private readonly authorization: AuthorizationService,
  ) {}

  canActivate(context: ExecutionContext): boolean {
    const requirement = this.reflector.getAllAndOverride<RequirePermissionOptions | undefined>(
      PERMISSION_KEY,
      [context.getHandler(), context.getClass()],
    );
    if (!requirement) {
      return true;
    }
    const request = context.switchToHttp().getRequest<Request & { access?: AccessContext }>();
    if (!request.access) {
      throw new ApiError(HttpStatus.UNAUTHORIZED, ErrorCodes.UNAUTHORIZED, 'Bạn cần đăng nhập.');
    }
    this.authorization.assertPolicy(request.access, requirement);
    return true;
  }
}
