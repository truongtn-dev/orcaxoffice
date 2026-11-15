import { HttpStatus, Injectable } from '@nestjs/common';
import { ErrorCodes } from '@orcax/api-contracts';
import { ApiError } from '../../common/http/api-error';
import { AccessContext } from '../../tenancy/access-context';
import { can, canAssignRole, permissionsFor } from './catalog';
import { evaluatePolicy, PolicyRequirement } from './policy';

@Injectable()
export class AuthorizationService {
  permissions(roleCode: string): string[] {
    return [...permissionsFor(roleCode)];
  }

  assert(roleCode: string, permission: string): void {
    if (!can(roleCode, permission)) {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Bạn không có quyền thực hiện thao tác này.',
      );
    }
  }

  assertPolicy(access: AccessContext, requirement: PolicyRequirement): void {
    const decision = evaluatePolicy(access, requirement);
    if (decision === 'allow') {
      return;
    }
    throw new ApiError(
      HttpStatus.FORBIDDEN,
      ErrorCodes.FORBIDDEN,
      decision === 'deny_permission'
        ? 'Bạn không có quyền thực hiện thao tác này.'
        : 'Điều kiện không gian làm việc hoặc quyền sở hữu không cho phép thao tác này.',
    );
  }

  assertAssignable(actorRole: string, targetRole: string): void {
    if (!canAssignRole(actorRole, targetRole)) {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Bạn không thể gán vai trò này.',
      );
    }
  }
}
