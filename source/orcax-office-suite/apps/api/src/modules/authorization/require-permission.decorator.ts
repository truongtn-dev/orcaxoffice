import { SetMetadata } from '@nestjs/common';
import { PermissionCode } from './catalog';
import { WorkspaceKind } from '../../tenancy/workspace-scope';

export const PERMISSION_KEY = 'orcax.permission';

export interface RequirePermissionOptions {
  permission: PermissionCode;
  workspaceKinds?: readonly WorkspaceKind[];
  ownership?: 'self' | 'any';
  requireActiveMembership?: boolean;
}

export const RequirePermission = (options: RequirePermissionOptions | PermissionCode) =>
  SetMetadata(
    PERMISSION_KEY,
    typeof options === 'string'
      ? ({ permission: options } satisfies RequirePermissionOptions)
      : options,
  );
