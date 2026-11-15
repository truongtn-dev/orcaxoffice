import { Module } from '@nestjs/common';
import { APP_GUARD } from '@nestjs/core';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuthorizationService } from './authorization.service';
import { Permission } from './permission.entity';
import { PermissionsGuard } from './permissions.guard';
import { RoleCatalog } from './role-catalog';
import { RolePermission } from './role-permission.entity';
import { Role } from './role.entity';

@Module({
  imports: [TypeOrmModule.forFeature([Role, Permission, RolePermission])],
  providers: [
    RoleCatalog,
    AuthorizationService,
    PermissionsGuard,
    { provide: APP_GUARD, useClass: PermissionsGuard },
  ],
  exports: [RoleCatalog, AuthorizationService, PermissionsGuard],
})
export class AuthorizationModule {}
