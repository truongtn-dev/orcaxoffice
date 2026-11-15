import { Module, forwardRef } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuditModule } from '../audit/audit.module';
import { AuthorizationModule } from '../authorization/authorization.module';
import { IdentityModule } from '../identity/identity.module';
import { DepartmentController } from './department.controller';
import { DepartmentMember } from './department-member.entity';
import { Department } from './department.entity';
import { DepartmentService } from './department.service';
import { GroupMember } from './group-member.entity';
import { Invitation } from './invitation.entity';
import { MembershipController } from './membership.controller';
import { Membership } from './membership.entity';
import { MembershipService } from './membership.service';
import { OrgGroup } from './org-group.entity';
import { OrganizationController } from './organization.controller';
import { Organization } from './organization.entity';
import { WorkspaceController } from './workspace.controller';
import { Workspace } from './workspace.entity';
import { WorkspaceService } from './workspace.service';

@Module({
  imports: [
    TypeOrmModule.forFeature([
      Organization,
      Workspace,
      Membership,
      Invitation,
      Department,
      OrgGroup,
      DepartmentMember,
      GroupMember,
    ]),
    AuthorizationModule,
    AuditModule,
    forwardRef(() => IdentityModule),
  ],
  controllers: [
    OrganizationController,
    WorkspaceController,
    MembershipController,
    DepartmentController,
  ],
  providers: [WorkspaceService, MembershipService, DepartmentService],
  exports: [WorkspaceService, MembershipService, DepartmentService],
})
export class OrganizationModule {}
