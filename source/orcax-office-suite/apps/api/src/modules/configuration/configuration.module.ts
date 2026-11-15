import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuditModule } from '../audit/audit.module';
import { AuthorizationModule } from '../authorization/authorization.module';
import { License } from '../licensing/license.entity';
import { Membership } from '../organization/membership.entity';
import { Profile } from '../profile/profile.entity';
import { Template } from '../template/template.entity';
import { AdminOverviewService } from './admin-overview.service';
import { ConfigurationController } from './configuration.controller';
import { ConfigurationService } from './configuration.service';
import { WorkspaceSettingHistory } from './workspace-setting-history.entity';
import { WorkspaceSetting } from './workspace-setting.entity';

/** Workspace configuration records with version history and admin overview. */
@Module({
  imports: [
    TypeOrmModule.forFeature([
      WorkspaceSetting,
      WorkspaceSettingHistory,
      Membership,
      License,
      Profile,
      Template,
    ]),
    AuthorizationModule,
    AuditModule,
  ],
  controllers: [ConfigurationController],
  providers: [ConfigurationService, AdminOverviewService],
  exports: [ConfigurationService, AdminOverviewService],
})
export class ConfigurationModule {}
