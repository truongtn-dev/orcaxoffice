import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuditModule } from '../audit/audit.module';
import { AuthorizationModule } from '../authorization/authorization.module';
import { OrganizationModule } from '../organization/organization.module';
import { LegalReference } from './legal-reference.entity';
import { ProfileController } from './profile.controller';
import { ProfileDistribution } from './profile-distribution.entity';
import { ProfileRule } from './profile-rule.entity';
import { ProfileVersion } from './profile-version.entity';
import { Profile } from './profile.entity';
import { ProfileService } from './profile.service';
import { VerificationEvidence } from './verification-evidence.entity';

/** Governed profiles, versions, legal references, and rule evaluation contracts. */
@Module({
  imports: [
    TypeOrmModule.forFeature([
      Profile,
      ProfileVersion,
      ProfileRule,
      LegalReference,
      VerificationEvidence,
      ProfileDistribution,
    ]),
    AuthorizationModule,
    AuditModule,
    OrganizationModule,
  ],
  controllers: [ProfileController],
  providers: [ProfileService],
  exports: [ProfileService],
})
export class ProfileModule {}
