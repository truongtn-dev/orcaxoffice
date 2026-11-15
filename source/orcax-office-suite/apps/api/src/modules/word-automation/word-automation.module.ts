import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuditModule } from '../audit/audit.module';
import { AuthorizationModule } from '../authorization/authorization.module';
import { LicensingModule } from '../licensing/licensing.module';
import { ProfileModule } from '../profile/profile.module';
import { ValidationRun } from './validation-run.entity';
import { WordAutomationController } from './word-automation.controller';
import { WordAutomationService } from './word-automation.service';

/**
 * Word profile and validation contracts.
 * Office Object Model calls stay in the VSTO or Office.js adapter.
 */
@Module({
  imports: [
    TypeOrmModule.forFeature([ValidationRun]),
    AuthorizationModule,
    AuditModule,
    ProfileModule,
    LicensingModule,
  ],
  controllers: [WordAutomationController],
  providers: [WordAutomationService],
  exports: [WordAutomationService],
})
export class WordAutomationModule {}
