import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuditModule } from '../audit/audit.module';
import { AuthorizationModule } from '../authorization/authorization.module';
import { LicensingModule } from '../licensing/licensing.module';
import { ProfileModule } from '../profile/profile.module';
import { ExcelAutomationController } from './excel-automation.controller';
import { ExcelAutomationService } from './excel-automation.service';
import { PrintRun } from './print-run.entity';

/**
 * Excel print and validation contracts.
 * Workbook changes stay in the Office adapter.
 */
@Module({
  imports: [
    TypeOrmModule.forFeature([PrintRun]),
    AuthorizationModule,
    AuditModule,
    ProfileModule,
    LicensingModule,
  ],
  controllers: [ExcelAutomationController],
  providers: [ExcelAutomationService],
  exports: [ExcelAutomationService],
})
export class ExcelAutomationModule {}
