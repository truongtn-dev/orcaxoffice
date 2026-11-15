import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuditModule } from '../audit/audit.module';
import { AuthorizationModule } from '../authorization/authorization.module';
import { LicenseActivation } from './license-activation.entity';
import { License } from './license.entity';
import { LicensingController } from './licensing.controller';
import { LicensingService } from './licensing.service';
import { PlanEntitlement } from './plan-entitlement.entity';
import { Plan } from './plan.entity';
import { Product } from './product.entity';

@Module({
  imports: [
    TypeOrmModule.forFeature([Product, Plan, PlanEntitlement, License, LicenseActivation]),
    AuthorizationModule,
    AuditModule,
  ],
  controllers: [LicensingController],
  providers: [LicensingService],
  exports: [LicensingService],
})
export class LicensingModule {}
