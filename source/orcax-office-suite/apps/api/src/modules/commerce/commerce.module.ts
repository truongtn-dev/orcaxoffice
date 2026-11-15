import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuditModule } from '../audit/audit.module';
import { AuthorizationModule } from '../authorization/authorization.module';
import { LicensingModule } from '../licensing/licensing.module';
import { Plan } from '../licensing/plan.entity';
import { Workspace } from '../organization/workspace.entity';
import { CommerceController } from './commerce.controller';
import { CommerceService } from './commerce.service';
import { Invoice } from './invoice.entity';
import { OrderItem } from './order-item.entity';
import { Order } from './order.entity';
import { Payment } from './payment.entity';
import { PlanPrice } from './plan-price.entity';
import { ReconciliationItem } from './reconciliation-item.entity';
import { Refund } from './refund.entity';

/** Orders, payments, refunds, invoices, and reconciliation. */
@Module({
  imports: [
    TypeOrmModule.forFeature([
      PlanPrice,
      Order,
      OrderItem,
      Payment,
      Refund,
      Invoice,
      ReconciliationItem,
      Plan,
      Workspace,
    ]),
    AuthorizationModule,
    AuditModule,
    LicensingModule,
  ],
  controllers: [CommerceController],
  providers: [CommerceService],
  exports: [CommerceService],
})
export class CommerceModule {}
