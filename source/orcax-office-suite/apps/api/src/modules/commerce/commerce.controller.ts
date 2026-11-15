import {
  Body,
  Controller,
  Get,
  Headers,
  HttpCode,
  Param,
  ParseUUIDPipe,
  Post,
  Req,
} from '@nestjs/common';
import { Request } from 'express';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { RequirePermission } from '../authorization/require-permission.decorator';
import {
  CheckoutDto,
  CreateOrderDto,
  CreateRefundDto,
  MockWebhookDto,
  ReconcileDto,
} from './commerce.dto';
import { CommerceService } from './commerce.service';

@Controller()
export class CommerceController {
  constructor(private readonly commerce: CommerceService) {}

  @Get('plan-prices')
  @RequirePermission('commerce.order.read')
  prices(@CurrentAccess() access: AccessContext) {
    return this.commerce.listPrices(access);
  }

  @Get('orders')
  @RequirePermission('commerce.order.read')
  list(@CurrentAccess() access: AccessContext) {
    return this.commerce.listOrders(access);
  }

  @Post('orders')
  @HttpCode(201)
  @RequirePermission('commerce.order.create')
  create(@CurrentAccess() access: AccessContext, @Body() body: CreateOrderDto) {
    return this.commerce.createOrder(access, body.items);
  }

  @Post('orders/:id/finalize')
  @HttpCode(200)
  @RequirePermission('commerce.order.create')
  finalize(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.commerce.finalize(access, id);
  }

  @Post('payments/checkout')
  @HttpCode(200)
  @RequirePermission('commerce.order.create')
  checkout(@CurrentAccess() access: AccessContext, @Body() body: CheckoutDto) {
    return this.commerce.checkout(access, body.orderId);
  }

  @Post('payments/webhooks/mock')
  @HttpCode(200)
  webhook(
    @Headers('x-orcax-webhook-secret') secret: string | undefined,
    @Body() body: MockWebhookDto,
    @Req() request: Request & { id?: string },
  ) {
    return this.commerce.handleMockWebhook({
      secret: secret || '',
      idempotencyKey: body.idempotencyKey,
      paymentId: body.paymentId,
      status: body.status,
      providerRef: body.providerRef,
      correlationId: request.id || body.paymentId,
    });
  }

  @Post('refunds')
  @HttpCode(201)
  @RequirePermission('commerce.refund')
  refund(@CurrentAccess() access: AccessContext, @Body() body: CreateRefundDto) {
    return this.commerce.createRefund(access, body);
  }

  @Get('reconciliation-items')
  @RequirePermission('commerce.reconcile')
  reconciliation(@CurrentAccess() access: AccessContext) {
    return this.commerce.listReconciliation(access);
  }

  @Post('reconciliation-items/:id')
  @HttpCode(200)
  @RequirePermission('commerce.reconcile')
  resolve(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: ReconcileDto,
  ) {
    return this.commerce.resolveReconciliation(access, id, body.status, body.note);
  }
}
