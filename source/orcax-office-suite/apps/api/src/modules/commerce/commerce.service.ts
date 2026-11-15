import { HttpStatus, Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { InjectRepository } from '@nestjs/typeorm';
import { ErrorCodes } from '@orcax/api-contracts';
import { createHash, randomUUID } from 'crypto';
import { Repository } from 'typeorm';
import { ApiError } from '../../common/http/api-error';
import { AppConfig } from '../../config/validate-env';
import { AccessContext } from '../../tenancy/access-context';
import { requireWorkspaceId } from '../../tenancy/workspace-scope';
import { AuditRecorder } from '../audit/audit-recorder.service';
import { AuthorizationService } from '../authorization/authorization.service';
import { LicensingService } from '../licensing/licensing.service';
import { Plan } from '../licensing/plan.entity';
import { Workspace } from '../organization/workspace.entity';
import { calcLineTotal, calcOrderTotal } from './commerce-catalog';
import { Invoice } from './invoice.entity';
import { OrderItem } from './order-item.entity';
import { Order } from './order.entity';
import { Payment } from './payment.entity';
import { PlanPrice } from './plan-price.entity';
import { ReconciliationItem } from './reconciliation-item.entity';
import { Refund } from './refund.entity';

interface BuiltOrderItem {
  planId: string;
  quantity: number;
  unitAmountMinor: number;
  lineTotalMinor: number;
  currency: string;
}

@Injectable()
export class CommerceService {
  constructor(
    @InjectRepository(PlanPrice)
    private readonly prices: Repository<PlanPrice>,
    @InjectRepository(Order)
    private readonly orders: Repository<Order>,
    @InjectRepository(OrderItem)
    private readonly items: Repository<OrderItem>,
    @InjectRepository(Payment)
    private readonly payments: Repository<Payment>,
    @InjectRepository(Refund)
    private readonly refunds: Repository<Refund>,
    @InjectRepository(Invoice)
    private readonly invoices: Repository<Invoice>,
    @InjectRepository(ReconciliationItem)
    private readonly reconciliation: Repository<ReconciliationItem>,
    @InjectRepository(Plan)
    private readonly plans: Repository<Plan>,
    @InjectRepository(Workspace)
    private readonly workspaces: Repository<Workspace>,
    private readonly licensing: LicensingService,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
    private readonly config: ConfigService<AppConfig, true>,
  ) {}

  async listPrices(actor: AccessContext) {
    this.authorization.assertPolicy(actor, { permission: 'commerce.order.read' });
    const rows = await this.prices.find({
      where: { status: 'active' },
      order: { amountMinor: 'ASC' },
    });
    const result = [];
    for (const price of rows) {
      const plan = await this.plans.findOneByOrFail({ id: price.planId });
      if (plan.workspaceKind !== actor.workspaceKind) {
        continue;
      }
      result.push({
        id: price.id,
        planId: plan.id,
        planCode: plan.code,
        planName: plan.name,
        currency: price.currency,
        amountMinor: price.amountMinor,
        billingPeriod: price.billingPeriod,
      });
    }
    return { prices: result };
  }

  async listOrders(actor: AccessContext) {
    this.authorization.assertPolicy(actor, { permission: 'commerce.order.read' });
    const workspaceId = requireWorkspaceId(actor);
    const rows = await this.orders.find({
      where: { workspaceId },
      order: { createdAt: 'DESC' },
    });
    const views = [];
    for (const order of rows) {
      views.push(await this.orderView(order));
    }
    return { orders: views };
  }

  async createOrder(actor: AccessContext, itemsInput: { planId: string; quantity: number }[]) {
    this.authorization.assertPolicy(actor, { permission: 'commerce.order.create' });
    if (!itemsInput.length) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Đơn hàng cần ít nhất một dòng.',
      );
    }
    const workspaceId = requireWorkspaceId(actor);
    const now = new Date();
    const built: BuiltOrderItem[] = [];
    for (const item of itemsInput) {
      const plan = await this.plans.findOne({ where: { id: item.planId, status: 'active' } });
      if (!plan || plan.workspaceKind !== actor.workspaceKind) {
        throw new ApiError(
          HttpStatus.BAD_REQUEST,
          ErrorCodes.BAD_REQUEST,
          'Gói không khả dụng cho không gian hiện tại.',
        );
      }
      const price = await this.prices.findOne({
        where: { planId: plan.id, status: 'active' },
      });
      if (!price) {
        throw new ApiError(HttpStatus.BAD_REQUEST, ErrorCodes.BAD_REQUEST, 'Gói chưa có giá.');
      }
      const lineTotalMinor = calcLineTotal(price.amountMinor, item.quantity);
      built.push({
        planId: plan.id,
        quantity: item.quantity,
        unitAmountMinor: price.amountMinor,
        lineTotalMinor,
        currency: price.currency,
      });
    }
    const currency = built[0].currency;
    if (built.some((item) => item.currency !== currency)) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Đơn hàng chỉ một loại tiền.',
      );
    }
    const totalMinor = calcOrderTotal(built);
    const orderId = randomUUID();
    await this.orders.manager.transaction(async (manager) => {
      await manager.insert(Order, {
        id: orderId,
        workspaceId,
        status: 'draft',
        currency,
        totalMinor,
        createdBy: actor.userId,
        finalizedAt: null,
        createdAt: now,
        updatedAt: now,
      });
      for (const item of built) {
        await manager.insert(OrderItem, {
          id: randomUUID(),
          orderId,
          planId: item.planId,
          quantity: item.quantity,
          unitAmountMinor: item.unitAmountMinor,
          lineTotalMinor: item.lineTotalMinor,
        });
      }
    });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'commerce.order.created',
      targetType: 'order',
      targetId: orderId,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { totalMinor, currency },
    });
    return this.orderView(await this.orders.findOneByOrFail({ id: orderId }));
  }

  async finalize(actor: AccessContext, orderId: string) {
    this.authorization.assertPolicy(actor, { permission: 'commerce.order.create' });
    const order = await this.requireWorkspaceOrder(actor, orderId);
    if (order.status !== 'draft') {
      throw new ApiError(HttpStatus.CONFLICT, ErrorCodes.CONFLICT, 'Chỉ chốt đơn nháp.');
    }
    const now = new Date();
    await this.orders.update(
      { id: order.id },
      { status: 'pending_payment', finalizedAt: now, updatedAt: now },
    );
    await this.recorder.append({
      workspaceId: order.workspaceId,
      actorId: actor.userId,
      action: 'commerce.order.finalized',
      targetType: 'order',
      targetId: order.id,
      outcome: 'success',
      correlationId: actor.correlationId,
    });
    return this.orderView(await this.orders.findOneByOrFail({ id: order.id }));
  }

  async checkout(actor: AccessContext, orderId: string) {
    this.authorization.assertPolicy(actor, { permission: 'commerce.order.create' });
    const order = await this.requireWorkspaceOrder(actor, orderId);
    if (order.status !== 'pending_payment' && order.status !== 'draft') {
      throw new ApiError(HttpStatus.CONFLICT, ErrorCodes.CONFLICT, 'Đơn không thể thanh toán.');
    }
    if (order.status === 'draft') {
      await this.finalize(actor, orderId);
    }
    const fresh = await this.orders.findOneByOrFail({ id: orderId });
    const now = new Date();
    const paymentId = randomUUID();
    const idempotencyKey = `checkout:${paymentId}`;
    await this.payments.insert({
      id: paymentId,
      orderId: fresh.id,
      workspaceId: fresh.workspaceId,
      provider: 'mock',
      providerRef: null,
      idempotencyKey,
      status: 'pending',
      amountMinor: fresh.totalMinor,
      createdAt: now,
      updatedAt: now,
    });
    if (fresh.totalMinor === 0) {
      await this.applyPaymentSuccess({
        paymentId,
        providerRef: 'free-zero',
        idempotencyKey,
        correlationId: actor.correlationId,
      });
    }
    return {
      paymentId,
      orderId: fresh.id,
      amountMinor: fresh.totalMinor,
      currency: fresh.currency,
      provider: 'mock',
      idempotencyKey,
      status: fresh.totalMinor === 0 ? 'succeeded' : 'pending',
    };
  }

  async handleMockWebhook(input: {
    secret: string;
    idempotencyKey: string;
    paymentId: string;
    status: 'succeeded' | 'failed';
    providerRef?: string;
    correlationId: string;
  }) {
    const expected = this.config.get('paymentWebhookSecret', { infer: true });
    if (!expected || input.secret !== expected) {
      throw new ApiError(HttpStatus.UNAUTHORIZED, ErrorCodes.UNAUTHORIZED, 'Webhook không hợp lệ.');
    }
    const existing = await this.payments.findOne({
      where: { idempotencyKey: input.idempotencyKey },
    });
    if (existing && existing.status !== 'pending') {
      return { replayed: true, paymentId: existing.id, status: existing.status };
    }
    const payment =
      existing ??
      (await this.payments.findOne({
        where: { id: input.paymentId },
      }));
    if (!payment) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy thanh toán.');
    }
    if (payment.idempotencyKey !== input.idempotencyKey) {
      await this.payments.update({ id: payment.id }, { idempotencyKey: input.idempotencyKey });
    }
    if (input.status === 'failed') {
      await this.payments.update(
        { id: payment.id },
        { status: 'failed', providerRef: input.providerRef ?? null, updatedAt: new Date() },
      );
      return { replayed: false, paymentId: payment.id, status: 'failed' };
    }
    await this.applyPaymentSuccess({
      paymentId: payment.id,
      providerRef:
        input.providerRef ??
        `mock-${createHash('sha256').update(input.idempotencyKey).digest('hex').slice(0, 12)}`,
      idempotencyKey: input.idempotencyKey,
      correlationId: input.correlationId,
    });
    return { replayed: false, paymentId: payment.id, status: 'succeeded' };
  }

  async createRefund(
    actor: AccessContext,
    input: { paymentId: string; amountMinor: number; reason: string },
  ) {
    this.authorization.assertPolicy(actor, { permission: 'commerce.refund' });
    const workspaceId = requireWorkspaceId(actor);
    const payment = await this.payments.findOne({
      where: { id: input.paymentId, workspaceId },
    });
    if (!payment || payment.status !== 'succeeded') {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Chỉ hoàn tiền trên thanh toán thành công.',
      );
    }
    if (input.amountMinor > payment.amountMinor) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Số hoàn vượt thanh toán.',
      );
    }
    const id = randomUUID();
    const now = new Date();
    await this.refunds.insert({
      id,
      paymentId: payment.id,
      orderId: payment.orderId,
      workspaceId,
      amountMinor: input.amountMinor,
      status: 'succeeded',
      providerRef: `refund-${id.slice(0, 8)}`,
      reason: input.reason,
      createdBy: actor.userId,
      createdAt: now,
    });
    const order = await this.orders.findOneByOrFail({ id: payment.orderId });
    const nextStatus = input.amountMinor >= order.totalMinor ? 'refunded' : 'partially_refunded';
    await this.orders.update({ id: order.id }, { status: nextStatus, updatedAt: now });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'commerce.refund.created',
      targetType: 'refund',
      targetId: id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { amountMinor: input.amountMinor, orderId: order.id },
    });
    return { id, status: 'succeeded', amountMinor: input.amountMinor };
  }

  async listReconciliation(actor: AccessContext) {
    this.authorization.assertPolicy(actor, { permission: 'commerce.reconcile' });
    const workspaceId = requireWorkspaceId(actor);
    const rows = await this.reconciliation.find({
      where: { workspaceId },
      order: { createdAt: 'DESC' },
    });
    return {
      items: rows.map((row) => ({
        id: row.id,
        paymentId: row.paymentId,
        status: row.status,
        note: row.note,
        createdAt: row.createdAt.toISOString(),
      })),
    };
  }

  async resolveReconciliation(
    actor: AccessContext,
    id: string,
    status: 'matched' | 'ignored',
    note?: string,
  ) {
    this.authorization.assertPolicy(actor, { permission: 'commerce.reconcile' });
    const workspaceId = requireWorkspaceId(actor);
    const item = await this.reconciliation.findOne({ where: { id, workspaceId } });
    if (!item) {
      throw new ApiError(
        HttpStatus.NOT_FOUND,
        ErrorCodes.NOT_FOUND,
        'Không tìm thấy mục đối soát.',
      );
    }
    await this.reconciliation.update(
      { id },
      { status, note: note ?? item.note, updatedAt: new Date() },
    );
    return { id, status };
  }

  private async applyPaymentSuccess(input: {
    paymentId: string;
    providerRef: string;
    idempotencyKey: string;
    correlationId: string;
  }) {
    const payment = await this.payments.findOneByOrFail({ id: input.paymentId });
    if (payment.status === 'succeeded') {
      return;
    }
    const order = await this.orders.findOneByOrFail({ id: payment.orderId });
    const now = new Date();
    if (
      order.status === 'paid' ||
      order.status === 'refunded' ||
      order.status === 'partially_refunded'
    ) {
      await this.payments.update(
        { id: payment.id },
        {
          status: 'succeeded',
          providerRef: input.providerRef,
          idempotencyKey: input.idempotencyKey,
          updatedAt: now,
        },
      );
      await this.reconciliation.insert({
        id: randomUUID(),
        workspaceId: order.workspaceId,
        paymentId: payment.id,
        status: 'unmatched',
        note: `Thanh toán trùng đơn đã ${order.status}: ${payment.id}`,
        createdAt: now,
        updatedAt: now,
      });
      return;
    }
    const workspace = await this.workspaces.findOneByOrFail({ id: order.workspaceId });
    const items = await this.items.find({ where: { orderId: order.id } });
    await this.orders.manager.transaction(async (manager) => {
      await manager.update(
        Payment,
        { id: payment.id },
        {
          status: 'succeeded',
          providerRef: input.providerRef,
          idempotencyKey: input.idempotencyKey,
          updatedAt: now,
        },
      );
      await manager.update(Order, { id: order.id }, { status: 'paid', updatedAt: now });
      const invoiceNumber = `INV-${now.getUTCFullYear()}-${order.id.slice(0, 8).toUpperCase()}`;
      await manager.insert(Invoice, {
        id: randomUUID(),
        orderId: order.id,
        workspaceId: order.workspaceId,
        invoiceNumber,
        status: 'issued',
        amountMinor: order.totalMinor,
        issuedAt: now,
      });
    });
    for (const item of items) {
      await this.licensing.issueFromPaidOrder({
        workspaceId: order.workspaceId,
        workspaceKind: workspace.kind,
        planId: item.planId,
        issuedBy: order.createdBy,
        correlationId: input.correlationId,
        orderId: order.id,
      });
    }
    await this.recorder.append({
      workspaceId: order.workspaceId,
      actorId: order.createdBy,
      action: 'commerce.payment.succeeded',
      targetType: 'payment',
      targetId: payment.id,
      outcome: 'success',
      correlationId: input.correlationId,
      metadata: { orderId: order.id, amountMinor: payment.amountMinor },
    });
  }

  private async requireWorkspaceOrder(actor: AccessContext, orderId: string) {
    const workspaceId = requireWorkspaceId(actor);
    const order = await this.orders.findOne({ where: { id: orderId, workspaceId } });
    if (!order) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy đơn hàng.');
    }
    return order;
  }

  private async orderView(order: Order) {
    const items = await this.items.find({ where: { orderId: order.id } });
    return {
      id: order.id,
      status: order.status,
      currency: order.currency,
      totalMinor: order.totalMinor,
      finalizedAt: order.finalizedAt ? order.finalizedAt.toISOString() : null,
      createdAt: order.createdAt.toISOString(),
      items: items.map((item) => ({
        id: item.id,
        planId: item.planId,
        quantity: item.quantity,
        unitAmountMinor: item.unitAmountMinor,
        lineTotalMinor: item.lineTotalMinor,
      })),
    };
  }
}
