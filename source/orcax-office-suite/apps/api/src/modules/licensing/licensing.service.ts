import { HttpStatus, Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { ErrorCodes } from '@orcax/api-contracts';
import { createHash, randomUUID } from 'crypto';
import { IsNull, Repository } from 'typeorm';
import { ApiError } from '../../common/http/api-error';
import { AccessContext } from '../../tenancy/access-context';
import { requireWorkspaceId } from '../../tenancy/workspace-scope';
import { AuditRecorder } from '../audit/audit-recorder.service';
import { AuthorizationService } from '../authorization/authorization.service';
import { deriveLicenseStatus, licenseAllowsUse, LicenseRuntimeStatus } from './license-catalog';
import { LicenseActivation } from './license-activation.entity';
import { License } from './license.entity';
import { PlanEntitlement } from './plan-entitlement.entity';
import { Plan } from './plan.entity';
import { Product } from './product.entity';

@Injectable()
export class LicensingService {
  constructor(
    @InjectRepository(Product)
    private readonly products: Repository<Product>,
    @InjectRepository(Plan)
    private readonly plans: Repository<Plan>,
    @InjectRepository(PlanEntitlement)
    private readonly planEntitlements: Repository<PlanEntitlement>,
    @InjectRepository(License)
    private readonly licenses: Repository<License>,
    @InjectRepository(LicenseActivation)
    private readonly activations: Repository<LicenseActivation>,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
  ) {}

  async listPlans(actor: AccessContext) {
    this.authorization.assertPolicy(actor, { permission: 'license.read' });
    const plans = await this.plans.find({
      where: { status: 'active', workspaceKind: actor.workspaceKind },
      order: { name: 'ASC' },
    });
    const result = [];
    for (const plan of plans) {
      const entitlements = await this.planEntitlements.find({ where: { planId: plan.id } });
      result.push(
        this.planView(
          plan,
          entitlements.map((item) => item.entitlementCode),
        ),
      );
    }
    return { plans: result };
  }

  async listLicenses(actor: AccessContext) {
    this.authorization.assertPolicy(actor, { permission: 'license.read' });
    const workspaceId = requireWorkspaceId(actor);
    const rows = await this.licenses.find({
      where: { workspaceId },
      order: { createdAt: 'DESC' },
    });
    const views = [];
    for (const license of rows) {
      views.push(await this.licenseView(license));
    }
    return { licenses: views };
  }

  async issue(actor: AccessContext, planId: string) {
    this.authorization.assertPolicy(actor, { permission: 'license.issue' });
    const workspaceId = requireWorkspaceId(actor);
    const plan = await this.plans.findOne({ where: { id: planId, status: 'active' } });
    if (!plan || plan.workspaceKind !== actor.workspaceKind) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Gói không khả dụng cho không gian hiện tại.',
      );
    }
    const now = new Date();
    const endsAt =
      plan.licenseType === 'perpetual' || plan.termDays === null
        ? null
        : new Date(now.getTime() + plan.termDays * 24 * 60 * 60 * 1000);
    const id = randomUUID();
    await this.licenses.insert({
      id,
      workspaceId,
      planId: plan.id,
      status: 'active',
      issuedBy: actor.userId,
      startsAt: now,
      endsAt,
      revokedAt: null,
      createdAt: now,
      updatedAt: now,
    });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'license.issued',
      targetType: 'license',
      targetId: id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { planCode: plan.code, licenseType: plan.licenseType },
    });
    const license = await this.licenses.findOneByOrFail({ id });
    return this.licenseView(license);
  }

  /** Trusted path for commerce after a verified paid payment. */
  async issueFromPaidOrder(input: {
    workspaceId: string;
    workspaceKind: 'personal' | 'organization';
    planId: string;
    issuedBy: string;
    correlationId: string;
    orderId: string;
  }) {
    const plan = await this.plans.findOne({ where: { id: input.planId, status: 'active' } });
    if (!plan || plan.workspaceKind !== input.workspaceKind) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Gói không khả dụng cho không gian hiện tại.',
      );
    }
    const now = new Date();
    const endsAt =
      plan.licenseType === 'perpetual' || plan.termDays === null
        ? null
        : new Date(now.getTime() + plan.termDays * 24 * 60 * 60 * 1000);
    const id = randomUUID();
    await this.licenses.insert({
      id,
      workspaceId: input.workspaceId,
      planId: plan.id,
      status: 'active',
      issuedBy: input.issuedBy,
      startsAt: now,
      endsAt,
      revokedAt: null,
      createdAt: now,
      updatedAt: now,
    });
    await this.recorder.append({
      workspaceId: input.workspaceId,
      actorId: input.issuedBy,
      action: 'license.issued',
      targetType: 'license',
      targetId: id,
      outcome: 'success',
      correlationId: input.correlationId,
      metadata: { planCode: plan.code, orderId: input.orderId, source: 'commerce' },
    });
    return this.licenseView(await this.licenses.findOneByOrFail({ id }));
  }

  async revoke(actor: AccessContext, licenseId: string) {
    this.authorization.assertPolicy(actor, { permission: 'license.manage' });
    const license = await this.requireWorkspaceLicense(actor, licenseId);
    if (license.revokedAt) {
      throw new ApiError(HttpStatus.CONFLICT, ErrorCodes.CONFLICT, 'Giấy phép đã bị thu hồi.');
    }
    const now = new Date();
    await this.licenses.manager.transaction(async (manager) => {
      await manager.update(
        License,
        { id: license.id },
        { status: 'revoked', revokedAt: now, updatedAt: now },
      );
      await manager.update(
        LicenseActivation,
        { licenseId: license.id, revokedAt: IsNull() },
        { revokedAt: now },
      );
      await this.recorder.append(
        {
          workspaceId: license.workspaceId,
          actorId: actor.userId,
          action: 'license.revoked',
          targetType: 'license',
          targetId: license.id,
          outcome: 'success',
          correlationId: actor.correlationId,
        },
        manager,
      );
    });
    return { revoked: true };
  }

  async renew(actor: AccessContext, licenseId: string) {
    this.authorization.assertPolicy(actor, { permission: 'license.manage' });
    const license = await this.requireWorkspaceLicense(actor, licenseId);
    if (license.revokedAt) {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Không thể gia hạn giấy phép đã thu hồi.',
      );
    }
    const plan = await this.plans.findOneByOrFail({ id: license.planId });
    if (plan.licenseType === 'perpetual' || plan.termDays === null) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Gói vĩnh viễn không cần gia hạn.',
      );
    }
    const now = new Date();
    const baseline = license.endsAt && license.endsAt > now ? license.endsAt : now;
    const endsAt = new Date(baseline.getTime() + plan.termDays * 24 * 60 * 60 * 1000);
    await this.licenses.update({ id: license.id }, { endsAt, status: 'active', updatedAt: now });
    await this.recorder.append({
      workspaceId: license.workspaceId,
      actorId: actor.userId,
      action: 'license.renewed',
      targetType: 'license',
      targetId: license.id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { endsAt: endsAt.toISOString() },
    });
    return this.licenseView(await this.licenses.findOneByOrFail({ id: license.id }));
  }

  async activate(
    actor: AccessContext,
    input: { licenseId: string; deviceKey: string; deviceLabel: string },
  ) {
    this.authorization.assertPolicy(actor, { permission: 'license.activate' });
    const license = await this.requireWorkspaceLicense(actor, input.licenseId);
    const plan = await this.plans.findOneByOrFail({ id: license.planId });
    const runtime = this.runtimeStatus(license, plan.graceDays);
    if (!licenseAllowsUse(runtime)) {
      throw new ApiError(HttpStatus.FORBIDDEN, ErrorCodes.FORBIDDEN, statusMessage(runtime));
    }
    const deviceKeyHash = sha256(input.deviceKey.trim());
    const existing = await this.activations.findOne({
      where: { licenseId: license.id, deviceKeyHash, revokedAt: IsNull() },
    });
    const now = new Date();
    if (existing) {
      await this.activations.update(
        { id: existing.id },
        { lastSeenAt: now, deviceLabel: input.deviceLabel.trim(), userId: actor.userId },
      );
      await this.recorder.append({
        workspaceId: license.workspaceId,
        actorId: actor.userId,
        action: 'license.activation.refreshed',
        targetType: 'license_activation',
        targetId: existing.id,
        outcome: 'success',
        correlationId: actor.correlationId,
      });
      return {
        activationId: existing.id,
        licenseId: license.id,
        status: runtime,
        deviceLimit: plan.deviceLimit,
      };
    }
    const liveCount = await this.activations.count({
      where: { licenseId: license.id, revokedAt: IsNull() },
    });
    if (liveCount >= plan.deviceLimit) {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Đã đạt giới hạn thiết bị của gói.',
      );
    }
    const id = randomUUID();
    await this.activations.insert({
      id,
      licenseId: license.id,
      workspaceId: license.workspaceId,
      userId: actor.userId,
      deviceKeyHash,
      deviceLabel: input.deviceLabel.trim(),
      activatedAt: now,
      lastSeenAt: now,
      revokedAt: null,
    });
    await this.recorder.append({
      workspaceId: license.workspaceId,
      actorId: actor.userId,
      action: 'license.activated',
      targetType: 'license_activation',
      targetId: id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { licenseId: license.id },
    });
    return {
      activationId: id,
      licenseId: license.id,
      status: runtime,
      deviceLimit: plan.deviceLimit,
    };
  }

  async listActivations(actor: AccessContext) {
    this.authorization.assertPolicy(actor, { permission: 'license.read' });
    const workspaceId = requireWorkspaceId(actor);
    const rows = await this.activations.find({
      where: { workspaceId, revokedAt: IsNull() },
      order: { activatedAt: 'DESC' },
    });
    return {
      activations: rows.map((row) => ({
        id: row.id,
        licenseId: row.licenseId,
        userId: row.userId,
        deviceLabel: row.deviceLabel,
        activatedAt: row.activatedAt.toISOString(),
        lastSeenAt: row.lastSeenAt.toISOString(),
        currentUser: row.userId === actor.userId,
      })),
    };
  }

  async revokeActivation(actor: AccessContext, activationId: string) {
    const workspaceId = requireWorkspaceId(actor);
    const activation = await this.activations.findOne({
      where: { id: activationId, workspaceId },
    });
    if (!activation || activation.revokedAt) {
      throw new ApiError(
        HttpStatus.NOT_FOUND,
        ErrorCodes.NOT_FOUND,
        'Không tìm thấy kích hoạt thiết bị.',
      );
    }
    const canManage =
      actor.roleCode === 'organization_owner' ||
      actor.roleCode === 'organization_administrator' ||
      actor.roleCode === 'personal_owner';
    if (!canManage && activation.userId !== actor.userId) {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Bạn chỉ thu hồi được thiết bị của chính mình.',
      );
    }
    this.authorization.assertPolicy(actor, {
      permission:
        canManage && activation.userId !== actor.userId ? 'license.manage' : 'license.activate',
    });
    await this.activations.update({ id: activation.id }, { revokedAt: new Date() });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'license.activation.revoked',
      targetType: 'license_activation',
      targetId: activation.id,
      outcome: 'success',
      correlationId: actor.correlationId,
    });
    return { revoked: true };
  }

  async checkEntitlement(actor: AccessContext, code: string) {
    this.authorization.assertPolicy(actor, { permission: 'license.read' });
    const workspaceId = requireWorkspaceId(actor);
    const normalized = code.trim().toLowerCase();
    const licenses = await this.licenses.find({ where: { workspaceId, status: 'active' } });
    for (const license of licenses) {
      const plan = await this.plans.findOneByOrFail({ id: license.planId });
      const runtime = this.runtimeStatus(license, plan.graceDays);
      if (!licenseAllowsUse(runtime)) {
        continue;
      }
      const hit = await this.planEntitlements.findOne({
        where: { planId: plan.id, entitlementCode: normalized },
      });
      if (hit) {
        return {
          allowed: true,
          entitlement: normalized,
          licenseId: license.id,
          status: runtime,
          planCode: plan.code,
        };
      }
    }
    return {
      allowed: false,
      entitlement: normalized,
      licenseId: null,
      status: null,
      planCode: null,
    };
  }

  private async requireWorkspaceLicense(actor: AccessContext, licenseId: string): Promise<License> {
    const workspaceId = requireWorkspaceId(actor);
    const license = await this.licenses.findOne({ where: { id: licenseId, workspaceId } });
    if (!license) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy giấy phép.');
    }
    return license;
  }

  private runtimeStatus(license: License, graceDays: number): LicenseRuntimeStatus {
    return deriveLicenseStatus({
      now: new Date(),
      startsAt: license.startsAt,
      endsAt: license.endsAt,
      revokedAt: license.revokedAt,
      graceDays,
    });
  }

  private async licenseView(license: License) {
    const plan = await this.plans.findOneByOrFail({ id: license.planId });
    const entitlements = await this.planEntitlements.find({ where: { planId: plan.id } });
    const runtime = this.runtimeStatus(license, plan.graceDays);
    const activationCount = await this.activations.count({
      where: { licenseId: license.id, revokedAt: IsNull() },
    });
    return {
      id: license.id,
      workspaceId: license.workspaceId,
      status: runtime,
      storedStatus: license.status,
      startsAt: license.startsAt.toISOString(),
      endsAt: license.endsAt ? license.endsAt.toISOString() : null,
      revokedAt: license.revokedAt ? license.revokedAt.toISOString() : null,
      activationCount,
      plan: this.planView(
        plan,
        entitlements.map((item) => item.entitlementCode),
      ),
      usable: licenseAllowsUse(runtime),
    };
  }

  private planView(plan: Plan, entitlements: string[]) {
    return {
      id: plan.id,
      code: plan.code,
      name: plan.name,
      licenseType: plan.licenseType,
      workspaceKind: plan.workspaceKind,
      deviceLimit: plan.deviceLimit,
      seatLimit: plan.seatLimit,
      termDays: plan.termDays,
      graceDays: plan.graceDays,
      entitlements,
    };
  }
}

function sha256(value: string): string {
  return createHash('sha256').update(value).digest('hex');
}

function statusMessage(status: LicenseRuntimeStatus): string {
  if (status === 'revoked') {
    return 'Giấy phép đã bị thu hồi.';
  }
  if (status === 'expired') {
    return 'Giấy phép đã hết hạn.';
  }
  if (status === 'pending') {
    return 'Giấy phép chưa đến ngày hiệu lực.';
  }
  return 'Giấy phép không khả dụng.';
}
