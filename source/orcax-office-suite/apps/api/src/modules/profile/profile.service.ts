import { HttpStatus, Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { ErrorCodes } from '@orcax/api-contracts';
import { createProfilePackage, validateProfilePackage } from '@orcax/profile-specifications';
import { buildValidationReport } from '@orcax/rule-definitions';
import { randomUUID } from 'crypto';
import { IsNull, Repository } from 'typeorm';
import { ApiError } from '../../common/http/api-error';
import { AccessContext } from '../../tenancy/access-context';
import { requireWorkspaceId } from '../../tenancy/workspace-scope';
import { AuditRecorder } from '../audit/audit-recorder.service';
import { AuthorizationService } from '../authorization/authorization.service';
import { can } from '../authorization/catalog';
import { DepartmentService } from '../organization/department.service';
import { LegalReference } from './legal-reference.entity';
import {
  AddLegalReferenceDto,
  CreateProfileDto,
  EvaluateProfileDto,
  ProfileRuleDto,
  VerifyLegalReferenceDto,
} from './profile.dto';
import {
  approvalBlockedByLegalReferences,
  assertProfileMutable,
  canTransitionProfile,
  nextProfileVersion,
  ProfileHost,
  ProfileStatus,
} from './profile-lifecycle';
import { ProfileDistribution } from './profile-distribution.entity';
import { ProfileRule } from './profile-rule.entity';
import { ProfileVersion } from './profile-version.entity';
import { Profile } from './profile.entity';
import { VerificationEvidence } from './verification-evidence.entity';

@Injectable()
export class ProfileService {
  constructor(
    @InjectRepository(Profile)
    private readonly profiles: Repository<Profile>,
    @InjectRepository(ProfileVersion)
    private readonly versions: Repository<ProfileVersion>,
    @InjectRepository(ProfileRule)
    private readonly rules: Repository<ProfileRule>,
    @InjectRepository(LegalReference)
    private readonly legalReferences: Repository<LegalReference>,
    @InjectRepository(VerificationEvidence)
    private readonly evidence: Repository<VerificationEvidence>,
    @InjectRepository(ProfileDistribution)
    private readonly distributions: Repository<ProfileDistribution>,
    private readonly departments: DepartmentService,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
  ) {}

  async list(actor: AccessContext, query: { host?: ProfileHost; status?: ProfileStatus } = {}) {
    this.authorization.assertPolicy(actor, { permission: 'profile.read' });
    const workspaceId = requireWorkspaceId(actor);
    const rows = await this.profiles.find({
      where: [{ workspaceId }, { scope: 'system', workspaceId: IsNull() }],
      order: { name: 'ASC' },
    });
    const targets = await this.departments.userTargets(workspaceId, actor.userId);
    const distributionRows = await this.distributions.find({ where: { workspaceId } });
    const byProfile = new Map<string, ProfileDistribution[]>();
    for (const row of distributionRows) {
      const list = byProfile.get(row.profileId) ?? [];
      list.push(row);
      byProfile.set(row.profileId, list);
    }
    const adminView = can(actor.roleCode, 'profile.publish');
    const filtered = rows.filter((row) => {
      if (query.host && row.host !== query.host) {
        return false;
      }
      if (query.status && row.status !== query.status) {
        return false;
      }
      if (row.scope === 'system' || adminView) {
        return true;
      }
      const scoped = byProfile.get(row.id) ?? [];
      if (scoped.length === 0) {
        return true;
      }
      return scoped.some((item) =>
        item.targetType === 'department'
          ? targets.departmentIds.includes(item.targetId)
          : targets.groupIds.includes(item.targetId),
      );
    });
    return { profiles: filtered.map((row) => this.summary(row)) };
  }

  async distribute(
    actor: AccessContext,
    profileId: string,
    input: { targetType: 'department' | 'group'; targetId: string },
  ) {
    this.authorization.assertPolicy(actor, {
      permission: 'profile.publish',
      workspaceKinds: ['organization'],
    });
    const profile = await this.requireOwnedProfile(actor, profileId);
    if (profile.status !== 'approved') {
      throw new ApiError(HttpStatus.CONFLICT, ErrorCodes.CONFLICT, 'Chỉ phân phối hồ sơ đã duyệt.');
    }
    const workspaceId = requireWorkspaceId(actor);
    const existing = await this.distributions.findOne({
      where: {
        profileId,
        workspaceId,
        targetType: input.targetType,
        targetId: input.targetId,
      },
    });
    if (existing) {
      return { id: existing.id, alreadyDistributed: true };
    }
    const id = randomUUID();
    await this.distributions.insert({
      id,
      profileId,
      workspaceId,
      targetType: input.targetType,
      targetId: input.targetId,
      distributedBy: actor.userId,
      distributedAt: new Date(),
    });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'profile.distributed',
      targetType: 'profile',
      targetId: profileId,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { targetType: input.targetType, targetId: input.targetId },
    });
    return { id, alreadyDistributed: false };
  }

  async listDistributions(actor: AccessContext, profileId: string) {
    this.authorization.assertPolicy(actor, { permission: 'profile.read' });
    await this.requireVisibleProfile(actor, profileId);
    const rows = await this.distributions.find({
      where: { profileId, workspaceId: requireWorkspaceId(actor) },
      order: { distributedAt: 'DESC' },
    });
    return {
      distributions: rows.map((row) => ({
        id: row.id,
        targetType: row.targetType,
        targetId: row.targetId,
        distributedBy: row.distributedBy,
        distributedAt: row.distributedAt.toISOString(),
      })),
    };
  }

  async get(actor: AccessContext, profileId: string) {
    this.authorization.assertPolicy(actor, { permission: 'profile.read' });
    const profile = await this.requireVisibleProfile(actor, profileId);
    return this.detail(profile);
  }

  async create(actor: AccessContext, body: CreateProfileDto) {
    this.authorization.assertPolicy(actor, { permission: 'profile.write' });
    const workspaceId = requireWorkspaceId(actor);
    const scope = actor.workspaceKind === 'personal' ? 'personal' : 'organization';
    const existing = await this.profiles.findOne({
      where: { workspaceId, code: body.code },
    });
    if (existing) {
      throw new ApiError(HttpStatus.CONFLICT, ErrorCodes.CONFLICT, 'Mã hồ sơ đã tồn tại.');
    }
    const now = new Date();
    const profileId = randomUUID();
    const versionId = randomUUID();
    await this.profiles.manager.transaction(async (manager) => {
      await manager.insert(Profile, {
        id: profileId,
        code: body.code,
        name: body.name,
        host: body.host,
        scope,
        workspaceId,
        ownerUserId: actor.userId,
        status: 'draft',
        currentVersion: 1,
        createdBy: actor.userId,
        createdAt: now,
        updatedAt: now,
      });
      await manager.insert(ProfileVersion, {
        id: versionId,
        profileId,
        version: 1,
        status: 'draft',
        source: body.source,
        effectiveAt: null,
        authorId: actor.userId,
        reviewerId: null,
        changeSummary: 'Tạo hồ sơ',
        immutable: false,
        createdAt: now,
        publishedAt: null,
      });
    });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'profile.created',
      targetType: 'profile',
      targetId: profileId,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { code: body.code, host: body.host },
    });
    return this.detail(await this.profiles.findOneByOrFail({ id: profileId }));
  }

  async replaceRules(actor: AccessContext, profileId: string, rules: ProfileRuleDto[]) {
    this.authorization.assertPolicy(actor, { permission: 'profile.write' });
    const profile = await this.requireOwnedProfile(actor, profileId);
    const version = await this.requireMutableHead(profile);
    await this.rules.delete({ profileVersionId: version.id });
    let order = 1;
    for (const rule of rules) {
      await this.rules.insert({
        id: randomUUID(),
        profileVersionId: version.id,
        ruleCode: rule.code,
        ruleVersion: rule.version,
        defaultSeverity: rule.defaultSeverity,
        title: rule.title,
        description: rule.description,
        checkKind: rule.checkKind,
        checkConfig: { ...(rule.checkConfig ?? {}) },
        sortOrder: order,
      });
      order += 1;
    }
    await this.touch(profile);
    return this.detail(profile);
  }

  async addLegalReference(actor: AccessContext, profileId: string, body: AddLegalReferenceDto) {
    this.authorization.assertPolicy(actor, { permission: 'profile.write' });
    const profile = await this.requireOwnedProfile(actor, profileId);
    const version = await this.requireMutableHead(profile);
    const id = randomUUID();
    await this.legalReferences.insert({
      id,
      profileVersionId: version.id,
      citation: body.citation,
      sourceUrl: body.sourceUrl ?? null,
      verified: false,
      notes: body.notes ?? null,
    });
    await this.touch(profile);
    return this.detail(await this.profiles.findOneByOrFail({ id: profile.id }));
  }

  async verifyLegalReference(
    actor: AccessContext,
    legalReferenceId: string,
    body: VerifyLegalReferenceDto,
  ) {
    this.authorization.assertPolicy(actor, { permission: 'profile.review' });
    const reference = await this.legalReferences.findOneBy({ id: legalReferenceId });
    if (!reference) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy trích dẫn.');
    }
    const version = await this.versions.findOneByOrFail({ id: reference.profileVersionId });
    const profile = await this.requireOwnedProfile(actor, version.profileId);
    if (!assertProfileMutable(version.status)) {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Phiên bản hồ sơ không còn chỉnh sửa được.',
      );
    }
    const now = new Date();
    await this.profiles.manager.transaction(async (manager) => {
      await manager.update(LegalReference, { id: reference.id }, { verified: true });
      await manager.insert(VerificationEvidence, {
        id: randomUUID(),
        legalReferenceId: reference.id,
        recordedBy: actor.userId,
        recordedAt: now,
        evidenceType: body.evidenceType,
        summary: body.summary,
        referenceUri: body.referenceUri ?? null,
      });
    });
    await this.recorder.append({
      workspaceId: requireWorkspaceId(actor),
      actorId: actor.userId,
      action: 'profile.legal.verified',
      targetType: 'legal_reference',
      targetId: reference.id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { profileId: profile.id },
    });
    return this.detail(await this.profiles.findOneByOrFail({ id: profile.id }));
  }

  async submitReview(actor: AccessContext, profileId: string, changeSummary?: string) {
    return this.transition(actor, profileId, 'in_review', 'profile.write', changeSummary);
  }

  async approve(actor: AccessContext, profileId: string, changeSummary?: string) {
    this.authorization.assertPolicy(actor, { permission: 'profile.publish' });
    const profile = await this.requireOwnedProfile(actor, profileId);
    const version = await this.requireHeadVersion(profile);
    if (!canTransitionProfile(version.status, 'approved')) {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Hồ sơ chưa ở trạng thái chờ duyệt.',
      );
    }
    const refs = await this.legalReferences.find({ where: { profileVersionId: version.id } });
    if (approvalBlockedByLegalReferences(refs)) {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Không thể duyệt khi còn trích dẫn pháp lý chưa xác minh.',
      );
    }
    if (!version.source.trim()) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Nguồn hồ sơ là bắt buộc trước khi duyệt.',
      );
    }
    const now = new Date();
    await this.profiles.manager.transaction(async (manager) => {
      await manager.update(
        ProfileVersion,
        { id: version.id },
        {
          status: 'approved',
          immutable: true,
          reviewerId: actor.userId,
          changeSummary: changeSummary ?? version.changeSummary,
          publishedAt: now,
          effectiveAt: version.effectiveAt ?? now,
        },
      );
      await manager.update(Profile, { id: profile.id }, { status: 'approved', updatedAt: now });
    });
    await this.recorder.append({
      workspaceId: requireWorkspaceId(actor),
      actorId: actor.userId,
      action: 'profile.published',
      targetType: 'profile',
      targetId: profile.id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { version: version.version },
    });
    return this.detail(await this.profiles.findOneByOrFail({ id: profile.id }));
  }

  async deprecate(actor: AccessContext, profileId: string, changeSummary?: string) {
    return this.transition(actor, profileId, 'deprecated', 'profile.publish', changeSummary);
  }

  async newDraftVersion(actor: AccessContext, profileId: string, changeSummary?: string) {
    this.authorization.assertPolicy(actor, { permission: 'profile.write' });
    const profile = await this.requireOwnedProfile(actor, profileId);
    if (profile.status !== 'approved' && profile.status !== 'deprecated') {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Chỉ tạo phiên bản mới từ hồ sơ đã duyệt hoặc đã ngừng.',
      );
    }
    const head = await this.requireHeadVersion(profile);
    const rules = await this.rules.find({
      where: { profileVersionId: head.id },
      order: { sortOrder: 'ASC' },
    });
    const refs = await this.legalReferences.find({ where: { profileVersionId: head.id } });
    const now = new Date();
    const versionNumber = nextProfileVersion(profile.currentVersion);
    const versionId = randomUUID();
    await this.profiles.manager.transaction(async (manager) => {
      await manager.insert(ProfileVersion, {
        id: versionId,
        profileId: profile.id,
        version: versionNumber,
        status: 'draft',
        source: head.source,
        effectiveAt: null,
        authorId: actor.userId,
        reviewerId: null,
        changeSummary: changeSummary ?? `Phiên bản ${versionNumber}`,
        immutable: false,
        createdAt: now,
        publishedAt: null,
      });
      for (const rule of rules) {
        await manager.insert(ProfileRule, {
          ...rule,
          id: randomUUID(),
          profileVersionId: versionId,
        });
      }
      for (const ref of refs) {
        await manager.insert(LegalReference, {
          id: randomUUID(),
          profileVersionId: versionId,
          citation: ref.citation,
          sourceUrl: ref.sourceUrl,
          verified: false,
          notes: ref.notes,
        });
      }
      await manager.update(
        Profile,
        { id: profile.id },
        { currentVersion: versionNumber, status: 'draft', updatedAt: now },
      );
    });
    return this.detail(await this.profiles.findOneByOrFail({ id: profile.id }));
  }

  async exportPackage(actor: AccessContext, profileId: string) {
    this.authorization.assertPolicy(actor, { permission: 'profile.read' });
    const profile = await this.requireVisibleProfile(actor, profileId);
    const version = await this.requireHeadVersion(profile);
    const rules = await this.rules.find({
      where: { profileVersionId: version.id },
      order: { sortOrder: 'ASC' },
    });
    const refs = await this.legalReferences.find({ where: { profileVersionId: version.id } });
    return createProfilePackage({
      profile: {
        code: profile.code,
        name: profile.name,
        host: profile.host,
        scope: profile.scope,
        source: version.source,
        effectiveAt: version.effectiveAt ? version.effectiveAt.toISOString() : null,
      },
      rules: rules.map((rule) => ({
        code: rule.ruleCode,
        version: rule.ruleVersion,
        title: rule.title,
        description: rule.description,
        checkKind: rule.checkKind,
        defaultSeverity: rule.defaultSeverity,
        checkConfig: { ...(rule.checkConfig as Record<string, unknown>) },
        sortOrder: rule.sortOrder,
      })),
      legalReferences: refs.map((item) => ({
        citation: item.citation,
        sourceUrl: item.sourceUrl,
        verified: item.verified,
        notes: item.notes,
      })),
    });
  }

  async importPackage(actor: AccessContext, raw: Record<string, unknown>) {
    this.authorization.assertPolicy(actor, { permission: 'profile.write' });
    const validation = validateProfilePackage(raw);
    if (!validation.ok) {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.VALIDATION_FAILED,
        validation.errors[0] || 'Gói hồ sơ không hợp lệ.',
      );
    }
    const pkg = raw as {
      profile: {
        code: string;
        name: string;
        host: ProfileHost;
        source?: string | null;
      };
      rules: ProfileRuleDto[];
      legalReferences?: Array<{
        citation: string;
        sourceUrl?: string | null;
        notes?: string | null;
      }>;
    };
    const created = await this.create(actor, {
      code: pkg.profile.code,
      name: pkg.profile.name,
      host: pkg.profile.host,
      source: pkg.profile.source || 'Nhập từ gói hồ sơ',
    });
    await this.replaceRules(
      actor,
      created.id,
      (pkg.rules || []).map((rule) => ({
        code: rule.code,
        version: rule.version,
        title: rule.title || rule.code,
        description: rule.description || '',
        checkKind: rule.checkKind,
        defaultSeverity: rule.defaultSeverity,
        checkConfig: rule.checkConfig,
      })),
    );
    for (const ref of pkg.legalReferences || []) {
      await this.addLegalReference(actor, created.id, {
        citation: ref.citation,
        sourceUrl: ref.sourceUrl ?? undefined,
        notes: ref.notes ?? undefined,
      });
    }
    return this.detail(await this.profiles.findOneByOrFail({ id: created.id }));
  }

  async evaluate(actor: AccessContext, profileId: string, body: EvaluateProfileDto) {
    this.authorization.assertPolicy(actor, { permission: 'profile.read' });
    const profile = await this.requireVisibleProfile(actor, profileId);
    if (profile.status !== 'approved') {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Chỉ đánh giá trên hồ sơ đã duyệt.',
      );
    }
    const version = await this.requireHeadVersion(profile);
    return buildValidationReport({
      profileId: profile.id,
      profileVersionId: version.id,
      profileVersion: version.version,
      profileCode: profile.code,
      results: body.results as never,
    });
  }

  private async transition(
    actor: AccessContext,
    profileId: string,
    to: ProfileStatus,
    permission: 'profile.write' | 'profile.publish',
    changeSummary?: string,
  ) {
    this.authorization.assertPolicy(actor, { permission });
    const profile = await this.requireOwnedProfile(actor, profileId);
    const version = await this.requireHeadVersion(profile);
    if (!canTransitionProfile(version.status, to)) {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Không thể chuyển trạng thái hồ sơ.',
      );
    }
    const now = new Date();
    await this.profiles.manager.transaction(async (manager) => {
      await manager.update(
        ProfileVersion,
        { id: version.id },
        {
          status: to,
          changeSummary: changeSummary ?? version.changeSummary,
          reviewerId: to === 'deprecated' ? actor.userId : version.reviewerId,
        },
      );
      await manager.update(Profile, { id: profile.id }, { status: to, updatedAt: now });
    });
    await this.recorder.append({
      workspaceId: requireWorkspaceId(actor),
      actorId: actor.userId,
      action: `profile.${to}`,
      targetType: 'profile',
      targetId: profile.id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { version: version.version },
    });
    return this.detail(await this.profiles.findOneByOrFail({ id: profile.id }));
  }

  private async requireVisibleProfile(actor: AccessContext, profileId: string) {
    const workspaceId = requireWorkspaceId(actor);
    const profile = await this.profiles.findOneBy({ id: profileId });
    if (!profile) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy hồ sơ.');
    }
    if (profile.scope === 'system' && profile.workspaceId === null) {
      return profile;
    }
    if (profile.workspaceId !== workspaceId) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy hồ sơ.');
    }
    if (can(actor.roleCode, 'profile.publish')) {
      return profile;
    }
    const scoped = await this.distributions.find({ where: { profileId, workspaceId } });
    if (scoped.length === 0) {
      return profile;
    }
    const targets = await this.departments.userTargets(workspaceId, actor.userId);
    const allowed = scoped.some((item) =>
      item.targetType === 'department'
        ? targets.departmentIds.includes(item.targetId)
        : targets.groupIds.includes(item.targetId),
    );
    if (!allowed) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy hồ sơ.');
    }
    return profile;
  }

  private async requireOwnedProfile(actor: AccessContext, profileId: string) {
    const profile = await this.requireVisibleProfile(actor, profileId);
    if (profile.scope === 'system') {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Không thể chỉnh sửa hồ sơ hệ thống.',
      );
    }
    return profile;
  }

  private async requireHeadVersion(profile: Profile) {
    const version = await this.versions.findOne({
      where: { profileId: profile.id, version: profile.currentVersion },
    });
    if (!version) {
      throw new ApiError(
        HttpStatus.INTERNAL_SERVER_ERROR,
        ErrorCodes.INTERNAL_ERROR,
        'Thiếu phiên bản hồ sơ.',
      );
    }
    return version;
  }

  private async requireMutableHead(profile: Profile) {
    const version = await this.requireHeadVersion(profile);
    if (!assertProfileMutable(version.status) || version.immutable) {
      throw new ApiError(HttpStatus.CONFLICT, ErrorCodes.CONFLICT, 'Phiên bản hồ sơ đã khóa.');
    }
    return version;
  }

  private async touch(profile: Profile) {
    await this.profiles.update({ id: profile.id }, { updatedAt: new Date() });
  }

  private summary(profile: Profile) {
    return {
      id: profile.id,
      code: profile.code,
      name: profile.name,
      host: profile.host,
      scope: profile.scope,
      status: profile.status,
      currentVersion: profile.currentVersion,
      workspaceId: profile.workspaceId,
    };
  }

  private async detail(profile: Profile) {
    const version = await this.requireHeadVersion(profile);
    const rules = await this.rules.find({
      where: { profileVersionId: version.id },
      order: { sortOrder: 'ASC' },
    });
    const refs = await this.legalReferences.find({ where: { profileVersionId: version.id } });
    const evidenceRows = [];
    for (const ref of refs) {
      const rows = await this.evidence.find({
        where: { legalReferenceId: ref.id },
        order: { recordedAt: 'DESC' },
      });
      evidenceRows.push(
        ...rows.map((item) => ({
          id: item.id,
          legalReferenceId: item.legalReferenceId,
          evidenceType: item.evidenceType,
          summary: item.summary,
          referenceUri: item.referenceUri,
          recordedAt: item.recordedAt.toISOString(),
        })),
      );
    }
    return {
      ...this.summary(profile),
      version: {
        id: version.id,
        version: version.version,
        status: version.status,
        source: version.source,
        effectiveAt: version.effectiveAt ? version.effectiveAt.toISOString() : null,
        authorId: version.authorId,
        reviewerId: version.reviewerId,
        changeSummary: version.changeSummary,
        immutable: version.immutable,
        publishedAt: version.publishedAt ? version.publishedAt.toISOString() : null,
      },
      rules: rules.map((rule) => ({
        id: rule.id,
        code: rule.ruleCode,
        version: rule.ruleVersion,
        title: rule.title,
        description: rule.description,
        checkKind: rule.checkKind,
        defaultSeverity: rule.defaultSeverity,
        checkConfig: rule.checkConfig,
        sortOrder: rule.sortOrder,
      })),
      legalReferences: refs.map((item) => ({
        id: item.id,
        citation: item.citation,
        sourceUrl: item.sourceUrl,
        verified: item.verified,
        notes: item.notes,
      })),
      verificationEvidence: evidenceRows,
    };
  }
}
