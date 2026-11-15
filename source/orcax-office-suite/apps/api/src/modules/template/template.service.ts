import { HttpStatus, Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { ErrorCodes } from '@orcax/api-contracts';
import { randomUUID } from 'crypto';
import { IsNull, Repository } from 'typeorm';
import { ApiError } from '../../common/http/api-error';
import { AccessContext } from '../../tenancy/access-context';
import { requireWorkspaceId } from '../../tenancy/workspace-scope';
import { AuditRecorder } from '../audit/audit-recorder.service';
import { AuthorizationService } from '../authorization/authorization.service';
import {
  canTransitionTemplate,
  nextProfileVersion,
  templateAvailableForGovernedUse,
  TemplateStatus,
} from '../profile/profile-lifecycle';
import { CreateTemplateDto, DataMappingDto, FieldDefinitionDto } from '../profile/profile.dto';
import { DataMapping } from './data-mapping.entity';
import { FieldDefinition } from './field-definition.entity';
import { TemplateVersion } from './template-version.entity';
import { Template } from './template.entity';

@Injectable()
export class TemplateService {
  constructor(
    @InjectRepository(Template)
    private readonly templates: Repository<Template>,
    @InjectRepository(TemplateVersion)
    private readonly versions: Repository<TemplateVersion>,
    @InjectRepository(FieldDefinition)
    private readonly fields: Repository<FieldDefinition>,
    @InjectRepository(DataMapping)
    private readonly mappings: Repository<DataMapping>,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
  ) {}

  async list(actor: AccessContext, host?: 'word' | 'excel') {
    this.authorization.assertPolicy(actor, { permission: 'template.read' });
    const workspaceId = requireWorkspaceId(actor);
    const rows = await this.templates.find({
      where: [{ workspaceId }, { scope: 'system', workspaceId: IsNull() }],
      order: { name: 'ASC' },
    });
    return {
      templates: rows
        .filter((row) => (host ? row.host === host : true))
        .map((row) => this.summary(row)),
    };
  }

  async get(actor: AccessContext, templateId: string) {
    this.authorization.assertPolicy(actor, { permission: 'template.read' });
    return this.detail(await this.requireVisible(actor, templateId));
  }

  async create(actor: AccessContext, body: CreateTemplateDto) {
    this.authorization.assertPolicy(actor, { permission: 'template.write' });
    const workspaceId = requireWorkspaceId(actor);
    const scope = actor.workspaceKind === 'personal' ? 'personal' : 'organization';
    const existing = await this.templates.findOne({ where: { workspaceId, code: body.code } });
    if (existing) {
      throw new ApiError(HttpStatus.CONFLICT, ErrorCodes.CONFLICT, 'Mã mẫu đã tồn tại.');
    }
    const now = new Date();
    const id = randomUUID();
    const versionId = randomUUID();
    await this.templates.manager.transaction(async (manager) => {
      await manager.insert(Template, {
        id,
        code: body.code,
        name: body.name,
        host: body.host,
        scope,
        workspaceId,
        status: 'draft',
        currentVersion: 1,
        createdBy: actor.userId,
        createdAt: now,
        updatedAt: now,
      });
      await manager.insert(TemplateVersion, {
        id: versionId,
        templateId: id,
        version: 1,
        status: 'draft',
        authorId: actor.userId,
        reviewerId: null,
        changeSummary: 'Tạo mẫu',
        immutable: false,
        createdAt: now,
        publishedAt: null,
      });
    });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'template.created',
      targetType: 'template',
      targetId: id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { code: body.code },
    });
    return this.detail(await this.templates.findOneByOrFail({ id }));
  }

  async replaceFields(actor: AccessContext, templateId: string, fields: FieldDefinitionDto[]) {
    this.authorization.assertPolicy(actor, { permission: 'template.write' });
    const template = await this.requireOwned(actor, templateId);
    const version = await this.requireMutableHead(template);
    await this.fields.delete({ templateVersionId: version.id });
    let order = 1;
    for (const field of fields) {
      await this.fields.insert({
        id: randomUUID(),
        templateVersionId: version.id,
        fieldKey: field.fieldKey,
        label: field.label,
        fieldType: field.fieldType,
        required: field.required !== false,
        defaultValue: field.defaultValue ?? null,
        validationJson: { ...(field.validationJson ?? {}) },
        sortOrder: order,
      });
      order += 1;
    }
    await this.templates.update({ id: template.id }, { updatedAt: new Date() });
    return this.detail(await this.templates.findOneByOrFail({ id: template.id }));
  }

  async replaceMappings(actor: AccessContext, templateId: string, mappings: DataMappingDto[]) {
    this.authorization.assertPolicy(actor, { permission: 'template.write' });
    const template = await this.requireOwned(actor, templateId);
    const version = await this.requireMutableHead(template);
    const fieldKeys = new Set(
      (await this.fields.find({ where: { templateVersionId: version.id } })).map(
        (item) => item.fieldKey,
      ),
    );
    for (const mapping of mappings) {
      if (!fieldKeys.has(mapping.fieldKey)) {
        throw new ApiError(
          HttpStatus.BAD_REQUEST,
          ErrorCodes.BAD_REQUEST,
          `Ánh xạ tham chiếu trường không tồn tại: ${mapping.fieldKey}`,
        );
      }
    }
    await this.mappings.delete({ templateVersionId: version.id });
    for (const mapping of mappings) {
      await this.mappings.insert({
        id: randomUUID(),
        templateVersionId: version.id,
        fieldKey: mapping.fieldKey,
        sourceKind: mapping.sourceKind,
        sourcePath: mapping.sourcePath,
        transform: mapping.transform,
      });
    }
    await this.templates.update({ id: template.id }, { updatedAt: new Date() });
    return this.detail(await this.templates.findOneByOrFail({ id: template.id }));
  }

  async submitReview(actor: AccessContext, templateId: string) {
    return this.transition(actor, templateId, 'review', 'template.write');
  }

  async approve(actor: AccessContext, templateId: string) {
    return this.transition(actor, templateId, 'approved', 'template.publish');
  }

  async publish(actor: AccessContext, templateId: string) {
    this.authorization.assertPolicy(actor, { permission: 'template.publish' });
    const template = await this.requireOwned(actor, templateId);
    const version = await this.requireHead(template);
    if (!canTransitionTemplate(version.status, 'published')) {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Chỉ phát hành mẫu đã được duyệt.',
      );
    }
    const now = new Date();
    await this.templates.manager.transaction(async (manager) => {
      await manager.update(
        TemplateVersion,
        { id: version.id },
        {
          status: 'published',
          immutable: true,
          reviewerId: actor.userId,
          publishedAt: now,
        },
      );
      await manager.update(Template, { id: template.id }, { status: 'published', updatedAt: now });
    });
    await this.recorder.append({
      workspaceId: requireWorkspaceId(actor),
      actorId: actor.userId,
      action: 'template.published',
      targetType: 'template',
      targetId: template.id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { version: version.version },
    });
    return this.detail(await this.templates.findOneByOrFail({ id: template.id }));
  }

  async deprecate(actor: AccessContext, templateId: string) {
    return this.transition(actor, templateId, 'deprecated', 'template.publish');
  }

  async newDraftVersion(actor: AccessContext, templateId: string) {
    this.authorization.assertPolicy(actor, { permission: 'template.write' });
    const template = await this.requireOwned(actor, templateId);
    if (template.status !== 'published' && template.status !== 'approved') {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Chỉ tạo phiên bản mới từ mẫu đã duyệt hoặc đã phát hành.',
      );
    }
    const head = await this.requireHead(template);
    const fields = await this.fields.find({
      where: { templateVersionId: head.id },
      order: { sortOrder: 'ASC' },
    });
    const mappings = await this.mappings.find({ where: { templateVersionId: head.id } });
    const now = new Date();
    const versionNumber = nextProfileVersion(template.currentVersion);
    const versionId = randomUUID();
    await this.templates.manager.transaction(async (manager) => {
      await manager.insert(TemplateVersion, {
        id: versionId,
        templateId: template.id,
        version: versionNumber,
        status: 'draft',
        authorId: actor.userId,
        reviewerId: null,
        changeSummary: `Phiên bản ${versionNumber}`,
        immutable: false,
        createdAt: now,
        publishedAt: null,
      });
      for (const field of fields) {
        await manager.insert(FieldDefinition, {
          ...field,
          id: randomUUID(),
          templateVersionId: versionId,
        });
      }
      for (const mapping of mappings) {
        await manager.insert(DataMapping, {
          ...mapping,
          id: randomUUID(),
          templateVersionId: versionId,
        });
      }
      await manager.update(
        Template,
        { id: template.id },
        { currentVersion: versionNumber, status: 'draft', updatedAt: now },
      );
    });
    return this.detail(await this.templates.findOneByOrFail({ id: template.id }));
  }

  mappingPreview(actor: AccessContext, templateId: string) {
    return this.get(actor, templateId).then((detail) => ({
      templateId: detail.id,
      templateVersion: detail.currentVersion,
      mappings: detail.mappings.map((item) => ({
        fieldKey: item.fieldKey,
        source: `${item.sourceKind}:${item.sourcePath}`,
        destination: item.fieldKey,
        transform: item.transform,
      })),
      governedUse: templateAvailableForGovernedUse(detail.status),
    }));
  }

  private async transition(
    actor: AccessContext,
    templateId: string,
    to: TemplateStatus,
    permission: 'template.write' | 'template.publish',
  ) {
    this.authorization.assertPolicy(actor, { permission });
    const template = await this.requireOwned(actor, templateId);
    const version = await this.requireHead(template);
    if (!canTransitionTemplate(version.status, to)) {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Không thể chuyển trạng thái mẫu.',
      );
    }
    const now = new Date();
    await this.templates.manager.transaction(async (manager) => {
      await manager.update(
        TemplateVersion,
        { id: version.id },
        {
          status: to,
          immutable: to === 'deprecated' ? true : version.immutable,
          reviewerId: actor.userId,
          publishedAt: to === 'published' ? now : version.publishedAt,
        },
      );
      await manager.update(Template, { id: template.id }, { status: to, updatedAt: now });
    });
    await this.recorder.append({
      workspaceId: requireWorkspaceId(actor),
      actorId: actor.userId,
      action: `template.${to}`,
      targetType: 'template',
      targetId: template.id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { version: version.version },
    });
    return this.detail(await this.templates.findOneByOrFail({ id: template.id }));
  }

  private async requireVisible(actor: AccessContext, templateId: string) {
    const workspaceId = requireWorkspaceId(actor);
    const template = await this.templates.findOneBy({ id: templateId });
    if (!template) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy mẫu.');
    }
    if (template.scope === 'system' && template.workspaceId === null) {
      return template;
    }
    if (template.workspaceId !== workspaceId) {
      throw new ApiError(HttpStatus.NOT_FOUND, ErrorCodes.NOT_FOUND, 'Không tìm thấy mẫu.');
    }
    return template;
  }

  private async requireOwned(actor: AccessContext, templateId: string) {
    const template = await this.requireVisible(actor, templateId);
    if (template.scope === 'system') {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Không thể chỉnh sửa mẫu hệ thống.',
      );
    }
    return template;
  }

  private async requireHead(template: Template) {
    const version = await this.versions.findOne({
      where: { templateId: template.id, version: template.currentVersion },
    });
    if (!version) {
      throw new ApiError(
        HttpStatus.INTERNAL_SERVER_ERROR,
        ErrorCodes.INTERNAL_ERROR,
        'Thiếu phiên bản mẫu.',
      );
    }
    return version;
  }

  private async requireMutableHead(template: Template) {
    const version = await this.requireHead(template);
    if (version.immutable || version.status === 'published' || version.status === 'deprecated') {
      throw new ApiError(HttpStatus.CONFLICT, ErrorCodes.CONFLICT, 'Phiên bản mẫu đã khóa.');
    }
    return version;
  }

  private summary(template: Template) {
    return {
      id: template.id,
      code: template.code,
      name: template.name,
      host: template.host,
      scope: template.scope,
      status: template.status,
      currentVersion: template.currentVersion,
      workspaceId: template.workspaceId,
      governedUse: templateAvailableForGovernedUse(template.status),
    };
  }

  private async detail(template: Template) {
    const version = await this.requireHead(template);
    const fields = await this.fields.find({
      where: { templateVersionId: version.id },
      order: { sortOrder: 'ASC' },
    });
    const mappings = await this.mappings.find({ where: { templateVersionId: version.id } });
    return {
      ...this.summary(template),
      version: {
        id: version.id,
        version: version.version,
        status: version.status,
        authorId: version.authorId,
        reviewerId: version.reviewerId,
        changeSummary: version.changeSummary,
        immutable: version.immutable,
        publishedAt: version.publishedAt ? version.publishedAt.toISOString() : null,
      },
      fields: fields.map((field) => ({
        id: field.id,
        fieldKey: field.fieldKey,
        label: field.label,
        fieldType: field.fieldType,
        required: field.required,
        defaultValue: field.defaultValue,
        validationJson: field.validationJson,
        sortOrder: field.sortOrder,
      })),
      mappings: mappings.map((item) => ({
        id: item.id,
        fieldKey: item.fieldKey,
        sourceKind: item.sourceKind,
        sourcePath: item.sourcePath,
        transform: item.transform,
      })),
    };
  }
}
