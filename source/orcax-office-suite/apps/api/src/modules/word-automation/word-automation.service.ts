import { HttpStatus, Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { ErrorCodes } from '@orcax/api-contracts';
import { randomUUID } from 'crypto';
import { Repository } from 'typeorm';
import { ApiError } from '../../common/http/api-error';
import { AccessContext } from '../../tenancy/access-context';
import { requireWorkspaceId } from '../../tenancy/workspace-scope';
import { AuditRecorder } from '../audit/audit-recorder.service';
import { AuthorizationService } from '../authorization/authorization.service';
import { LicensingService } from '../licensing/licensing.service';
import { ProfileService } from '../profile/profile.service';
import { WORD_CAPABILITIES } from './add-in-capabilities';
import { RecordValidationRunDto } from './automation.dto';
import { ValidationRun } from './validation-run.entity';

@Injectable()
export class WordAutomationService {
  constructor(
    @InjectRepository(ValidationRun)
    private readonly runs: Repository<ValidationRun>,
    private readonly profiles: ProfileService,
    private readonly licensing: LicensingService,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
  ) {}

  async bootstrap(actor: AccessContext) {
    this.authorization.assertPolicy(actor, { permission: 'addin.word.use' });
    const entitlement = await this.licensing.checkEntitlement(actor, WORD_CAPABILITIES.entitlement);
    const { profiles } = await this.profiles.list(actor, { host: 'word', status: 'approved' });
    return {
      host: 'word' as const,
      workspaceId: requireWorkspaceId(actor),
      entitlement,
      capabilities: WORD_CAPABILITIES,
      profiles,
    };
  }

  async recordValidation(actor: AccessContext, body: RecordValidationRunDto) {
    this.authorization.assertPolicy(actor, { permission: 'addin.word.use' });
    const entitlement = await this.licensing.checkEntitlement(actor, WORD_CAPABILITIES.entitlement);
    if (!entitlement.allowed) {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Workspace chưa có quyền lợi Word.',
      );
    }
    const detail = await this.profiles.get(actor, body.profileId);
    if (detail.host !== 'word' || detail.status !== 'approved') {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Hồ sơ Word phải ở trạng thái đã duyệt.',
      );
    }
    if (
      detail.version.id !== body.profileVersionId ||
      detail.currentVersion !== body.profileVersion
    ) {
      throw new ApiError(
        HttpStatus.CONFLICT,
        ErrorCodes.CONFLICT,
        'Phiên bản hồ sơ không khớp phiên bản đang duyệt.',
      );
    }
    const workspaceId = requireWorkspaceId(actor);
    const id = randomUUID();
    const now = new Date();
    await this.runs.insert({
      id,
      workspaceId,
      userId: actor.userId,
      host: 'word',
      profileId: body.profileId,
      profileVersionId: body.profileVersionId,
      profileVersion: body.profileVersion,
      profileCode: body.profileCode,
      summaryPass: body.summary.pass,
      summaryFail: body.summary.fail,
      summaryWarning: body.summary.warning,
      summaryCorrected: body.summary.corrected,
      adapter: body.adapter,
      createdAt: now,
    });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'word.validation.recorded',
      targetType: 'validation_run',
      targetId: id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: {
        profileCode: body.profileCode,
        profileVersion: body.profileVersion,
        adapter: body.adapter,
        fail: body.summary.fail,
      },
    });
    return {
      id,
      host: 'word' as const,
      profileId: body.profileId,
      profileVersion: body.profileVersion,
      createdAt: now.toISOString(),
    };
  }
}
