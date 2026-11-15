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
import { EXCEL_CAPABILITIES } from '../word-automation/add-in-capabilities';
import { RecordPrintRunDto } from '../word-automation/automation.dto';
import { PrintRun } from './print-run.entity';

@Injectable()
export class ExcelAutomationService {
  constructor(
    @InjectRepository(PrintRun)
    private readonly runs: Repository<PrintRun>,
    private readonly profiles: ProfileService,
    private readonly licensing: LicensingService,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
  ) {}

  async bootstrap(actor: AccessContext) {
    this.authorization.assertPolicy(actor, { permission: 'addin.excel.use' });
    const entitlement = await this.licensing.checkEntitlement(
      actor,
      EXCEL_CAPABILITIES.entitlement,
    );
    const { profiles } = await this.profiles.list(actor, { host: 'excel', status: 'approved' });
    return {
      host: 'excel' as const,
      workspaceId: requireWorkspaceId(actor),
      entitlement,
      capabilities: EXCEL_CAPABILITIES,
      profiles,
    };
  }

  async recordPrintRun(actor: AccessContext, body: RecordPrintRunDto) {
    this.authorization.assertPolicy(actor, { permission: 'addin.excel.use' });
    const entitlement = await this.licensing.checkEntitlement(
      actor,
      EXCEL_CAPABILITIES.entitlement,
    );
    if (!entitlement.allowed) {
      throw new ApiError(
        HttpStatus.FORBIDDEN,
        ErrorCodes.FORBIDDEN,
        'Workspace chưa có quyền lợi Excel.',
      );
    }
    const detail = await this.profiles.get(actor, body.profileId);
    if (detail.host !== 'excel' || detail.status !== 'approved') {
      throw new ApiError(
        HttpStatus.BAD_REQUEST,
        ErrorCodes.BAD_REQUEST,
        'Hồ sơ Excel phải ở trạng thái đã duyệt.',
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
      profileId: body.profileId,
      profileVersionId: body.profileVersionId,
      profileVersion: body.profileVersion,
      profileCode: body.profileCode,
      paperSize: body.paperSize,
      orientation: body.orientation,
      applied: body.applied,
      pdfExported: Boolean(body.pdfExported),
      adapter: body.adapter,
      createdAt: now,
    });
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'excel.print.recorded',
      targetType: 'print_run',
      targetId: id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: {
        profileCode: body.profileCode,
        profileVersion: body.profileVersion,
        applied: body.applied,
        pdfExported: Boolean(body.pdfExported),
        adapter: body.adapter,
      },
    });
    return {
      id,
      host: 'excel' as const,
      profileId: body.profileId,
      profileVersion: body.profileVersion,
      createdAt: now.toISOString(),
    };
  }
}
