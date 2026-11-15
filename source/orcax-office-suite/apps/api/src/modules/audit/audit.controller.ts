import { Controller, Get, Query } from '@nestjs/common';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { RequirePermission } from '../authorization/require-permission.decorator';
import { AuditQueryDto } from './audit-query.dto';
import { AuditQueryService } from './audit-query.service';

@Controller('audit-events')
export class AuditController {
  constructor(private readonly audits: AuditQueryService) {}

  @Get()
  @RequirePermission('workspace.audit.read')
  search(@CurrentAccess() access: AccessContext, @Query() query: AuditQueryDto) {
    return this.audits.search(access, query);
  }
}
