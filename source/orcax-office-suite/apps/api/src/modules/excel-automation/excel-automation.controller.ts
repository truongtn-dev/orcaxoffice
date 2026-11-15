import { Body, Controller, Get, HttpCode, Post } from '@nestjs/common';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { RequirePermission } from '../authorization/require-permission.decorator';
import { RecordPrintRunDto } from '../word-automation/automation.dto';
import { ExcelAutomationService } from './excel-automation.service';

@Controller()
export class ExcelAutomationController {
  constructor(private readonly excel: ExcelAutomationService) {}

  @Get('add-in/excel/bootstrap')
  @RequirePermission('addin.excel.use')
  bootstrap(@CurrentAccess() access: AccessContext) {
    return this.excel.bootstrap(access);
  }

  @Post('add-in/excel/print-runs')
  @HttpCode(201)
  @RequirePermission('addin.excel.use')
  record(@CurrentAccess() access: AccessContext, @Body() body: RecordPrintRunDto) {
    return this.excel.recordPrintRun(access, body);
  }
}
