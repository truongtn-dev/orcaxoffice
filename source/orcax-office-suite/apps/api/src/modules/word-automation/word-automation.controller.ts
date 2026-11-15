import { Body, Controller, Get, HttpCode, Post } from '@nestjs/common';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { RequirePermission } from '../authorization/require-permission.decorator';
import { RecordValidationRunDto } from './automation.dto';
import { WordAutomationService } from './word-automation.service';

@Controller()
export class WordAutomationController {
  constructor(private readonly word: WordAutomationService) {}

  @Get('add-in/word/bootstrap')
  @RequirePermission('addin.word.use')
  bootstrap(@CurrentAccess() access: AccessContext) {
    return this.word.bootstrap(access);
  }

  @Post('add-in/word/validation-runs')
  @HttpCode(201)
  @RequirePermission('addin.word.use')
  record(@CurrentAccess() access: AccessContext, @Body() body: RecordValidationRunDto) {
    return this.word.recordValidation(access, body);
  }
}
