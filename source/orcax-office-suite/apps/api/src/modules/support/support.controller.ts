import {
  Body,
  Controller,
  Get,
  HttpCode,
  Param,
  ParseUUIDPipe,
  Patch,
  Post,
  Query,
} from '@nestjs/common';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { RequirePermission } from '../authorization/require-permission.decorator';
import {
  AddCommentDto,
  CreateHelpArticleDto,
  CreateTicketDto,
  UpdateTicketStatusDto,
} from './support.dto';
import { SupportService } from './support.service';

@Controller('support')
export class SupportController {
  constructor(private readonly support: SupportService) {}

  @Get('tickets')
  @RequirePermission('support.ticket.read')
  listTickets(@CurrentAccess() access: AccessContext) {
    return this.support.listTickets(access);
  }

  @Post('tickets')
  @HttpCode(201)
  @RequirePermission('support.ticket.create')
  createTicket(@CurrentAccess() access: AccessContext, @Body() body: CreateTicketDto) {
    return this.support.createTicket(access, body);
  }

  @Get('tickets/:id')
  @RequirePermission('support.ticket.read')
  getTicket(@CurrentAccess() access: AccessContext, @Param('id', ParseUUIDPipe) id: string) {
    return this.support.getTicket(access, id);
  }

  @Post('tickets/:id/comments')
  @HttpCode(201)
  @RequirePermission('support.ticket.create')
  addComment(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: AddCommentDto,
  ) {
    return this.support.addComment(access, id, body);
  }

  @Patch('tickets/:id/status')
  @HttpCode(200)
  @RequirePermission('support.ticket.manage')
  updateStatus(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: UpdateTicketStatusDto,
  ) {
    return this.support.updateStatus(access, id, body.status);
  }

  @Get('help/articles')
  @RequirePermission('support.help.read')
  listHelpArticles(@Query('category') category?: string) {
    return this.support.listHelpArticles(category);
  }

  @Get('help/articles/:slug')
  @RequirePermission('support.help.read')
  getHelpArticle(@Param('slug') slug: string) {
    return this.support.getHelpArticle(slug);
  }

  @Post('help/articles')
  @HttpCode(201)
  @RequirePermission('support.help.manage')
  createHelpArticle(@CurrentAccess() access: AccessContext, @Body() body: CreateHelpArticleDto) {
    return this.support.createHelpArticle(access, body);
  }
}
