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
import { HelpArticle } from './help-article.entity';
import { SupportComment } from './support-comment.entity';
import { SupportTicket, TicketStatus } from './support-ticket.entity';
import { AddCommentDto, CreateHelpArticleDto, CreateTicketDto } from './support.dto';

@Injectable()
export class SupportService {
  constructor(
    @InjectRepository(SupportTicket)
    private readonly tickets: Repository<SupportTicket>,
    @InjectRepository(SupportComment)
    private readonly comments: Repository<SupportComment>,
    @InjectRepository(HelpArticle)
    private readonly helpArticles: Repository<HelpArticle>,
    private readonly authorization: AuthorizationService,
    private readonly recorder: AuditRecorder,
  ) {}

  async createTicket(actor: AccessContext, dto: CreateTicketDto) {
    this.authorization.assertPolicy(actor, { permission: 'support.ticket.create' });
    const workspaceId = requireWorkspaceId(actor);
    const now = new Date();
    const ticketId = randomUUID();
    const commentId = randomUUID();

    await this.tickets.manager.transaction(async (manager) => {
      await manager.insert(SupportTicket, {
        id: ticketId,
        workspaceId,
        userId: actor.userId,
        title: dto.title,
        category: dto.category,
        priority: dto.priority,
        status: 'open',
        createdAt: now,
        updatedAt: now,
      });

      await manager.insert(SupportComment, {
        id: commentId,
        ticketId,
        userId: actor.userId,
        message: dto.description,
        isStaff: false,
        createdAt: now,
      });
    });

    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'support.ticket.created',
      targetType: 'support_ticket',
      targetId: ticketId,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { category: dto.category, priority: dto.priority },
    });

    return this.getTicket(actor, ticketId);
  }

  async listTickets(actor: AccessContext) {
    this.authorization.assertPolicy(actor, { permission: 'support.ticket.read' });
    const workspaceId = requireWorkspaceId(actor);
    const rows = await this.tickets.find({
      where: { workspaceId },
      order: { createdAt: 'DESC' },
    });
    return {
      tickets: rows.map((ticket) => ({
        id: ticket.id,
        title: ticket.title,
        category: ticket.category,
        priority: ticket.priority,
        status: ticket.status,
        userId: ticket.userId,
        createdAt: ticket.createdAt.toISOString(),
        updatedAt: ticket.updatedAt.toISOString(),
      })),
    };
  }

  async getTicket(actor: AccessContext, id: string) {
    this.authorization.assertPolicy(actor, { permission: 'support.ticket.read' });
    const workspaceId = requireWorkspaceId(actor);
    const ticket = await this.tickets.findOne({ where: { id, workspaceId } });
    if (!ticket) {
      throw new ApiError(
        HttpStatus.NOT_FOUND,
        ErrorCodes.NOT_FOUND,
        'Không tìm thấy yêu cầu hỗ trợ.',
      );
    }

    const comments = await this.comments.find({
      where: { ticketId: id },
      order: { createdAt: 'ASC' },
    });

    return {
      id: ticket.id,
      title: ticket.title,
      category: ticket.category,
      priority: ticket.priority,
      status: ticket.status,
      userId: ticket.userId,
      createdAt: ticket.createdAt.toISOString(),
      updatedAt: ticket.updatedAt.toISOString(),
      comments: comments.map((c) => ({
        id: c.id,
        userId: c.userId,
        message: c.message,
        isStaff: c.isStaff,
        createdAt: c.createdAt.toISOString(),
      })),
    };
  }

  async addComment(actor: AccessContext, ticketId: string, dto: AddCommentDto) {
    this.authorization.assertPolicy(actor, { permission: 'support.ticket.create' });
    const workspaceId = requireWorkspaceId(actor);
    const ticket = await this.tickets.findOne({ where: { id: ticketId, workspaceId } });
    if (!ticket) {
      throw new ApiError(
        HttpStatus.NOT_FOUND,
        ErrorCodes.NOT_FOUND,
        'Không tìm thấy yêu cầu hỗ trợ.',
      );
    }

    const now = new Date();
    const commentId = randomUUID();
    const isStaff =
      actor.roleCode === 'organization_administrator' || actor.roleCode === 'organization_owner';

    await this.comments.insert({
      id: commentId,
      ticketId,
      userId: actor.userId,
      message: dto.message,
      isStaff,
      createdAt: now,
    });

    await this.tickets.update({ id: ticketId }, { updatedAt: now });

    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'support.ticket.commented',
      targetType: 'support_ticket',
      targetId: ticketId,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { commentId, isStaff },
    });

    return { id: commentId, message: dto.message, isStaff, createdAt: now.toISOString() };
  }

  async updateStatus(actor: AccessContext, ticketId: string, status: TicketStatus) {
    this.authorization.assertPolicy(actor, { permission: 'support.ticket.manage' });
    const workspaceId = requireWorkspaceId(actor);
    const ticket = await this.tickets.findOne({ where: { id: ticketId, workspaceId } });
    if (!ticket) {
      throw new ApiError(
        HttpStatus.NOT_FOUND,
        ErrorCodes.NOT_FOUND,
        'Không tìm thấy yêu cầu hỗ trợ.',
      );
    }

    const now = new Date();
    await this.tickets.update({ id: ticketId }, { status, updatedAt: now });

    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'support.ticket.status_changed',
      targetType: 'support_ticket',
      targetId: ticketId,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { previousStatus: ticket.status, nextStatus: status },
    });

    return { id: ticketId, status, updatedAt: now.toISOString() };
  }

  async listHelpArticles(category?: string) {
    const where: { status: 'published'; category?: string } = { status: 'published' };
    if (category) {
      where.category = category;
    }
    const articles = await this.helpArticles.find({
      where,
      order: { createdAt: 'ASC' },
    });
    return {
      articles: articles.map((article) => ({
        id: article.id,
        slug: article.slug,
        title: article.title,
        category: article.category,
        productVersion: article.productVersion,
        createdAt: article.createdAt.toISOString(),
      })),
    };
  }

  async getHelpArticle(slug: string) {
    const article = await this.helpArticles.findOne({
      where: { slug, status: 'published' },
    });
    if (!article) {
      throw new ApiError(
        HttpStatus.NOT_FOUND,
        ErrorCodes.NOT_FOUND,
        'Không tìm thấy bài viết hỗ trợ.',
      );
    }
    return {
      id: article.id,
      slug: article.slug,
      title: article.title,
      category: article.category,
      content: article.content,
      productVersion: article.productVersion,
      createdAt: article.createdAt.toISOString(),
      updatedAt: article.updatedAt.toISOString(),
    };
  }

  async createHelpArticle(actor: AccessContext, dto: CreateHelpArticleDto) {
    this.authorization.assertPolicy(actor, { permission: 'support.help.manage' });
    const existing = await this.helpArticles.findOne({ where: { slug: dto.slug } });
    if (existing) {
      throw new ApiError(HttpStatus.CONFLICT, ErrorCodes.CONFLICT, 'Slug bài viết đã tồn tại.');
    }

    const now = new Date();
    const id = randomUUID();
    await this.helpArticles.insert({
      id,
      slug: dto.slug,
      title: dto.title,
      category: dto.category,
      content: dto.content,
      productVersion: dto.productVersion ?? '1.0',
      status: 'published',
      createdAt: now,
      updatedAt: now,
    });

    const workspaceId = actor.workspaceId ?? '00000000-0000-0000-0000-000000000000';
    await this.recorder.append({
      workspaceId,
      actorId: actor.userId,
      action: 'support.help.created',
      targetType: 'help_article',
      targetId: id,
      outcome: 'success',
      correlationId: actor.correlationId,
      metadata: { slug: dto.slug, title: dto.title },
    });

    return this.getHelpArticle(dto.slug);
  }
}
