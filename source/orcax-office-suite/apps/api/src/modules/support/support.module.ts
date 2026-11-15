import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuditModule } from '../audit/audit.module';
import { AuthorizationModule } from '../authorization/authorization.module';
import { HelpArticle } from './help-article.entity';
import { SupportComment } from './support-comment.entity';
import { SupportController } from './support.controller';
import { SupportService } from './support.service';
import { SupportTicket } from './support-ticket.entity';

/** Support tickets and help content. Part of E09. */
@Module({
  imports: [
    TypeOrmModule.forFeature([SupportTicket, SupportComment, HelpArticle]),
    AuthorizationModule,
    AuditModule,
  ],
  controllers: [SupportController],
  providers: [SupportService],
  exports: [SupportService],
})
export class SupportModule {}
