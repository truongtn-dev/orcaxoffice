import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuthorizationModule } from '../authorization/authorization.module';
import { AuditController } from './audit.controller';
import { AuditEvent } from './audit-event.entity';
import { AuditQueryService } from './audit-query.service';
import { AuditRecorder } from './audit-recorder.service';

@Module({
  imports: [TypeOrmModule.forFeature([AuditEvent]), AuthorizationModule],
  controllers: [AuditController],
  providers: [AuditRecorder, AuditQueryService],
  exports: [AuditRecorder],
})
export class AuditModule {}
