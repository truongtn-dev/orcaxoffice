import { Module, forwardRef } from '@nestjs/common';
import { APP_GUARD } from '@nestjs/core';
import { TypeOrmModule } from '@nestjs/typeorm';
import { AuditModule } from '../audit/audit.module';
import { AuthorizationModule } from '../authorization/authorization.module';
import { OrganizationModule } from '../organization/organization.module';
import { AccessGuard } from './access.guard';
import { AuthController } from './auth.controller';
import { AuthService } from './auth.service';
import { ChallengeService } from './challenge.service';
import { DevMailboxController } from './dev-mailbox.controller';
import { EmailSender } from './email-sender';
import { SessionController } from './session.controller';
import { Session } from './session.entity';
import { SessionService } from './session.service';
import { TokenService } from './token.service';
import { UserDirectory } from './user-directory';
import { User } from './user.entity';
import { VerificationChallenge } from './verification-challenge.entity';
import { WorkspaceSwitchController } from './workspace-switch.controller';

@Module({
  imports: [
    TypeOrmModule.forFeature([User, Session, VerificationChallenge]),
    AuditModule,
    AuthorizationModule,
    forwardRef(() => OrganizationModule),
  ],
  controllers: [AuthController, SessionController, WorkspaceSwitchController, DevMailboxController],
  providers: [
    AuthService,
    SessionService,
    ChallengeService,
    TokenService,
    EmailSender,
    UserDirectory,
    { provide: APP_GUARD, useClass: AccessGuard },
  ],
  exports: [UserDirectory, EmailSender, SessionService],
})
export class IdentityModule {}
