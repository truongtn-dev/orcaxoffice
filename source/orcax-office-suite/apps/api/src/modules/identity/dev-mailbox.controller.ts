import { Controller, Get, NotFoundException, Query } from '@nestjs/common';
import { EmailDto } from './auth.dto';
import { EmailSender } from './email-sender';

@Controller('dev/mailbox')
export class DevMailboxController {
  constructor(private readonly email: EmailSender) {}

  @Get()
  list(@Query() query: EmailDto) {
    if (!this.email.devMailboxEnabled()) {
      throw new NotFoundException('Không tìm thấy.');
    }
    return { messages: this.email.listDevMailbox(query.email.trim().toLowerCase()) };
  }
}
