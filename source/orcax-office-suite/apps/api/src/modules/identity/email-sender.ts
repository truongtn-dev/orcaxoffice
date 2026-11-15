import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { AppConfig } from '../../config/validate-env';

export interface MailMessage {
  to: string;
  subject: string;
  text: string;
  createdAt: string;
}

interface MailTransport {
  sendMail(input: { from: string; to: string; subject: string; text: string }): Promise<unknown>;
}

interface MailerModule {
  createTransport(options: Record<string, unknown>): MailTransport;
}

@Injectable()
export class EmailSender {
  private readonly logger = new Logger(EmailSender.name);
  private readonly mailbox: MailMessage[] = [];

  constructor(private readonly config: ConfigService) {}

  async send(input: { to: string; subject: string; text: string }): Promise<void> {
    if (this.config.getOrThrow<AppConfig['emailTransport']>('emailTransport') === 'log') {
      this.mailbox.push({ ...input, createdAt: new Date().toISOString() });
      if (this.mailbox.length > 100) {
        this.mailbox.shift();
      }
      this.logger.log('Development mailbox stored a message');
      return;
    }
    const settings = this.smtp();
    const mailer = asMailer(await import('nodemailer'));
    const transport = mailer.createTransport({
      host: settings.smtpHost,
      port: settings.smtpPort,
      secure: settings.smtpPort === 465,
      auth: settings.smtpUser
        ? { user: settings.smtpUser, pass: settings.smtpPassword }
        : undefined,
    });
    await transport.sendMail({
      from: settings.emailFrom,
      to: input.to,
      subject: input.subject,
      text: input.text,
    });
  }

  devMailboxEnabled(): boolean {
    return (
      this.config.getOrThrow<AppConfig['nodeEnv']>('nodeEnv') !== 'production' &&
      this.config.getOrThrow<AppConfig['emailTransport']>('emailTransport') === 'log'
    );
  }

  listDevMailbox(email: string): MailMessage[] {
    return this.mailbox.filter((message) => message.to === email).slice(-20);
  }

  private smtp(): Pick<
    AppConfig,
    'smtpHost' | 'smtpPort' | 'smtpUser' | 'smtpPassword' | 'emailFrom'
  > {
    return {
      smtpHost: this.config.getOrThrow<string>('smtpHost'),
      smtpPort: this.config.getOrThrow<number>('smtpPort'),
      smtpUser: this.config.getOrThrow<string>('smtpUser'),
      smtpPassword: this.config.getOrThrow<string>('smtpPassword'),
      emailFrom: this.config.getOrThrow<string>('emailFrom'),
    };
  }
}

function asMailer(loaded: unknown): MailerModule {
  const record = loaded as {
    createTransport?: MailerModule['createTransport'];
    default?: MailerModule;
  };
  if (record.createTransport) {
    return { createTransport: record.createTransport };
  }
  if (record.default?.createTransport) {
    return record.default;
  }
  throw new Error('SMTP transport is unavailable');
}
