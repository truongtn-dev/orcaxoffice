import { Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { createHmac, timingSafeEqual } from 'crypto';
import { isUuid } from '../../common/ids/uuid';
import { AppConfig } from '../../config/validate-env';

export interface AccessClaims {
  sub: string;
  sid: string;
  wid: string;
  exp: number;
  iat: number;
}

@Injectable()
export class TokenService {
  constructor(private readonly config: ConfigService) {}

  sign(
    userId: string,
    sessionId: string,
    workspaceId: string,
  ): { token: string; expiresIn: number } {
    const now = Math.floor(Date.now() / 1000);
    const expiresIn = this.settings().accessTtlSeconds;
    const claims: AccessClaims = {
      sub: userId,
      sid: sessionId,
      wid: workspaceId,
      iat: now,
      exp: now + expiresIn,
    };
    const header = base64Url(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
    const body = base64Url(JSON.stringify(claims));
    const data = `${header}.${body}`;
    const signature = createHmac('sha256', this.settings().authTokenSecret)
      .update(data)
      .digest('base64url');
    return { token: `${data}.${signature}`, expiresIn };
  }

  verify(token: string): AccessClaims | null {
    const parts = token.split('.');
    if (parts.length !== 3 || !parts[0] || !parts[1] || !parts[2]) {
      return null;
    }
    const expected = createHmac('sha256', this.settings().authTokenSecret)
      .update(`${parts[0]}.${parts[1]}`)
      .digest();
    const actual = Buffer.from(parts[2], 'base64url');
    if (actual.length !== expected.length || !timingSafeEqual(actual, expected)) {
      return null;
    }
    let payload: Partial<AccessClaims>;
    try {
      payload = JSON.parse(
        Buffer.from(parts[1], 'base64url').toString('utf8'),
      ) as Partial<AccessClaims>;
    } catch {
      return null;
    }
    if (
      typeof payload.sub !== 'string' ||
      typeof payload.sid !== 'string' ||
      typeof payload.wid !== 'string' ||
      typeof payload.exp !== 'number' ||
      typeof payload.iat !== 'number'
    ) {
      return null;
    }
    if (!isUuid(payload.sub) || !isUuid(payload.sid) || !isUuid(payload.wid)) {
      return null;
    }
    if (payload.exp <= Math.floor(Date.now() / 1000)) {
      return null;
    }
    return {
      sub: payload.sub,
      sid: payload.sid,
      wid: payload.wid,
      exp: payload.exp,
      iat: payload.iat,
    };
  }

  private settings(): Pick<AppConfig, 'authTokenSecret' | 'accessTtlSeconds'> {
    return {
      authTokenSecret: this.config.getOrThrow<string>('authTokenSecret'),
      accessTtlSeconds: this.config.getOrThrow<number>('accessTtlSeconds'),
    };
  }
}

function base64Url(value: string): string {
  return Buffer.from(value).toString('base64url');
}
