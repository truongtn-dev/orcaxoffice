import { createHash, timingSafeEqual } from 'crypto';

export const SESSION_COOKIE = 'orcax_session';

export function sha256(value: string): string {
  return createHash('sha256').update(value).digest('hex');
}

export function sha256Equal(value: string, expectedHex: string): boolean {
  const actual = createHash('sha256').update(value).digest();
  let expected: Buffer;
  try {
    expected = Buffer.from(expectedHex, 'hex');
  } catch {
    return false;
  }
  if (actual.length !== expected.length || expectedHex.length !== actual.length * 2) {
    return false;
  }
  return timingSafeEqual(actual, expected);
}

export function readCookie(header: string | undefined, name: string): string | undefined {
  if (!header) {
    return undefined;
  }
  for (const part of header.split(';')) {
    const separator = part.indexOf('=');
    if (separator === -1) {
      continue;
    }
    const key = part.slice(0, separator).trim();
    if (key !== name) {
      continue;
    }
    return decodeURIComponent(part.slice(separator + 1).trim());
  }
  return undefined;
}

export function sessionCookie(
  value: string,
  options: { maxAgeSeconds: number; secure: boolean },
): string {
  const parts = [
    `${SESSION_COOKIE}=${encodeURIComponent(value)}`,
    'HttpOnly',
    'Path=/api/v1',
    'SameSite=Lax',
    `Max-Age=${options.maxAgeSeconds}`,
  ];
  if (options.secure) {
    parts.push('Secure');
  }
  return parts.join('; ');
}

export function clearSessionCookie(secure: boolean): string {
  return sessionCookie('', { maxAgeSeconds: 0, secure });
}

export function packSession(sessionId: string, secret: string): string {
  return `${sessionId}.${secret}`;
}

export function unpackSession(value: string): { sessionId: string; secret: string } | null {
  const separator = value.indexOf('.');
  if (separator <= 0) {
    return null;
  }
  const sessionId = value.slice(0, separator);
  const secret = value.slice(separator + 1);
  if (!secret) {
    return null;
  }
  return { sessionId, secret };
}
