import { randomBytes, scrypt, timingSafeEqual } from 'crypto';

const SCRYPT_N = 16384;
const SCRYPT_R = 8;
const SCRYPT_P = 1;
const KEY_LENGTH = 32;

export async function hashSecret(value: string): Promise<string> {
  const salt = randomBytes(16);
  const derived = await deriveKey(value, salt, KEY_LENGTH);
  return [
    'scrypt',
    SCRYPT_N,
    SCRYPT_R,
    SCRYPT_P,
    salt.toString('base64url'),
    derived.toString('base64url'),
  ].join('$');
}

export async function verifySecret(value: string, encoded: string): Promise<boolean> {
  const parts = encoded.split('$');
  if (parts.length !== 6 || parts[0] !== 'scrypt') {
    return false;
  }
  const salt = Buffer.from(parts[4], 'base64url');
  const expected = Buffer.from(parts[5], 'base64url');
  const derived = await deriveKey(value, salt, expected.length, {
    N: Number(parts[1]),
    r: Number(parts[2]),
    p: Number(parts[3]),
  });
  if (derived.length !== expected.length) {
    return false;
  }
  return timingSafeEqual(derived, expected);
}

let dummyHash: Promise<string> | undefined;

export function dummyPasswordHash(): Promise<string> {
  dummyHash ??= hashSecret('not-a-user-password');
  return dummyHash;
}

function deriveKey(
  value: string,
  salt: Buffer,
  keyLength: number,
  options = { N: SCRYPT_N, r: SCRYPT_R, p: SCRYPT_P },
): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    scrypt(value, salt, keyLength, options, (error, derived) => {
      if (error) {
        reject(error);
        return;
      }
      resolve(derived);
    });
  });
}
