import { resolve } from 'path';

export function repoEnvFilePath(): string {
  return resolve(__dirname, '../../../../.env');
}
