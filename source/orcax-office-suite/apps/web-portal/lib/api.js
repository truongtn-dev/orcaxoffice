import { clearAccessToken, readAccessToken, writeAccessToken } from './session';

const origin = process.env.NEXT_PUBLIC_API_ORIGIN || 'http://localhost:3001';

export class ApiRequestError extends Error {
  constructor(message, status, code) {
    super(message);
    this.name = 'ApiRequestError';
    this.status = status;
    this.code = code;
  }
}

export async function api(path, options = {}) {
  return request(path, options, false);
}

async function request(path, options, retried) {
  const headers = { accept: 'application/json' };
  if (options.body !== undefined) {
    headers['content-type'] = 'application/json';
  }
  const token = options.token ?? readAccessToken();
  if (token) {
    headers.authorization = `Bearer ${token}`;
  }
  let response;
  try {
    response = await fetch(`${origin}/api/v1${path}`, {
      method: options.method ?? 'GET',
      headers,
      body: options.body === undefined ? undefined : JSON.stringify(options.body),
      credentials: 'include',
    });
  } catch {
    throw new ApiRequestError('Không kết nối được máy chủ.', 0, 'NETWORK');
  }
  const payload = await response.json().catch(() => null);
  if (response.status === 401 && !options.skipRefresh && !retried && path !== '/auth/refresh') {
    const refreshed = await request('/auth/refresh', { method: 'POST', token: '' }, true);
    writeAccessToken(refreshed.accessToken);
    return request(path, { ...options, token: refreshed.accessToken }, true);
  }
  if (!response.ok) {
    throw new ApiRequestError(
      payload?.error?.message || 'Yêu cầu không thành công.',
      response.status,
      payload?.error?.code || 'REQUEST_FAILED',
    );
  }
  return payload?.data;
}
