const ACCESS_TOKEN_KEY = 'orcax.accessToken';

export function readAccessToken() {
  if (typeof sessionStorage === 'undefined') {
    return '';
  }
  return sessionStorage.getItem(ACCESS_TOKEN_KEY) ?? '';
}

export function writeAccessToken(token) {
  sessionStorage.setItem(ACCESS_TOKEN_KEY, token);
}

export function clearAccessToken() {
  sessionStorage.removeItem(ACCESS_TOKEN_KEY);
}
