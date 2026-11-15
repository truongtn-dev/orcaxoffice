'use client';

import Link from 'next/link';
import { useRouter } from 'next/navigation';
import { useState } from 'react';
import { AuthFrame } from '../../components/AuthFrame';
import { api, ApiRequestError } from '../../lib/api';
import { writeAccessToken } from '../../lib/session';

export default function LoginPage() {
  const router = useRouter();
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [pending, setPending] = useState(false);

  async function submit(event) {
    event.preventDefault();
    if (password.length < 1 || !email.includes('@')) {
      setError('Nhập email và mật khẩu.');
      return;
    }
    setPending(true);
    setError('');
    try {
      const result = await api('/auth/login', {
        method: 'POST',
        body: { email, password },
        skipRefresh: true,
        token: '',
      });
      writeAccessToken(result.accessToken);
      router.push('/');
    } catch (reason) {
      if (reason instanceof ApiRequestError && reason.code === 'EMAIL_NOT_VERIFIED') {
        router.push(`/xac-minh-email?email=${encodeURIComponent(email)}`);
        return;
      }
      setError(reason.message);
      setPending(false);
    }
  }

  return (
    <AuthFrame
      title="Đăng nhập"
      lead="Sau khi đăng nhập, không gian đang hoạt động là không gian cá nhân."
    >
      <form className="stack-form" onSubmit={submit}>
        <label>
          Email
          <input
            type="email"
            autoComplete="username"
            value={email}
            onChange={(event) => setEmail(event.target.value)}
            required
          />
        </label>
        <label>
          Mật khẩu
          <input
            type="password"
            autoComplete="current-password"
            value={password}
            onChange={(event) => setPassword(event.target.value)}
            required
          />
        </label>
        {error ? <p className="form-error">{error}</p> : null}
        <button type="submit" className="button" disabled={pending}>
          {pending ? 'Đang đăng nhập' : 'Đăng nhập'}
        </button>
      </form>
      <p className="auth-links">
        <Link href="/quen-mat-khau">Quên mật khẩu</Link>
        <Link href="/dang-ky">Tạo tài khoản</Link>
      </p>
    </AuthFrame>
  );
}
