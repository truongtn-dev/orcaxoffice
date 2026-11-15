'use client';

import Link from 'next/link';
import { useRouter } from 'next/navigation';
import { useState } from 'react';
import { AuthFrame } from '../../components/AuthFrame';
import { api } from '../../lib/api';

export default function RegisterPage() {
  const router = useRouter();
  const [displayName, setDisplayName] = useState('');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [pending, setPending] = useState(false);

  async function submit(event) {
    event.preventDefault();
    if (displayName.trim().length < 1 || password.length < 10 || password.length > 72) {
      setError('Tên hiển thị không để trống. Mật khẩu từ 10 đến 72 ký tự.');
      return;
    }
    setPending(true);
    setError('');
    try {
      await api('/auth/register', {
        method: 'POST',
        body: { displayName, email, password },
        skipRefresh: true,
        token: '',
      });
      router.push(`/xac-minh-email?email=${encodeURIComponent(email)}`);
    } catch (reason) {
      setError(reason.message);
      setPending(false);
    }
  }

  return (
    <AuthFrame
      title="Đăng ký"
      lead="Tài khoản mới nhận một không gian cá nhân. Email cần được xác minh trước khi đăng nhập."
    >
      <form className="stack-form" onSubmit={submit}>
        <label>
          Tên hiển thị
          <input
            value={displayName}
            onChange={(event) => setDisplayName(event.target.value)}
            required
            maxLength={120}
          />
        </label>
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
            autoComplete="new-password"
            value={password}
            onChange={(event) => setPassword(event.target.value)}
            minLength={10}
            maxLength={72}
            required
          />
        </label>
        {error ? <p className="form-error">{error}</p> : null}
        <button type="submit" className="button" disabled={pending}>
          {pending ? 'Đang tạo tài khoản' : 'Tạo tài khoản'}
        </button>
      </form>
      <p className="auth-links">
        <Link href="/dang-nhap">Đã có tài khoản</Link>
      </p>
    </AuthFrame>
  );
}
