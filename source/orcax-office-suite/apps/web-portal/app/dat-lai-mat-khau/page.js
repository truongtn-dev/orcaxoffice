'use client';

import Link from 'next/link';
import { useRouter, useSearchParams } from 'next/navigation';
import { Suspense, useState } from 'react';
import { AuthFrame } from '../../components/AuthFrame';
import { api } from '../../lib/api';

export default function ResetPasswordPage() {
  return (
    <Suspense fallback={<AuthFrame title="Đặt lại mật khẩu" />}>
      <ResetForm />
    </Suspense>
  );
}

function ResetForm() {
  const params = useSearchParams();
  const router = useRouter();
  const [email, setEmail] = useState(params.get('email') ?? '');
  const [code, setCode] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [pending, setPending] = useState(false);

  async function submit(event) {
    event.preventDefault();
    if (!/^\d{6}$/.test(code) || password.length < 10 || password.length > 72) {
      setError('Mã gồm 6 chữ số. Mật khẩu mới từ 10 đến 72 ký tự.');
      return;
    }
    setPending(true);
    setError('');
    try {
      await api('/auth/password/reset', {
        method: 'POST',
        body: { email, code, password },
        skipRefresh: true,
        token: '',
      });
      router.push('/dang-nhap');
    } catch (reason) {
      setError(reason.message);
      setPending(false);
    }
  }

  return (
    <AuthFrame
      title="Đặt lại mật khẩu"
      lead="Sau khi đặt lại, mọi phiên đăng nhập hiện có bị thu hồi."
    >
      <form className="stack-form" onSubmit={submit}>
        <label>
          Email
          <input
            type="email"
            value={email}
            onChange={(event) => setEmail(event.target.value)}
            required
          />
        </label>
        <label>
          Mã
          <input
            inputMode="numeric"
            value={code}
            onChange={(event) => setCode(event.target.value)}
            required
          />
        </label>
        <label>
          Mật khẩu mới
          <input
            type="password"
            autoComplete="new-password"
            value={password}
            onChange={(event) => setPassword(event.target.value)}
            required
          />
        </label>
        {error ? <p className="form-error">{error}</p> : null}
        <button type="submit" className="button" disabled={pending}>
          {pending ? 'Đang lưu' : 'Lưu mật khẩu'}
        </button>
      </form>
      <p className="auth-links">
        <Link href="/dang-nhap">Về đăng nhập</Link>
      </p>
    </AuthFrame>
  );
}
