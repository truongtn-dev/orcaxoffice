'use client';

import Link from 'next/link';
import { useState } from 'react';
import { AuthFrame } from '../../components/AuthFrame';
import { api } from '../../lib/api';

export default function ForgotPasswordPage() {
  const [email, setEmail] = useState('');
  const [note, setNote] = useState('');
  const [error, setError] = useState('');

  async function submit(event) {
    event.preventDefault();
    if (!email.includes('@')) {
      setError('Nhập email.');
      return;
    }
    setError('');
    try {
      const result = await api('/auth/password/forgot', {
        method: 'POST',
        body: { email },
        skipRefresh: true,
        token: '',
      });
      setNote(result.message);
    } catch (reason) {
      setError(reason.message);
    }
  }

  return (
    <AuthFrame
      title="Quên mật khẩu"
      lead="Nếu tài khoản đã xác minh, một mã đặt lại sẽ được gửi. Trang này không cho biết email có tồn tại hay không."
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
        {error ? <p className="form-error">{error}</p> : null}
        {note ? <p className="form-note">{note}</p> : null}
        <button type="submit" className="button">
          Gửi mã
        </button>
      </form>
      <p className="auth-links">
        <Link href={`/dat-lai-mat-khau?email=${encodeURIComponent(email)}`}>Đã có mã</Link>
        <Link href="/dang-nhap">Về đăng nhập</Link>
      </p>
    </AuthFrame>
  );
}
