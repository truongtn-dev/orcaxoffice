'use client';

import Link from 'next/link';
import { useRouter, useSearchParams } from 'next/navigation';
import { Suspense, useState } from 'react';
import { AuthFrame } from '../../components/AuthFrame';
import { api } from '../../lib/api';

export default function VerifyEmailPage() {
  return (
    <Suspense fallback={<AuthFrame title="Xác minh email" />}>
      <VerifyForm />
    </Suspense>
  );
}

function VerifyForm() {
  const params = useSearchParams();
  const router = useRouter();
  const [email, setEmail] = useState(params.get('email') ?? '');
  const [code, setCode] = useState('');
  const [error, setError] = useState('');
  const [note, setNote] = useState('');
  const [mailbox, setMailbox] = useState([]);
  const [pending, setPending] = useState(false);

  async function submit(event) {
    event.preventDefault();
    if (!/^\d{6}$/.test(code)) {
      setError('Mã xác minh gồm 6 chữ số.');
      return;
    }
    setPending(true);
    setError('');
    try {
      await api('/auth/email/verify', {
        method: 'POST',
        body: { email, code },
        skipRefresh: true,
        token: '',
      });
      router.push('/dang-nhap');
    } catch (reason) {
      setError(reason.message);
      setPending(false);
    }
  }

  async function resend() {
    setError('');
    setNote('');
    try {
      const result = await api('/auth/email/resend', {
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

  async function loadMailbox() {
    setError('');
    try {
      const result = await api(`/dev/mailbox?email=${encodeURIComponent(email)}`, {
        skipRefresh: true,
        token: '',
      });
      setMailbox(result.messages ?? []);
    } catch (reason) {
      setError(reason.message);
    }
  }

  return (
    <AuthFrame
      title="Xác minh email"
      lead="Nhập mã 6 chữ số đã gửi tới email. Mã hết hạn sau 15 phút."
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
          Mã xác minh
          <input
            inputMode="numeric"
            autoComplete="one-time-code"
            value={code}
            onChange={(event) => setCode(event.target.value)}
            required
          />
        </label>
        {error ? <p className="form-error">{error}</p> : null}
        {note ? <p className="form-note">{note}</p> : null}
        <button type="submit" className="button" disabled={pending}>
          {pending ? 'Đang xác minh' : 'Xác minh'}
        </button>
      </form>
      <div className="form-actions">
        <button type="button" className="button secondary" onClick={resend}>
          Gửi lại mã
        </button>
        {process.env.NEXT_PUBLIC_EMAIL_DEV_MAILBOX === '1' ? (
          <button type="button" className="button secondary" onClick={loadMailbox}>
            Xem hộp thư phát triển
          </button>
        ) : null}
      </div>
      {mailbox.map((message) => (
        <pre key={message.createdAt} className="mailbox">
          {message.text}
        </pre>
      ))}
      <p className="auth-links">
        <Link href="/dang-nhap">Về đăng nhập</Link>
      </p>
    </AuthFrame>
  );
}
