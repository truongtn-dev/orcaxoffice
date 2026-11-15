'use client';

import { useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';

export default function InvitationPage() {
  return <PortalScreen>{() => <AcceptInvitation />}</PortalScreen>;
}

function AcceptInvitation() {
  const [token, setToken] = useState('');
  const [error, setError] = useState('');
  const [note, setNote] = useState('');

  async function submit(event) {
    event.preventDefault();
    if (token.trim().length < 20) {
      setError('Hãy dán mã lời mời.');
      return;
    }
    setError('');
    setNote('');
    try {
      const result = await api('/invitations/accept', {
        method: 'POST',
        body: { token: token.trim() },
      });
      setNote(
        `Đã tham gia với vai trò ${result.role}. Workspace hiện tại không đổi. Hãy xác nhận nếu bạn muốn chuyển.`,
      );
    } catch (reason) {
      setError(reason.message);
    }
  }

  return (
    <section className="panel">
      <p className="eyebrow">Lời mời</p>
      <h1>Nhận lời mời</h1>
      <p className="lead">Mã lời mời phải trùng email của tài khoản đang đăng nhập.</p>
      <form className="stack-form" onSubmit={submit}>
        <label>
          Mã lời mời
          <textarea
            value={token}
            onChange={(event) => setToken(event.target.value)}
            rows={4}
            required
          />
        </label>
        {error ? <p className="form-error">{error}</p> : null}
        {note ? <p className="form-note">{note}</p> : null}
        <button type="submit" className="button">
          Tham gia
        </button>
      </form>
    </section>
  );
}
