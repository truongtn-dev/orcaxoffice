'use client';

import { useState } from 'react';
import { API_VERSION } from '@orcax/api-contracts';
import { PortalScreen } from '../components/PortalScreen';
import { useSwitchRequest } from '../components/AppShell';
import { api } from '../lib/api';
import { kindLabel, roleLabel } from '../lib/labels';

export default function HomePage() {
  return (
    <PortalScreen allowAnonymous>
      {(workspace) => (workspace.session ? <SignedIn workspace={workspace} /> : <SignedOut />)}
    </PortalScreen>
  );
}

function SignedOut() {
  return (
    <section className="panel">
      <p className="eyebrow">Cổng thông tin</p>
      <h1>Bàn làm việc OrcaX Office</h1>
      <p className="lead">
        Đăng nhập để vào không gian cá nhân. Không gian tổ chức chỉ mở sau khi bạn xác nhận chuyển
        workspace.
      </p>
      <dl className="facts">
        <div>
          <dt>Hợp đồng API</dt>
          <dd>{API_VERSION}</dd>
        </div>
        <div>
          <dt>Kiến trúc</dt>
          <dd>Modular monolith</dd>
        </div>
        <div>
          <dt>Cơ sở dữ liệu</dt>
          <dd>PostgreSQL</dd>
        </div>
        <div>
          <dt>Sức khỏe dịch vụ</dt>
          <dd>/health và /health/ready</dd>
        </div>
      </dl>
      <aside className="callout">
        <strong>Tài liệu Office ở lại trên máy người dùng.</strong>
        <span>Cổng thông tin này không nhận nội dung Word hoặc Excel.</span>
      </aside>
    </section>
  );
}

function SignedIn({ workspace }) {
  const askSwitch = useSwitchRequest();
  const { session, reload } = workspace;
  const [displayName, setDisplayName] = useState(session.user.displayName);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const suggestion = session.suggestedWorkspace;

  async function save(event) {
    event.preventDefault();
    setMessage('');
    setError('');
    try {
      const next = await api('/auth/me', { method: 'PATCH', body: { displayName } });
      setMessage('Đã lưu tên hiển thị.');
      await reload();
      if (next?.user?.displayName) {
        setDisplayName(next.user.displayName);
      }
    } catch (reason) {
      setError(reason.message);
    }
  }

  return (
    <section className="panel">
      <p className="eyebrow">{kindLabel(session.workspace.kind)}</p>
      <h1>{session.workspace.name}</h1>
      <p className="lead">
        {session.user.email} · {roleLabel(session.workspace.role)}
      </p>
      {suggestion ? (
        <aside className="callout">
          <strong>Gợi ý: {suggestion.name}</strong>
          <span>Bạn vẫn đang ở {session.workspace.name}. Gợi ý này không tự chuyển workspace.</span>
          <button
            type="button"
            className="button secondary"
            onClick={() => askSwitch?.(suggestion.id)}
          >
            Xem hướng chuyển
          </button>
        </aside>
      ) : null}
      <form className="stack-form" onSubmit={save}>
        <label>
          Tên hiển thị
          <input
            value={displayName}
            onChange={(event) => setDisplayName(event.target.value)}
            maxLength={120}
            required
          />
        </label>
        {error ? <p className="form-error">{error}</p> : null}
        {message ? <p className="form-note">{message}</p> : null}
        <button type="submit" className="button">
          Lưu tên
        </button>
      </form>
    </section>
  );
}
