'use client';

import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';

export default function AuditPage() {
  return (
    <PortalScreen>
      {(workspace) => <AuditLog key={workspace.session.workspace.id} session={workspace.session} />}
    </PortalScreen>
  );
}

function AuditLog({ session }) {
  const canRead = (session.permissions ?? []).includes('workspace.audit.read');
  if (!canRead) {
    return (
      <section className="panel">
        <p className="eyebrow">Nhật ký</p>
        <h1>Không có quyền xem nhật ký</h1>
        <p className="lead">
          Chỉ chủ sở hữu, quản trị hoặc kiểm toán của không gian hiện tại mới xem được nhật ký kiểm
          toán.
        </p>
      </section>
    );
  }
  return <AuditLogPanel />;
}

function AuditLogPanel() {
  const [action, setAction] = useState('');
  const [events, setEvents] = useState([]);
  const [error, setError] = useState('');
  const [ready, setReady] = useState(false);

  async function load(filters = {}) {
    setError('');
    try {
      const query = new URLSearchParams();
      if (filters.action) {
        query.set('action', filters.action);
      }
      query.set('limit', '50');
      const result = await api(`/audit-events?${query}`);
      setEvents(result.events ?? []);
    } catch (reason) {
      setError(reason.message);
      setEvents([]);
    } finally {
      setReady(true);
    }
  }

  useEffect(() => {
    let ignore = false;
    api('/audit-events?limit=50')
      .then((result) => {
        if (!ignore) {
          setEvents(result.events ?? []);
          setError('');
        }
      })
      .catch((reason) => {
        if (!ignore) {
          setError(reason.message);
          setEvents([]);
        }
      })
      .finally(() => {
        if (!ignore) {
          setReady(true);
        }
      });
    return () => {
      ignore = true;
    };
  }, []);

  return (
    <section className="panel">
      <p className="eyebrow">Nhật ký</p>
      <h1>Nhật ký kiểm toán</h1>
      <p className="lead">
        Chỉ hiện sự kiện của không gian đang hoạt động. Việc tìm kiếm cũng được ghi nhật ký.
      </p>
      <form
        className="stack-form"
        onSubmit={(event) => {
          event.preventDefault();
          load({ action: action.trim() });
        }}
      >
        <label>
          Lọc theo hành động
          <input
            value={action}
            onChange={(event) => setAction(event.target.value)}
            placeholder="ví dụ membership.role.changed"
          />
        </label>
        <button type="submit" className="button">
          Tìm
        </button>
      </form>
      {error ? <p className="form-error">{error}</p> : null}
      {!ready ? <p className="lead">Đang tải…</p> : null}
      <ul className="record-list">
        {events.map((item) => (
          <li key={item.id}>
            <div>
              <strong>{item.action}</strong>
              <span>
                {item.outcome === 'success' ? 'Thành công' : 'Thất bại'} · {item.targetType}
                {item.targetId ? ` · ${item.targetId}` : ''} ·{' '}
                {new Date(item.occurredAt).toLocaleString('vi-VN')}
              </span>
            </div>
          </li>
        ))}
      </ul>
      {ready && events.length === 0 ? <p className="form-note">Chưa có sự kiện phù hợp.</p> : null}
    </section>
  );
}
