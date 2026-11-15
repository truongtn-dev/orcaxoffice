'use client';

import Link from 'next/link';
import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';
import { kindLabel, roleLabel } from '../../lib/labels';

export default function AdminPage() {
  return (
    <PortalScreen>
      {(workspace) => <AdminHub key={workspace.session.workspace.id} session={workspace.session} />}
    </PortalScreen>
  );
}

function AdminHub({ session }) {
  const [overview, setOverview] = useState(null);
  const [error, setError] = useState('');

  useEffect(() => {
    let ignore = false;
    api('/admin/overview')
      .then((result) => {
        if (!ignore) {
          setOverview(result);
        }
      })
      .catch((reason) => {
        if (!ignore) {
          setError(reason.message);
        }
      });
    return () => {
      ignore = true;
    };
  }, []);

  return (
    <section className="panel">
      <p className="eyebrow">Quản trị</p>
      <h1>Bảng điều khiển không gian</h1>
      <p className="lead">
        Widget theo quyền của {roleLabel(session.workspace.role)} ·{' '}
        {kindLabel(session.workspace.kind)}: {session.workspace.name}. Không dùng cá nhân hóa AI.
      </p>
      {error ? <p className="form-error">{error}</p> : null}
      <div className="admin-grid">
        {(overview?.widgets ?? []).map((widget) => (
          <Link key={widget.id} className="admin-card" href={widget.href}>
            <strong>{widget.title}</strong>
            <ul>
              {widget.metrics.map((metric) => (
                <li key={metric.label}>
                  {metric.label}: {metric.value}
                </li>
              ))}
            </ul>
          </Link>
        ))}
      </div>
      {!overview && !error ? <p className="form-note">Đang tải…</p> : null}
    </section>
  );
}
