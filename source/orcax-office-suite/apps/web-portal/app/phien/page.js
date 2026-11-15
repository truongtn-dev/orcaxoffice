'use client';

import { useRouter } from 'next/navigation';
import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';
import { clearAccessToken } from '../../lib/session';

export default function SessionsPage() {
  return <PortalScreen>{() => <Sessions />}</PortalScreen>;
}

function Sessions() {
  const router = useRouter();
  const [sessions, setSessions] = useState([]);
  const [error, setError] = useState('');

  useEffect(() => {
    let ignore = false;
    api('/sessions')
      .then((result) => {
        if (!ignore) {
          setSessions(result.sessions ?? []);
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

  async function revoke(session) {
    setError('');
    try {
      await api(`/sessions/${session.id}`, { method: 'DELETE' });
      if (session.current) {
        clearAccessToken();
        router.push('/dang-nhap');
        return;
      }
      const result = await api('/sessions');
      setSessions(result.sessions ?? []);
    } catch (reason) {
      setError(reason.message);
    }
  }

  return (
    <section className="panel">
      <p className="eyebrow">Phiên đăng nhập</p>
      <h1>Các phiên của bạn</h1>
      {error ? <p className="form-error">{error}</p> : null}
      <ul className="record-list">
        {sessions.map((session) => (
          <li key={session.id}>
            <div>
              <strong>{session.current ? 'Phiên hiện tại' : 'Phiên khác'}</strong>
              <span>
                {session.userAgent || 'Không rõ thiết bị'} · hết hạn{' '}
                {new Date(session.expiresAt).toLocaleString('vi-VN')}
              </span>
            </div>
            <button type="button" className="button secondary" onClick={() => revoke(session)}>
              Thu hồi
            </button>
          </li>
        ))}
      </ul>
    </section>
  );
}
