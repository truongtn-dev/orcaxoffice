'use client';

import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';

const KEYS = [
  { key: 'portal.locale', label: 'Ngôn ngữ cổng', sample: { value: 'vi-VN' } },
  { key: 'portal.timezone', label: 'Múi giờ', sample: { value: 'Asia/Ho_Chi_Minh' } },
  {
    key: 'notifications.email.enabled',
    label: 'Email thông báo',
    sample: { enabled: true },
  },
];

export default function ConfigPage() {
  return (
    <PortalScreen>
      {(workspace) => (
        <ConfigPanel key={workspace.session.workspace.id} session={workspace.session} />
      )}
    </PortalScreen>
  );
}

function ConfigPanel({ session }) {
  const permissions = new Set(session.permissions ?? []);
  const canRead = permissions.has('workspace.config.read');
  const canManage = permissions.has('workspace.config.manage');
  const [settings, setSettings] = useState([]);
  const [key, setKey] = useState(KEYS[0].key);
  const [json, setJson] = useState(JSON.stringify(KEYS[0].sample, null, 2));
  const [history, setHistory] = useState([]);
  const [error, setError] = useState('');
  const [note, setNote] = useState('');

  async function reload() {
    const result = await api('/workspace-settings');
    setSettings(result.settings ?? []);
  }

  useEffect(() => {
    if (!canRead) {
      return;
    }
    let ignore = false;
    api('/workspace-settings')
      .then((result) => {
        if (!ignore) {
          setSettings(result.settings ?? []);
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
  }, [canRead]);

  if (!canRead) {
    return (
      <section className="panel">
        <p className="eyebrow">Cấu hình</p>
        <h1>Không có quyền xem cấu hình</h1>
      </section>
    );
  }

  return (
    <section className="panel">
      <p className="eyebrow">Cấu hình workspace</p>
      <h1>Thiết lập có phiên bản</h1>
      <p className="lead">
        Thay đổi được kiểm tra khóa cho phép, lưu lịch sử và ghi audit. Không ghi secret vào đây.
      </p>
      {error ? <p className="form-error">{error}</p> : null}
      {note ? <p className="form-note">{note}</p> : null}

      <h2>Giá trị hiện tại</h2>
      <ul className="record-list">
        {settings.map((item) => (
          <li key={item.key}>
            <strong>{item.key}</strong>
            <span>
              v{item.version} · {JSON.stringify(item.value)}
            </span>
          </li>
        ))}
      </ul>

      {canManage ? (
        <form
          className="stack-form"
          onSubmit={async (event) => {
            event.preventDefault();
            setError('');
            try {
              const value = JSON.parse(json);
              await api('/workspace-settings', { method: 'POST', body: { key, value } });
              setNote('Đã lưu cấu hình.');
              await reload();
              const hist = await api(`/workspace-settings/${encodeURIComponent(key)}/history`);
              setHistory(hist.history ?? []);
            } catch (reason) {
              setError(reason.message || 'JSON không hợp lệ.');
            }
          }}
        >
          <h2>Cập nhật</h2>
          <label>
            Khóa
            <select
              value={key}
              onChange={(event) => {
                const next = event.target.value;
                setKey(next);
                const preset = KEYS.find((item) => item.key === next);
                setJson(JSON.stringify(preset?.sample ?? {}, null, 2));
              }}
            >
              {KEYS.map((item) => (
                <option key={item.key} value={item.key}>
                  {item.label} ({item.key})
                </option>
              ))}
            </select>
          </label>
          <label>
            Giá trị JSON
            <textarea rows={5} value={json} onChange={(event) => setJson(event.target.value)} />
          </label>
          <button type="submit" className="button">
            Lưu
          </button>
        </form>
      ) : null}

      {history.length ? (
        <>
          <h2>Lịch sử gần đây</h2>
          <ul className="record-list">
            {history.map((item) => (
              <li key={`${item.version}-${item.changedAt}`}>
                <strong>v{item.version}</strong>
                <span>
                  {new Date(item.changedAt).toLocaleString('vi-VN')} · {JSON.stringify(item.value)}
                </span>
              </li>
            ))}
          </ul>
        </>
      ) : null}
    </section>
  );
}
