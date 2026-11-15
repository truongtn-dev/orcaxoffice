'use client';

import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';

export default function TemplatesPage() {
  return (
    <PortalScreen>
      {(workspace) => (
        <Templates key={workspace.session.workspace.id} session={workspace.session} />
      )}
    </PortalScreen>
  );
}

function Templates({ session }) {
  const permissions = new Set(session.permissions ?? []);
  if (!permissions.has('template.read')) {
    return (
      <section className="panel">
        <p className="eyebrow">Mẫu</p>
        <h1>Không có quyền xem mẫu</h1>
      </section>
    );
  }
  return (
    <TemplatesPanel
      canWrite={permissions.has('template.write')}
      canPublish={permissions.has('template.publish')}
    />
  );
}

function TemplatesPanel({ canWrite, canPublish }) {
  const [templates, setTemplates] = useState([]);
  const [detail, setDetail] = useState(null);
  const [preview, setPreview] = useState(null);
  const [code, setCode] = useState('');
  const [name, setName] = useState('');
  const [host, setHost] = useState('word');
  const [error, setError] = useState('');
  const [note, setNote] = useState('');

  async function reload() {
    const result = await api('/templates');
    setTemplates(result.templates ?? []);
    return result.templates ?? [];
  }

  async function loadDetail(id) {
    const [detailResult, previewResult] = await Promise.all([
      api(`/templates/${id}`),
      api(`/templates/${id}/mapping-preview`),
    ]);
    setDetail(detailResult);
    setPreview(previewResult);
  }

  useEffect(() => {
    let ignore = false;
    api('/templates')
      .then(async (result) => {
        if (ignore) {
          return;
        }
        const rows = result.templates ?? [];
        setTemplates(rows);
        if (rows[0]?.id) {
          const [detailResult, previewResult] = await Promise.all([
            api(`/templates/${rows[0].id}`),
            api(`/templates/${rows[0].id}/mapping-preview`),
          ]);
          if (!ignore) {
            setDetail(detailResult);
            setPreview(previewResult);
          }
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
      <p className="eyebrow">Mẫu tài liệu</p>
      <h1>Kho mẫu Word / Excel</h1>
      <p className="lead">
        Chỉ mẫu đã duyệt hoặc đã phát hành dùng cho thao tác governed. Ánh xạ trường là
        deterministic.
      </p>
      {error ? <p className="form-error">{error}</p> : null}
      {note ? <p className="form-note">{note}</p> : null}

      {canWrite ? (
        <form
          className="stack-form"
          onSubmit={async (event) => {
            event.preventDefault();
            setError('');
            try {
              const created = await api('/templates', {
                method: 'POST',
                body: { code, name, host },
              });
              setNote('Đã tạo mẫu nháp.');
              setCode('');
              setName('');
              await reload();
              await loadDetail(created.id);
            } catch (reason) {
              setError(reason.message);
            }
          }}
        >
          <h2>Tạo mẫu</h2>
          <label>
            Mã
            <input value={code} onChange={(event) => setCode(event.target.value)} required />
          </label>
          <label>
            Tên
            <input value={name} onChange={(event) => setName(event.target.value)} required />
          </label>
          <label>
            Host
            <select value={host} onChange={(event) => setHost(event.target.value)}>
              <option value="word">Word</option>
              <option value="excel">Excel</option>
            </select>
          </label>
          <button type="submit" className="button">
            Tạo nháp
          </button>
        </form>
      ) : null}

      <h2>Danh mục</h2>
      <ul className="record-list">
        {templates.map((item) => (
          <li key={item.id}>
            <button
              type="button"
              className="button secondary"
              onClick={async () => {
                setError('');
                try {
                  await loadDetail(item.id);
                } catch (reason) {
                  setError(reason.message);
                }
              }}
            >
              {item.name}
            </button>
            <span>
              {item.host} · {statusLabel(item.status)} · v{item.currentVersion}
              {item.governedUse ? ' · dùng được' : ''}
            </span>
          </li>
        ))}
      </ul>

      {detail ? (
        <div className="stack-form">
          <h2>
            {detail.name}{' '}
            <small>
              ({detail.code} · v{detail.currentVersion})
            </small>
          </h2>
          <h3>Trường</h3>
          <ul className="record-list">
            {detail.fields.map((field) => (
              <li key={field.id}>
                <strong>{field.fieldKey}</strong>
                <span>
                  {field.label} · {field.fieldType}
                  {field.required ? ' · bắt buộc' : ''}
                </span>
              </li>
            ))}
          </ul>
          <h3>Ánh xạ</h3>
          <ul className="record-list">
            {(preview?.mappings ?? []).map((item) => (
              <li key={`${item.fieldKey}-${item.source}`}>
                <strong>{item.destination}</strong>
                <span>
                  {item.source} → {item.transform}
                </span>
              </li>
            ))}
          </ul>
          <div className="form-actions">
            {canWrite && detail.scope !== 'system' && detail.status === 'draft' ? (
              <button
                type="button"
                className="button"
                onClick={async () => {
                  setError('');
                  try {
                    await api(`/templates/${detail.id}/submit-review`, { method: 'POST' });
                    setNote('Đã gửi duyệt mẫu.');
                    await reload();
                    await loadDetail(detail.id);
                  } catch (reason) {
                    setError(reason.message);
                  }
                }}
              >
                Gửi duyệt
              </button>
            ) : null}
            {canPublish && detail.scope !== 'system' && detail.status === 'review' ? (
              <button
                type="button"
                className="button"
                onClick={async () => {
                  setError('');
                  try {
                    await api(`/templates/${detail.id}/approve`, { method: 'POST' });
                    setNote('Đã duyệt mẫu.');
                    await reload();
                    await loadDetail(detail.id);
                  } catch (reason) {
                    setError(reason.message);
                  }
                }}
              >
                Duyệt
              </button>
            ) : null}
            {canPublish && detail.scope !== 'system' && detail.status === 'approved' ? (
              <button
                type="button"
                className="button"
                onClick={async () => {
                  setError('');
                  try {
                    await api(`/templates/${detail.id}/publish`, { method: 'POST' });
                    setNote('Đã phát hành mẫu.');
                    await reload();
                    await loadDetail(detail.id);
                  } catch (reason) {
                    setError(reason.message);
                  }
                }}
              >
                Phát hành
              </button>
            ) : null}
          </div>
        </div>
      ) : null}
    </section>
  );
}

function statusLabel(status) {
  if (status === 'draft') return 'Nháp';
  if (status === 'review') return 'Chờ duyệt';
  if (status === 'approved') return 'Đã duyệt';
  if (status === 'published') return 'Đã phát hành';
  if (status === 'deprecated') return 'Ngừng dùng';
  return status;
}
