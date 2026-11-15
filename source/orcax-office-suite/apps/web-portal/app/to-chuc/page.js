'use client';

import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';

export default function OrganizationPage() {
  return (
    <PortalScreen>
      {(workspace) => <Organization key={workspace.session.workspace.id} workspace={workspace} />}
    </PortalScreen>
  );
}

function Organization({ workspace }) {
  const { session, reload } = workspace;
  const [name, setName] = useState('');
  const [rename, setRename] = useState('');
  const [current, setCurrent] = useState(null);
  const [error, setError] = useState('');
  const [note, setNote] = useState('');
  const canRename =
    session.workspace.kind === 'organization' && session.permissions.includes('workspace.update');
  const canLifecycle =
    session.workspace.kind === 'organization' &&
    session.workspace.role === 'organization_owner' &&
    session.permissions.includes('workspace.update');
  const canCreate = session.workspace.kind === 'personal';

  useEffect(() => {
    if (session.workspace.kind !== 'organization') {
      return;
    }
    let ignore = false;
    api('/organizations/current')
      .then((organization) => {
        if (ignore) {
          return;
        }
        setCurrent(organization);
        setRename(organization.name);
      })
      .catch((reason) => {
        if (!ignore) {
          setError(reason.message);
        }
      });
    return () => {
      ignore = true;
    };
  }, [session.workspace.kind]);

  async function create(event) {
    event.preventDefault();
    if (name.trim().length < 2) {
      setError('Tên tổ chức cần từ 2 ký tự.');
      return;
    }
    setError('');
    setNote('');
    try {
      const created = await api('/organizations', { method: 'POST', body: { name } });
      setNote(
        `Đã tạo ${created.name}. Bạn vẫn ở không gian hiện tại. Hãy dùng nút workspace và xác nhận khi muốn chuyển.`,
      );
      setName('');
      await reload();
    } catch (reason) {
      setError(reason.message);
    }
  }

  async function saveName(event) {
    event.preventDefault();
    setError('');
    setNote('');
    try {
      const updated = await api('/organizations/current', {
        method: 'PATCH',
        body: { name: rename },
      });
      setNote(`Đã đổi tên thành ${updated.name}.`);
      await reload();
    } catch (reason) {
      setError(reason.message);
    }
  }

  return (
    <section className="panel">
      <p className="eyebrow">Tổ chức</p>
      <h1>{current ? current.name : 'Tạo tổ chức'}</h1>
      <p className="lead">
        Tổ chức chỉ được tạo khi bạn đang ở không gian cá nhân. Dữ liệu cá nhân và dữ liệu tổ chức
        nằm ở hai workspace khác nhau; tạo xong không tự chuyển sang tổ chức.
      </p>
      {error ? <p className="form-error">{error}</p> : null}
      {note ? <p className="form-note">{note}</p> : null}
      {canRename ? (
        <form className="stack-form" onSubmit={saveName}>
          <label>
            Tên tổ chức hiện tại
            <input
              value={rename}
              onChange={(event) => setRename(event.target.value)}
              minLength={2}
              maxLength={160}
              required
            />
          </label>
          <button type="submit" className="button">
            Lưu tên tổ chức
          </button>
        </form>
      ) : null}
      {current ? (
        <p className="form-note">Trạng thái tổ chức: {statusLabel(current.status)}</p>
      ) : null}
      {canLifecycle && current ? (
        <div className="form-actions">
          {current.status !== 'suspended' ? (
            <button
              type="button"
              className="button secondary"
              onClick={async () => {
                setError('');
                try {
                  const updated = await api('/organizations/current/status', {
                    method: 'POST',
                    body: { status: 'suspended' },
                  });
                  setCurrent(updated);
                  setNote('Đã tạm ngưng tổ chức. Workspace tổ chức không còn switchable.');
                  await reload();
                } catch (reason) {
                  setError(reason.message);
                }
              }}
            >
              Tạm ngưng tổ chức
            </button>
          ) : (
            <button
              type="button"
              className="button"
              onClick={async () => {
                setError('');
                try {
                  const updated = await api('/organizations/current/status', {
                    method: 'POST',
                    body: { status: 'active' },
                  });
                  setCurrent(updated);
                  setNote('Đã kích hoạt lại tổ chức.');
                  await reload();
                } catch (reason) {
                  setError(reason.message);
                }
              }}
            >
              Kích hoạt lại
            </button>
          )}
          {current.status !== 'archived' ? (
            <button
              type="button"
              className="button secondary"
              onClick={async () => {
                if (!window.confirm('Lưu trữ tổ chức? Thao tác dành cho chủ sở hữu.')) {
                  return;
                }
                setError('');
                try {
                  const updated = await api('/organizations/current/status', {
                    method: 'POST',
                    body: { status: 'archived' },
                  });
                  setCurrent(updated);
                  setNote('Đã lưu trữ tổ chức.');
                  await reload();
                } catch (reason) {
                  setError(reason.message);
                }
              }}
            >
              Lưu trữ
            </button>
          ) : null}
        </div>
      ) : null}
      {canCreate ? (
        <form className="stack-form" onSubmit={create}>
          <label>
            Tên tổ chức mới
            <input
              value={name}
              onChange={(event) => setName(event.target.value)}
              minLength={2}
              maxLength={160}
              required
            />
          </label>
          <button type="submit" className="button">
            Tạo tổ chức
          </button>
        </form>
      ) : (
        <p className="form-note">
          Để tạo tổ chức khác, hãy chuyển về không gian cá nhân và xác nhận trước khi tạo.
        </p>
      )}
    </section>
  );
}

function statusLabel(status) {
  if (status === 'active') return 'Đang hoạt động';
  if (status === 'suspended') return 'Tạm ngưng';
  if (status === 'archived') return 'Đã lưu trữ';
  return status;
}
