'use client';

import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';

export default function ProfilesPage() {
  return (
    <PortalScreen>
      {(workspace) => <Profiles key={workspace.session.workspace.id} session={workspace.session} />}
    </PortalScreen>
  );
}

function Profiles({ session }) {
  const permissions = new Set(session.permissions ?? []);
  if (!permissions.has('profile.read')) {
    return (
      <section className="panel">
        <p className="eyebrow">Hồ sơ</p>
        <h1>Không có quyền xem hồ sơ</h1>
      </section>
    );
  }
  return (
    <ProfilesPanel
      canWrite={permissions.has('profile.write')}
      canReview={permissions.has('profile.review')}
      canPublish={permissions.has('profile.publish')}
    />
  );
}

function ProfilesPanel({ canWrite, canReview, canPublish }) {
  const [profiles, setProfiles] = useState([]);
  const [detail, setDetail] = useState(null);
  const [departments, setDepartments] = useState([]);
  const [groups, setGroups] = useState([]);
  const [distributeTargetType, setDistributeTargetType] = useState('department');
  const [distributeTargetId, setDistributeTargetId] = useState('');
  const [code, setCode] = useState('');
  const [name, setName] = useState('');
  const [host, setHost] = useState('word');
  const [source, setSource] = useState('');
  const [citation, setCitation] = useState('');
  const [evidenceSummary, setEvidenceSummary] = useState('');
  const [error, setError] = useState('');
  const [note, setNote] = useState('');
  const [report, setReport] = useState(null);

  async function reloadList() {
    const result = await api('/profiles');
    setProfiles(result.profiles ?? []);
    return result.profiles ?? [];
  }

  async function loadDetail(id) {
    if (!id) {
      setDetail(null);
      return;
    }
    const result = await api(`/profiles/${id}`);
    setDetail(result);
  }

  useEffect(() => {
    let ignore = false;
    Promise.all([
      api('/profiles'),
      api('/departments').catch(() => ({ departments: [] })),
      api('/groups').catch(() => ({ groups: [] })),
    ])
      .then(([result, deptResult, groupResult]) => {
        if (ignore) {
          return;
        }
        const rows = result.profiles ?? [];
        setProfiles(rows);
        setDepartments(deptResult.departments ?? []);
        setGroups(groupResult.groups ?? []);
        setDistributeTargetId(deptResult.departments?.[0]?.id || groupResult.groups?.[0]?.id || '');
        const first = rows[0]?.id || '';
        if (first) {
          return api(`/profiles/${first}`).then((detailResult) => {
            if (!ignore) {
              setDetail(detailResult);
            }
          });
        }
        setDetail(null);
        return undefined;
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

  async function createProfile(event) {
    event.preventDefault();
    setError('');
    setNote('');
    try {
      const created = await api('/profiles', {
        method: 'POST',
        body: { code, name, host, source },
      });
      setNote('Đã tạo hồ sơ nháp.');
      setCode('');
      setName('');
      setSource('');
      const rows = await reloadList();
      if (!rows.find((item) => item.id === created.id)) {
        setProfiles((current) => [created, ...current]);
      }
      await loadDetail(created.id);
    } catch (reason) {
      setError(reason.message);
    }
  }

  return (
    <section className="panel">
      <p className="eyebrow">Hồ sơ & quy tắc</p>
      <h1>Hồ sơ Word / Excel</h1>
      <p className="lead">
        Hồ sơ hệ thống dùng chung; hồ sơ cá nhân hoặc tổ chức gắn workspace hiện tại. Duyệt bị chặn
        khi còn trích dẫn pháp lý chưa xác minh.
      </p>
      {error ? <p className="form-error">{error}</p> : null}
      {note ? <p className="form-note">{note}</p> : null}

      {canWrite ? (
        <form className="stack-form" onSubmit={createProfile}>
          <h2>Tạo hồ sơ</h2>
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
          <label>
            Nguồn
            <input value={source} onChange={(event) => setSource(event.target.value)} required />
          </label>
          <button type="submit" className="button">
            Tạo nháp
          </button>
        </form>
      ) : null}

      <h2>Danh mục</h2>
      <ul className="record-list">
        {profiles.map((item) => (
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
              {item.host} · {scopeLabel(item.scope)} · {statusLabel(item.status)} · v
              {item.currentVersion}
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
          <p>
            Trạng thái: {statusLabel(detail.status)} · Nguồn: {detail.version.source}
          </p>
          <h3>Quy tắc</h3>
          <ul className="record-list">
            {detail.rules.map((rule) => (
              <li key={rule.id}>
                <strong>
                  {rule.code}@{rule.version}
                </strong>
                <span>
                  {rule.title} · {rule.defaultSeverity}
                </span>
              </li>
            ))}
          </ul>
          <h3>Trích dẫn pháp lý</h3>
          <ul className="record-list">
            {detail.legalReferences.map((item) => (
              <li key={item.id}>
                <div>
                  <strong>{item.citation}</strong>
                  <span>{item.verified ? 'Đã xác minh' : 'Chưa xác minh'}</span>
                </div>
                {canReview && !item.verified ? (
                  <button
                    type="button"
                    className="button secondary"
                    onClick={async () => {
                      setError('');
                      try {
                        await api(`/legal-references/${item.id}/verify`, {
                          method: 'POST',
                          body: {
                            evidenceType: 'statement',
                            summary: evidenceSummary || 'Xác minh thủ công từ cổng thông tin.',
                          },
                        });
                        setNote('Đã ghi bằng chứng xác minh.');
                        await loadDetail(detail.id);
                      } catch (reason) {
                        setError(reason.message);
                      }
                    }}
                  >
                    Xác minh
                  </button>
                ) : null}
              </li>
            ))}
          </ul>
          {canWrite && detail.scope !== 'system' ? (
            <form
              className="stack-form"
              onSubmit={async (event) => {
                event.preventDefault();
                setError('');
                try {
                  await api(`/profiles/${detail.id}/legal-references`, {
                    method: 'POST',
                    body: { citation },
                  });
                  setCitation('');
                  setNote('Đã thêm trích dẫn.');
                  await loadDetail(detail.id);
                } catch (reason) {
                  setError(reason.message);
                }
              }}
            >
              <label>
                Thêm trích dẫn
                <input
                  value={citation}
                  onChange={(event) => setCitation(event.target.value)}
                  required
                />
              </label>
              <button type="submit" className="button secondary">
                Thêm
              </button>
            </form>
          ) : null}
          {canReview ? (
            <label>
              Tóm tắt bằng chứng xác minh
              <input
                value={evidenceSummary}
                onChange={(event) => setEvidenceSummary(event.target.value)}
              />
            </label>
          ) : null}
          <div className="form-actions">
            {canWrite && detail.scope !== 'system' && detail.status === 'draft' ? (
              <button
                type="button"
                className="button"
                onClick={async () => {
                  setError('');
                  try {
                    await api(`/profiles/${detail.id}/submit-review`, { method: 'POST', body: {} });
                    setNote('Đã gửi duyệt.');
                    await reloadList();
                    await loadDetail(detail.id);
                  } catch (reason) {
                    setError(reason.message);
                  }
                }}
              >
                Gửi duyệt
              </button>
            ) : null}
            {canPublish && detail.scope !== 'system' && detail.status === 'in_review' ? (
              <button
                type="button"
                className="button"
                onClick={async () => {
                  setError('');
                  try {
                    await api(`/profiles/${detail.id}/approve`, { method: 'POST', body: {} });
                    setNote('Đã duyệt hồ sơ.');
                    await reloadList();
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
                className="button secondary"
                onClick={async () => {
                  setError('');
                  try {
                    await api(`/profiles/${detail.id}/deprecate`, { method: 'POST', body: {} });
                    setNote('Đã ngừng hồ sơ.');
                    await reloadList();
                    await loadDetail(detail.id);
                  } catch (reason) {
                    setError(reason.message);
                  }
                }}
              >
                Ngừng dùng
              </button>
            ) : null}
            {canPublish &&
            detail.scope === 'organization' &&
            detail.status === 'approved' &&
            (departments.length > 0 || groups.length > 0) ? (
              <form
                className="stack-form"
                onSubmit={async (event) => {
                  event.preventDefault();
                  setError('');
                  try {
                    await api(`/profiles/${detail.id}/distribute`, {
                      method: 'POST',
                      body: {
                        targetType: distributeTargetType,
                        targetId: distributeTargetId,
                      },
                    });
                    setNote('Đã phân phối hồ sơ tới phòng ban/nhóm.');
                  } catch (reason) {
                    setError(reason.message);
                  }
                }}
              >
                <h3>Phân phối</h3>
                <label>
                  Loại đích
                  <select
                    value={distributeTargetType}
                    onChange={(event) => {
                      const next = event.target.value;
                      setDistributeTargetType(next);
                      setDistributeTargetId(
                        next === 'department' ? departments[0]?.id || '' : groups[0]?.id || '',
                      );
                    }}
                  >
                    <option value="department">Phòng ban</option>
                    <option value="group">Nhóm</option>
                  </select>
                </label>
                <label>
                  Đích
                  <select
                    value={distributeTargetId}
                    onChange={(event) => setDistributeTargetId(event.target.value)}
                  >
                    {(distributeTargetType === 'department' ? departments : groups).map((item) => (
                      <option key={item.id} value={item.id}>
                        {item.name}
                      </option>
                    ))}
                  </select>
                </label>
                <button type="submit" className="button secondary">
                  Phân phối
                </button>
              </form>
            ) : null}
            {detail.status === 'approved' ? (
              <button
                type="button"
                className="button secondary"
                onClick={async () => {
                  setError('');
                  try {
                    const result = await api(`/profiles/${detail.id}/evaluate`, {
                      method: 'POST',
                      body: {
                        results: detail.rules.map((rule) => ({
                          ruleCode: rule.code,
                          ruleVersion: rule.version,
                          passed: true,
                        })),
                      },
                    });
                    setReport(result);
                    setNote('Đã tạo báo cáo đánh giá mẫu (mọi rule pass).');
                  } catch (reason) {
                    setError(reason.message);
                  }
                }}
              >
                Chạy báo cáo mẫu
              </button>
            ) : null}
          </div>
          {report ? (
            <p className="form-note">
              Báo cáo v{report.profileVersion}: pass {report.summary.pass}, fail{' '}
              {report.summary.fail}, warning {report.summary.warning}.
            </p>
          ) : null}
        </div>
      ) : null}
    </section>
  );
}

function statusLabel(status) {
  if (status === 'draft') return 'Nháp';
  if (status === 'in_review') return 'Chờ duyệt';
  if (status === 'approved') return 'Đã duyệt';
  if (status === 'deprecated') return 'Ngừng dùng';
  return status;
}

function scopeLabel(scope) {
  if (scope === 'system') return 'Hệ thống';
  if (scope === 'personal') return 'Cá nhân';
  if (scope === 'organization') return 'Tổ chức';
  return scope;
}
