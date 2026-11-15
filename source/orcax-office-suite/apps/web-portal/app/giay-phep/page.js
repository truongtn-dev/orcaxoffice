'use client';

import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';

export default function LicensesPage() {
  return (
    <PortalScreen>
      {(workspace) => <Licenses key={workspace.session.workspace.id} session={workspace.session} />}
    </PortalScreen>
  );
}

function Licenses({ session }) {
  const permissions = new Set(session.permissions ?? []);
  const canRead = permissions.has('license.read');
  if (!canRead) {
    return (
      <section className="panel">
        <p className="eyebrow">Giấy phép</p>
        <h1>Không có quyền xem giấy phép</h1>
      </section>
    );
  }
  return (
    <LicensesPanel
      canIssue={permissions.has('license.issue')}
      canManage={permissions.has('license.manage')}
      canActivate={permissions.has('license.activate')}
    />
  );
}

function LicensesPanel({ canIssue, canManage, canActivate }) {
  const [plans, setPlans] = useState([]);
  const [licenses, setLicenses] = useState([]);
  const [activations, setActivations] = useState([]);
  const [planId, setPlanId] = useState('');
  const [licenseId, setLicenseId] = useState('');
  const [deviceKey, setDeviceKey] = useState('');
  const [deviceLabel, setDeviceLabel] = useState('');
  const [entitlement, setEntitlement] = useState('word.core');
  const [checkResult, setCheckResult] = useState(null);
  const [error, setError] = useState('');
  const [note, setNote] = useState('');

  async function reload() {
    const [planResult, licenseResult, activationResult] = await Promise.all([
      api('/plans'),
      api('/licenses'),
      api('/license-activations'),
    ]);
    setPlans(planResult.plans ?? []);
    setLicenses(licenseResult.licenses ?? []);
    setActivations(activationResult.activations ?? []);
    setPlanId((current) => current || planResult.plans?.[0]?.id || '');
    setLicenseId((current) => current || licenseResult.licenses?.[0]?.id || '');
  }

  useEffect(() => {
    let ignore = false;
    Promise.all([api('/plans'), api('/licenses'), api('/license-activations')])
      .then(([planResult, licenseResult, activationResult]) => {
        if (ignore) {
          return;
        }
        setPlans(planResult.plans ?? []);
        setLicenses(licenseResult.licenses ?? []);
        setActivations(activationResult.activations ?? []);
        setPlanId(planResult.plans?.[0]?.id || '');
        setLicenseId(licenseResult.licenses?.[0]?.id || '');
        setError('');
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

  async function issue(event) {
    event.preventDefault();
    setError('');
    setNote('');
    try {
      await api('/licenses', { method: 'POST', body: { planId } });
      setNote('Đã phát hành giấy phép cho không gian hiện tại.');
      await reload();
    } catch (reason) {
      setError(reason.message);
    }
  }

  async function activate(event) {
    event.preventDefault();
    setError('');
    setNote('');
    try {
      await api('/licenses/activate', {
        method: 'POST',
        body: { licenseId, deviceKey, deviceLabel },
      });
      setNote('Đã kích hoạt thiết bị.');
      setDeviceKey('');
      await reload();
    } catch (reason) {
      setError(reason.message);
    }
  }

  async function check(event) {
    event.preventDefault();
    setError('');
    try {
      const result = await api(`/entitlements/check?code=${encodeURIComponent(entitlement)}`);
      setCheckResult(result);
    } catch (reason) {
      setError(reason.message);
    }
  }

  return (
    <section className="panel">
      <p className="eyebrow">Giấy phép</p>
      <h1>Giấy phép không gian hiện tại</h1>
      <p className="lead">
        Giấy phép gắn theo workspace đang hoạt động. Không dùng chung giữa cá nhân và tổ chức.
      </p>
      {error ? <p className="form-error">{error}</p> : null}
      {note ? <p className="form-note">{note}</p> : null}

      {canIssue ? (
        <form className="stack-form" onSubmit={issue}>
          <label>
            Gói phát hành
            <select value={planId} onChange={(event) => setPlanId(event.target.value)} required>
              {plans.map((plan) => (
                <option key={plan.id} value={plan.id}>
                  {plan.name} · {plan.licenseType === 'perpetual' ? 'Vĩnh viễn' : 'Thuê bao'} ·{' '}
                  {plan.deviceLimit} thiết bị
                </option>
              ))}
            </select>
          </label>
          <button type="submit" className="button">
            Phát hành
          </button>
        </form>
      ) : null}

      <h2>Giấy phép</h2>
      <ul className="record-list">
        {licenses.map((item) => (
          <li key={item.id}>
            <div>
              <strong>{item.plan.name}</strong>
              <span>
                {statusLabel(item.status)} · {item.activationCount}/{item.plan.deviceLimit} thiết bị
                {item.endsAt
                  ? ` · hết hạn ${new Date(item.endsAt).toLocaleDateString('vi-VN')}`
                  : ''}
              </span>
            </div>
            {canManage ? (
              <div className="form-actions">
                {item.plan.licenseType !== 'perpetual' ? (
                  <button
                    type="button"
                    className="button secondary"
                    onClick={async () => {
                      setError('');
                      try {
                        await api(`/licenses/${item.id}/renew`, { method: 'POST' });
                        setNote('Đã gia hạn giấy phép.');
                        await reload();
                      } catch (reason) {
                        setError(reason.message);
                      }
                    }}
                  >
                    Gia hạn
                  </button>
                ) : null}
                {item.status !== 'revoked' ? (
                  <button
                    type="button"
                    className="button secondary"
                    onClick={async () => {
                      setError('');
                      try {
                        await api(`/licenses/${item.id}/revoke`, { method: 'POST' });
                        setNote('Đã thu hồi giấy phép.');
                        await reload();
                      } catch (reason) {
                        setError(reason.message);
                      }
                    }}
                  >
                    Thu hồi
                  </button>
                ) : null}
              </div>
            ) : null}
          </li>
        ))}
      </ul>

      {canActivate ? (
        <form className="stack-form" onSubmit={activate}>
          <h2>Kích hoạt thiết bị</h2>
          <label>
            Giấy phép
            <select
              value={licenseId}
              onChange={(event) => setLicenseId(event.target.value)}
              required
            >
              {licenses
                .filter((item) => item.usable)
                .map((item) => (
                  <option key={item.id} value={item.id}>
                    {item.plan.name}
                  </option>
                ))}
            </select>
          </label>
          <label>
            Khóa thiết bị
            <input
              value={deviceKey}
              onChange={(event) => setDeviceKey(event.target.value)}
              minLength={8}
              required
            />
          </label>
          <label>
            Tên thiết bị
            <input
              value={deviceLabel}
              onChange={(event) => setDeviceLabel(event.target.value)}
              required
            />
          </label>
          <button type="submit" className="button">
            Kích hoạt
          </button>
        </form>
      ) : null}

      <h2>Thiết bị đang gắn</h2>
      <ul className="record-list">
        {activations.map((item) => (
          <li key={item.id}>
            <div>
              <strong>{item.deviceLabel}</strong>
              <span>
                {item.currentUser ? 'Thiết bị của bạn' : 'Thiết bị thành viên khác'} ·{' '}
                {new Date(item.lastSeenAt).toLocaleString('vi-VN')}
              </span>
            </div>
            {canActivate || canManage ? (
              <button
                type="button"
                className="button secondary"
                onClick={async () => {
                  setError('');
                  try {
                    await api(`/license-activations/${item.id}`, { method: 'DELETE' });
                    setNote('Đã thu hồi kích hoạt thiết bị.');
                    await reload();
                  } catch (reason) {
                    setError(reason.message);
                  }
                }}
              >
                Thu hồi thiết bị
              </button>
            ) : null}
          </li>
        ))}
      </ul>

      <form className="stack-form" onSubmit={check}>
        <h2>Kiểm tra quyền lợi</h2>
        <label>
          Mã quyền lợi
          <input value={entitlement} onChange={(event) => setEntitlement(event.target.value)} />
        </label>
        <button type="submit" className="button secondary">
          Kiểm tra
        </button>
      </form>
      {checkResult ? (
        <p className="form-note">
          {checkResult.allowed
            ? `Được phép (${checkResult.planCode}, trạng thái ${statusLabel(checkResult.status)}).`
            : 'Không có giấy phép hợp lệ cho quyền lợi này.'}
        </p>
      ) : null}
    </section>
  );
}

function statusLabel(status) {
  if (status === 'active') return 'Đang hiệu lực';
  if (status === 'grace') return 'Gia hạn tạm';
  if (status === 'expired') return 'Hết hạn';
  if (status === 'revoked') return 'Đã thu hồi';
  if (status === 'pending') return 'Chưa hiệu lực';
  return status;
}
