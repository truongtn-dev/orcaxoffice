'use client';

import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';
import { roleLabel } from '../../lib/labels';

export default function MembersPage() {
  return <PortalScreen>{(workspace) => <Members session={workspace.session} />}</PortalScreen>;
}

function Members({ session }) {
  const [members, setMembers] = useState([]);
  const [email, setEmail] = useState('');
  const [role, setRole] = useState('member');
  const [error, setError] = useState('');
  const [note, setNote] = useState('');
  const permissions = new Set(session.permissions ?? []);
  const canInvite =
    session.workspace.kind === 'organization' && permissions.has('workspace.member.invite');

  useEffect(() => {
    let ignore = false;
    api('/members')
      .then((result) => {
        if (!ignore) {
          setMembers(result.members ?? []);
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
  }, [session.workspace.id]);

  async function invite(event) {
    event.preventDefault();
    setError('');
    setNote('');
    try {
      await api('/invitations', { method: 'POST', body: { email, role } });
      setNote(
        'Đã gửi lời mời. Người được mời vẫn ở workspace hiện tại của họ cho đến khi tự chuyển.',
      );
      setEmail('');
    } catch (reason) {
      setError(reason.message);
    }
  }

  async function act(id, action) {
    setError('');
    try {
      await api(`/memberships/${id}/${action}`, { method: 'POST' });
      const result = await api('/members');
      setMembers(result.members ?? []);
    } catch (reason) {
      setError(reason.message);
    }
  }

  return (
    <section className="panel">
      <p className="eyebrow">Thành viên</p>
      <h1>{session.workspace.name}</h1>
      {error ? <p className="form-error">{error}</p> : null}
      {note ? <p className="form-note">{note}</p> : null}
      {canInvite ? (
        <form className="stack-form" onSubmit={invite}>
          <label>
            Email mời
            <input
              type="email"
              value={email}
              onChange={(event) => setEmail(event.target.value)}
              required
            />
          </label>
          <label>
            Vai trò
            <select value={role} onChange={(event) => setRole(event.target.value)}>
              <option value="member">Thành viên</option>
              <option value="auditor">Kiểm toán</option>
              {permissions.has('workspace.member.role') ? (
                <>
                  <option value="organization_administrator">Quản trị</option>
                  <option value="organization_owner">Chủ sở hữu</option>
                </>
              ) : null}
            </select>
          </label>
          <button type="submit" className="button">
            Gửi lời mời
          </button>
        </form>
      ) : (
        <p className="lead">Không gian này không mở lời mời từ vai trò hiện tại.</p>
      )}
      <ul className="record-list">
        {members.map((member) => (
          <li key={member.id}>
            <div>
              <strong>{member.displayName || member.email}</strong>
              <span>
                {member.email} · {roleLabel(member.role)} ·{' '}
                {member.status === 'active' ? 'Đang tham gia' : 'Tạm ngưng'}
              </span>
            </div>
            {member.userId !== session.user.id && member.status === 'active' ? (
              <div className="form-actions">
                {permissions.has('workspace.member.role') ? (
                  <select
                    aria-label={`Đổi vai trò ${member.email}`}
                    defaultValue={member.role}
                    onChange={async (event) => {
                      setError('');
                      try {
                        await api(`/memberships/${member.id}/role`, {
                          method: 'POST',
                          body: { role: event.target.value },
                        });
                        const result = await api('/members');
                        setMembers(result.members ?? []);
                        setNote('Đã đổi vai trò.');
                      } catch (reason) {
                        setError(reason.message);
                      }
                    }}
                  >
                    <option value="member">Thành viên</option>
                    <option value="auditor">Kiểm toán</option>
                    {session.workspace.role === 'organization_owner' ? (
                      <>
                        <option value="organization_administrator">Quản trị</option>
                        <option value="organization_owner">Chủ sở hữu</option>
                      </>
                    ) : null}
                  </select>
                ) : null}
                {permissions.has('workspace.member.suspend') ? (
                  <button
                    type="button"
                    className="button secondary"
                    onClick={() => act(member.id, 'suspend')}
                  >
                    Tạm ngưng
                  </button>
                ) : null}
                {permissions.has('workspace.member.remove') ? (
                  <button
                    type="button"
                    className="button secondary"
                    onClick={() => act(member.id, 'remove')}
                  >
                    Gỡ
                  </button>
                ) : null}
              </div>
            ) : null}
          </li>
        ))}
      </ul>
    </section>
  );
}
