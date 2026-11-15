'use client';

import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';

export default function DepartmentsPage() {
  return (
    <PortalScreen>
      {(workspace) => (
        <Departments key={workspace.session.workspace.id} session={workspace.session} />
      )}
    </PortalScreen>
  );
}

function Departments({ session }) {
  const permissions = new Set(session.permissions ?? []);
  if (
    session.workspace.kind !== 'organization' ||
    !permissions.has('workspace.department.manage')
  ) {
    return (
      <section className="panel">
        <p className="eyebrow">Phòng ban</p>
        <h1>Không có quyền quản lý phòng ban / nhóm</h1>
        <p className="lead">Chỉ quản trị tổ chức mới tạo phòng ban và nhóm.</p>
      </section>
    );
  }
  return <DepartmentsPanel />;
}

function DepartmentsPanel() {
  const [departments, setDepartments] = useState([]);
  const [groups, setGroups] = useState([]);
  const [members, setMembers] = useState([]);
  const [deptCode, setDeptCode] = useState('');
  const [deptName, setDeptName] = useState('');
  const [groupCode, setGroupCode] = useState('');
  const [groupName, setGroupName] = useState('');
  const [deptId, setDeptId] = useState('');
  const [groupId, setGroupId] = useState('');
  const [userId, setUserId] = useState('');
  const [error, setError] = useState('');
  const [note, setNote] = useState('');

  async function reload() {
    const [deptResult, groupResult, memberResult] = await Promise.all([
      api('/departments'),
      api('/groups'),
      api('/members'),
    ]);
    setDepartments(deptResult.departments ?? []);
    setGroups(groupResult.groups ?? []);
    setMembers((memberResult.members ?? []).filter((item) => item.status === 'active'));
    setDeptId((current) => current || deptResult.departments?.[0]?.id || '');
    setGroupId((current) => current || groupResult.groups?.[0]?.id || '');
    setUserId((current) => current || memberResult.members?.[0]?.userId || '');
  }

  useEffect(() => {
    let ignore = false;
    Promise.all([api('/departments'), api('/groups'), api('/members')])
      .then(([deptResult, groupResult, memberResult]) => {
        if (ignore) {
          return;
        }
        setDepartments(deptResult.departments ?? []);
        setGroups(groupResult.groups ?? []);
        setMembers((memberResult.members ?? []).filter((item) => item.status === 'active'));
        setDeptId(deptResult.departments?.[0]?.id || '');
        setGroupId(groupResult.groups?.[0]?.id || '');
        setUserId(memberResult.members?.[0]?.userId || '');
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
      <p className="eyebrow">Phòng ban &amp; nhóm</p>
      <h1>Cơ cấu tổ chức</h1>
      <p className="lead">
        Phòng ban có thể lồng nhau. Phân phối hồ sơ đã duyệt gắn theo phòng ban hoặc nhóm.
      </p>
      {error ? <p className="form-error">{error}</p> : null}
      {note ? <p className="form-note">{note}</p> : null}

      <form
        className="stack-form"
        onSubmit={async (event) => {
          event.preventDefault();
          setError('');
          try {
            await api('/departments', { method: 'POST', body: { code: deptCode, name: deptName } });
            setNote('Đã tạo phòng ban.');
            setDeptCode('');
            setDeptName('');
            await reload();
          } catch (reason) {
            setError(reason.message);
          }
        }}
      >
        <h2>Tạo phòng ban</h2>
        <label>
          Mã
          <input value={deptCode} onChange={(event) => setDeptCode(event.target.value)} required />
        </label>
        <label>
          Tên
          <input value={deptName} onChange={(event) => setDeptName(event.target.value)} required />
        </label>
        <button type="submit" className="button">
          Tạo phòng ban
        </button>
      </form>

      <form
        className="stack-form"
        onSubmit={async (event) => {
          event.preventDefault();
          setError('');
          try {
            await api('/groups', { method: 'POST', body: { code: groupCode, name: groupName } });
            setNote('Đã tạo nhóm.');
            setGroupCode('');
            setGroupName('');
            await reload();
          } catch (reason) {
            setError(reason.message);
          }
        }}
      >
        <h2>Tạo nhóm</h2>
        <label>
          Mã
          <input
            value={groupCode}
            onChange={(event) => setGroupCode(event.target.value)}
            required
          />
        </label>
        <label>
          Tên
          <input
            value={groupName}
            onChange={(event) => setGroupName(event.target.value)}
            required
          />
        </label>
        <button type="submit" className="button">
          Tạo nhóm
        </button>
      </form>

      <form
        className="stack-form"
        onSubmit={async (event) => {
          event.preventDefault();
          setError('');
          try {
            await api(`/departments/${deptId}/members`, { method: 'POST', body: { userId } });
            setNote('Đã gán thành viên vào phòng ban.');
            await reload();
          } catch (reason) {
            setError(reason.message);
          }
        }}
      >
        <h2>Gán thành viên</h2>
        <label>
          Phòng ban
          <select value={deptId} onChange={(event) => setDeptId(event.target.value)}>
            {departments.map((item) => (
              <option key={item.id} value={item.id}>
                {item.name}
              </option>
            ))}
          </select>
        </label>
        <label>
          Nhóm
          <select value={groupId} onChange={(event) => setGroupId(event.target.value)}>
            {groups.map((item) => (
              <option key={item.id} value={item.id}>
                {item.name}
              </option>
            ))}
          </select>
        </label>
        <label>
          Thành viên
          <select value={userId} onChange={(event) => setUserId(event.target.value)}>
            {members.map((item) => (
              <option key={item.userId} value={item.userId}>
                {item.displayName || item.email}
              </option>
            ))}
          </select>
        </label>
        <div className="form-actions">
          <button type="submit" className="button">
            Gán vào phòng ban
          </button>
          <button
            type="button"
            className="button secondary"
            onClick={async () => {
              setError('');
              try {
                await api(`/groups/${groupId}/members`, { method: 'POST', body: { userId } });
                setNote('Đã gán thành viên vào nhóm.');
                await reload();
              } catch (reason) {
                setError(reason.message);
              }
            }}
          >
            Gán vào nhóm
          </button>
        </div>
      </form>

      <h2>Phòng ban</h2>
      <ul className="record-list">
        {departments.map((item) => (
          <li key={item.id}>
            <strong>
              {item.name} ({item.code})
            </strong>
            <span>{item.memberUserIds.length} thành viên</span>
          </li>
        ))}
      </ul>
      <h2>Nhóm</h2>
      <ul className="record-list">
        {groups.map((item) => (
          <li key={item.id}>
            <strong>
              {item.name} ({item.code})
            </strong>
            <span>{item.memberUserIds.length} thành viên</span>
          </li>
        ))}
      </ul>
    </section>
  );
}
