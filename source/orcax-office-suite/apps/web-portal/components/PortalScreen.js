'use client';

import Link from 'next/link';
import { AppShell } from './AppShell';
import { useWorkspaceSession } from './use-workspace-session';

export function PortalScreen({ children, allowAnonymous = false }) {
  const workspace = useWorkspaceSession();

  if (!workspace.ready) {
    return (
      <AppShell session={null}>
        <section className="panel">
          <h1>Đang tải phiên làm việc</h1>
        </section>
      </AppShell>
    );
  }

  if (!workspace.session && !allowAnonymous) {
    return (
      <main className="auth-screen" id="noi-dung">
        <div className="auth-card">
          <h1>Bạn cần đăng nhập</h1>
          <p className="lead">Phiên làm việc chưa có hoặc đã hết hạn.</p>
          <Link className="button" href="/dang-nhap">
            Đến trang đăng nhập
          </Link>
        </div>
      </main>
    );
  }

  return (
    <AppShell
      session={workspace.session}
      workspaces={workspace.workspaces}
      onSwitch={workspace.switchWorkspace}
      onLogout={workspace.logout}
    >
      {workspace.error ? <p className="form-error">{workspace.error}</p> : null}
      {typeof children === 'function' ? children(workspace) : children}
    </AppShell>
  );
}
