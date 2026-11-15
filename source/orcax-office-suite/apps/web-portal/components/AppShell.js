'use client';

import Link from 'next/link';
import { usePathname, useRouter } from 'next/navigation';
import { createContext, useContext, useState } from 'react';
import { kindLabel } from '../lib/labels';
import { WorkspaceDialog } from './WorkspaceDialog';

const SwitchContext = createContext(null);

export function useSwitchRequest() {
  return useContext(SwitchContext);
}

const LINKS = [
  { href: '/', label: 'Tổng quan' },
  { href: '/quan-tri', label: 'Quản trị' },
  { href: '/thanh-vien', label: 'Thành viên' },
  { href: '/phong-ban', label: 'Phòng ban' },
  { href: '/phien', label: 'Phiên đăng nhập' },
  { href: '/to-chuc', label: 'Tổ chức' },
  { href: '/loi-moi', label: 'Lời mời' },
  { href: '/nhat-ky', label: 'Nhật ký' },
  { href: '/giay-phep', label: 'Giấy phép' },
  { href: '/don-hang', label: 'Đơn hàng' },
  { href: '/ho-so', label: 'Hồ sơ' },
  { href: '/mau', label: 'Mẫu' },
  { href: '/cau-hinh', label: 'Cấu hình' },
  { href: '/add-in', label: 'Add-in' },
  { href: '/ho-tro', label: 'Hỗ trợ' },
];

export function AppShell({ children, session, workspaces = [], onSwitch, onLogout }) {
  const pathname = usePathname();
  const router = useRouter();
  const [dialog, setDialog] = useState(null);
  const workspace = session?.workspace;

  async function logout() {
    await onLogout();
    router.push('/dang-nhap');
  }

  return (
    <>
      <a className="skip-link" href="#noi-dung">
        Đi tới nội dung
      </a>
      <div className="shell">
        <aside className="sidebar">
          <div className="brand">
            <span className="brand-mark" aria-hidden="true">
              OX
            </span>
            <span>
              <strong>OrcaX</strong>
              <small>Office</small>
            </span>
          </div>
          <nav aria-label="Điều hướng chính">
            <p className="nav-label">Khung điều hướng</p>
            {(session ? LINKS : [{ href: '/', label: 'Tổng quan' }]).map((item) => (
              <Link
                key={item.href}
                className={pathname === item.href ? 'nav-link is-current' : 'nav-link'}
                href={item.href}
                aria-current={pathname === item.href ? 'page' : undefined}
              >
                {item.label}
              </Link>
            ))}
          </nav>
        </aside>
        <header className="topbar">
          <div>
            <p className="topbar-kicker">Cổng thông tin</p>
            <p className="topbar-title">{session ? session.user.displayName : 'Bàn làm việc'}</p>
          </div>
          <div className="topbar-actions">
            {session ? (
              <>
                <button
                  type="button"
                  className="workspace-switch is-ready"
                  onClick={() => setDialog({ initialId: '' })}
                >
                  {kindLabel(workspace.kind)}: {workspace.name}
                </button>
                <button type="button" className="button secondary" onClick={logout}>
                  Đăng xuất
                </button>
              </>
            ) : (
              <>
                <Link className="button secondary" href="/dang-nhap">
                  Đăng nhập
                </Link>
                <Link className="button" href="/dang-ky">
                  Đăng ký
                </Link>
              </>
            )}
          </div>
        </header>
        <div className="stage">
          <section className="quick-access" aria-label="Truy cập nhanh">
            <h2>Truy cập nhanh</h2>
            <div className="slots">
              {session ? (
                <>
                  <Link className="slot slot-link" href="/quan-tri">
                    Quản trị
                  </Link>
                  <Link className="slot slot-link" href="/don-hang">
                    Đơn hàng
                  </Link>
                  <Link className="slot slot-link" href="/ho-tro">
                    Hỗ trợ
                  </Link>
                </>
              ) : (
                <>
                  <p className="slot">Chưa có lối tắt</p>
                  <p className="slot">Chưa có lối tắt</p>
                  <p className="slot">Chưa có lối tắt</p>
                </>
              )}
            </div>
          </section>
          <main id="noi-dung" className="main">
            <SwitchContext.Provider
              value={(initialId) => setDialog({ initialId: initialId || '' })}
            >
              {children}
            </SwitchContext.Provider>
          </main>
        </div>
      </div>
      {dialog ? (
        <WorkspaceDialog
          workspaces={workspaces}
          currentId={workspace?.id}
          initialId={dialog.initialId}
          onCancel={() => setDialog(null)}
          onConfirm={async (workspaceId) => {
            await onSwitch(workspaceId);
            setDialog(null);
          }}
        />
      ) : null}
    </>
  );
}
