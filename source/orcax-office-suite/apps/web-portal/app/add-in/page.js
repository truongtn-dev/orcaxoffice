'use client';

import { PortalScreen } from '../../components/PortalScreen';

export default function AddInPage() {
  return (
    <PortalScreen>
      {(workspace) => (
        <AddInGuide key={workspace.session.workspace.id} session={workspace.session} />
      )}
    </PortalScreen>
  );
}

function AddInGuide({ session }) {
  const permissions = new Set(session.permissions ?? []);
  const canWord = permissions.has('addin.word.use');
  const canExcel = permissions.has('addin.excel.use');

  return (
    <section className="panel">
      <p className="eyebrow">Add-in Office</p>
      <h1>Word &amp; Excel</h1>
      <p className="lead">
        Add-in xử lý tài liệu cục bộ trên máy. API chỉ nhận metadata báo cáo và cấu hình in, không
        nhận nội dung file.
      </p>

      <h2>Quyền hiện tại</h2>
      <ul className="record-list">
        <li>
          <strong>Word</strong>
          <span>{canWord ? 'Được dùng add-in Word' : 'Chưa có quyền addin.word.use'}</span>
        </li>
        <li>
          <strong>Excel</strong>
          <span>{canExcel ? 'Được dùng add-in Excel' : 'Chưa có quyền addin.excel.use'}</span>
        </li>
      </ul>

      <h2>Sideload Office.js</h2>
      <ol>
        <li>
          Build tĩnh: <code>npm run build -w @orcax/office-js</code>
        </li>
        <li>
          Phục vụ thư mục <code>apps/office-js/dist</code> qua HTTPS (mặc định manifest dùng{' '}
          <code>https://localhost:3002</code>).
        </li>
        <li>
          Sideload <code>manifests/word-manifest.xml</code> hoặc{' '}
          <code>manifests/excel-manifest.xml</code> trong Office.
        </li>
        <li>
          Đăng nhập trong task pane; bảo đảm workspace có giấy phép <code>word.core</code> /{' '}
          <code>excel.core</code>.
        </li>
      </ol>

      <h2>VSTO</h2>
      <p>
        Adapter Windows nằm ở <code>apps/office-vsto</code> (contract C#). Project designer đầy đủ
        chưa ship; Office.js là MVP chạy được trên host hiện đại.
      </p>
    </section>
  );
}
