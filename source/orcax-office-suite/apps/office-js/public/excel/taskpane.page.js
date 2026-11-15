/* global OrcaXExcel */
(function () {
  const TOKEN_KEY = 'orcax.accessToken';
  const API_KEY = 'orcax.apiOrigin';
  let state = {
    token: null,
    apiOrigin: 'http://localhost:3001',
    bootstrap: null,
    profile: null,
    plan: null,
  };

  function $(id) {
    return document.getElementById(id);
  }

  function setStatus(message, isError) {
    const node = $('status');
    node.textContent = message || '';
    node.className = isError ? 'status error' : 'status';
  }

  function save(key, value) {
    try {
      localStorage.setItem(key, value);
    } catch {
      /* ignore */
    }
    if (Office.context && Office.context.roamingSettings) {
      Office.context.roamingSettings.set(key, value);
      Office.context.roamingSettings.saveAsync(function () {});
    }
  }

  function read(key) {
    if (Office.context && Office.context.roamingSettings) {
      const value = Office.context.roamingSettings.get(key);
      if (value) return value;
    }
    try {
      return localStorage.getItem(key);
    } catch {
      return null;
    }
  }

  async function api(path, options) {
    options = options || {};
    const headers = { accept: 'application/json' };
    if (state.token) headers.authorization = 'Bearer ' + state.token;
    if (options.body !== undefined) headers['content-type'] = 'application/json';
    const response = await fetch(state.apiOrigin.replace(/\/$/, '') + '/api/v1' + path, {
      method: options.method || 'GET',
      headers: headers,
      body: options.body === undefined ? undefined : JSON.stringify(options.body),
    });
    const payload = await response.json().catch(function () {
      return null;
    });
    if (!response.ok) {
      throw new Error((payload && payload.error && payload.error.message) || 'Yêu cầu thất bại.');
    }
    return payload.data;
  }

  function inchesToMm(value) {
    const inches = Number(value);
    if (!Number.isFinite(inches)) return null;
    return Math.round(inches * 25.4 * 10) / 10;
  }

  function mmToInches(value) {
    return Number(value) / 25.4;
  }

  async function readSnapshot() {
    return Excel.run(async function (context) {
      const pageSetup = context.workbook.worksheets.getActiveWorksheet().pageLayout;
      pageSetup.load([
        'paperSize',
        'orientation',
        'topMargin',
        'bottomMargin',
        'leftMargin',
        'rightMargin',
      ]);
      await context.sync();
      return {
        pageSetup: {
          paperSize: String(pageSetup.paperSize || ''),
          orientation: String(pageSetup.orientation || ''),
          marginMm: {
            top: inchesToMm(pageSetup.topMargin),
            bottom: inchesToMm(pageSetup.bottomMargin),
            left: inchesToMm(pageSetup.leftMargin),
            right: inchesToMm(pageSetup.rightMargin),
          },
        },
      };
    });
  }

  async function applySettings(settings) {
    return Excel.run(async function (context) {
      const pageSetup = context.workbook.worksheets.getActiveWorksheet().pageLayout;
      if (settings.orientation === 'portrait')
        pageSetup.orientation = Excel.PageOrientation.portrait;
      if (settings.orientation === 'landscape')
        pageSetup.orientation = Excel.PageOrientation.landscape;
      if (Excel.PaperSize && Excel.PaperSize.a4) pageSetup.paperSize = Excel.PaperSize.a4;
      pageSetup.topMargin = mmToInches(settings.marginMm.top);
      pageSetup.bottomMargin = mmToInches(settings.marginMm.bottom);
      pageSetup.leftMargin = mmToInches(settings.marginMm.left);
      pageSetup.rightMargin = mmToInches(settings.marginMm.right);
      await context.sync();
    });
  }

  function renderProfiles() {
    const select = $('profileId');
    select.innerHTML = '';
    (state.bootstrap.profiles || []).forEach(function (profile) {
      const option = document.createElement('option');
      option.value = profile.id;
      option.textContent = profile.name + ' (v' + profile.currentVersion + ')';
      select.appendChild(option);
    });
  }

  async function bootstrap() {
    state.token = read(TOKEN_KEY);
    state.apiOrigin = read(API_KEY) || $('apiOrigin').value;
    if (!state.token) throw new Error('Cần đăng nhập.');
    state.bootstrap = await api('/add-in/excel/bootstrap');
    renderProfiles();
    if (!state.bootstrap.entitlement || !state.bootstrap.entitlement.allowed) {
      setStatus('Workspace chưa có quyền lợi excel.core.', true);
    }
  }

  async function recordRun(extra) {
    await api('/add-in/excel/print-runs', {
      method: 'POST',
      body: Object.assign(
        {
          profileId: state.profile.id,
          profileVersionId: state.profile.version.id,
          profileVersion: state.profile.currentVersion,
          profileCode: state.profile.code,
          paperSize: state.plan.preview.paperSize,
          orientation: state.plan.preview.orientation,
          adapter: 'office-js',
        },
        extra,
      ),
    });
  }

  $('loginForm').addEventListener('submit', async function (event) {
    event.preventDefault();
    try {
      state.apiOrigin = $('apiOrigin').value.trim();
      state.token = null;
      const data = await api('/auth/login', {
        method: 'POST',
        body: { email: $('email').value.trim(), password: $('password').value },
      });
      state.token = data.accessToken;
      save(TOKEN_KEY, state.token);
      save(API_KEY, state.apiOrigin);
      await bootstrap();
      setStatus('Đăng nhập thành công.');
    } catch (error) {
      setStatus(error.message, true);
    }
  });

  $('logoutBtn').addEventListener('click', function () {
    state.token = null;
    save(TOKEN_KEY, '');
    setStatus('Đã đăng xuất.');
  });

  $('previewBtn').addEventListener('click', async function () {
    try {
      await bootstrap();
      state.profile = await api('/profiles/' + $('profileId').value);
      const snapshot = await readSnapshot();
      state.plan = OrcaXExcel.planExcelPrint(state.profile, snapshot);
      $('preview').textContent = JSON.stringify(state.plan.preview, null, 2);
      setStatus(
        state.plan.changes.length
          ? 'Có khác biệt. Xác nhận trước khi áp dụng.'
          : 'Trang in đã khớp hồ sơ.',
      );
    } catch (error) {
      setStatus(error.message, true);
    }
  });

  $('applyBtn').addEventListener('click', async function () {
    try {
      if (!state.plan || !state.profile) {
        setStatus('Hãy xem trước trước.', true);
        return;
      }
      const confirmed = window.confirm('Áp dụng cấu hình in theo hồ sơ?');
      const result = OrcaXExcel.applyExcelPrintPlan(state.plan, confirmed, function (settings) {
        return applySettings(settings);
      });
      if (result.cancelled) {
        setStatus(result.message);
        return;
      }
      if (result.applied) await applySettings(state.plan.preview);
      await recordRun({ applied: result.applied, pdfExported: false });
      setStatus(result.message);
    } catch (error) {
      setStatus(error.message, true);
    }
  });

  $('pdfBtn').addEventListener('click', async function () {
    try {
      await bootstrap();
      if (!state.profile || !state.plan) {
        setStatus('Hãy xem trước / chọn hồ sơ trước.', true);
        return;
      }
      await new Promise(function (resolve) {
        Office.context.document.getFileAsync(Office.FileType.Pdf, function (result) {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            setStatus(result.error?.message || 'Host không hỗ trợ xuất PDF.', true);
            resolve();
            return;
          }
          result.value.closeAsync(async function () {
            await recordRun({ applied: false, pdfExported: true });
            setStatus('Đã tạo luồng PDF qua host Office.');
            resolve();
          });
        });
      });
    } catch (error) {
      setStatus(error.message, true);
    }
  });

  Office.onReady(async function () {
    $('apiOrigin').value = read(API_KEY) || 'http://localhost:3001';
    try {
      await bootstrap();
      setStatus('Sẵn sàng.');
    } catch {
      setStatus('Đăng nhập để tải hồ sơ Excel.');
    }
  });
})();
