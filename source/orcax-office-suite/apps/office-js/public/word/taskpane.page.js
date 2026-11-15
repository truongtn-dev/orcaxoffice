/* global OrcaXWord */
(function () {
  const TOKEN_KEY = 'orcax.accessToken';
  const API_KEY = 'orcax.apiOrigin';
  let state = {
    token: null,
    apiOrigin: 'http://localhost:3001',
    bootstrap: null,
    profile: null,
    report: null,
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

  async function readSnapshot() {
    return Word.run(async function (context) {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();
      for (let i = 0; i < paragraphs.items.length; i += 1) {
        paragraphs.items[i].load(['font/name', 'font/size', 'outlineLevel']);
      }
      await context.sync();
      const fontRuns = [];
      const headings = [];
      for (let i = 0; i < paragraphs.items.length; i += 1) {
        const paragraph = paragraphs.items[i];
        const location = 'paragraph:' + (i + 1);
        fontRuns.push({
          location: location,
          fontName: paragraph.font.name,
          fontSizePt: paragraph.font.size,
        });
        const level = Number(paragraph.outlineLevel);
        if (level >= 1 && level <= 9) headings.push({ location: location, level: level });
      }
      return { fontRuns: fontRuns, headings: headings };
    });
  }

  async function applyChange(change) {
    const payload = change.payload || {};
    return Word.run(async function (context) {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();
      for (let i = 0; i < paragraphs.items.length; i += 1) {
        if (payload.fontName) paragraphs.items[i].font.name = payload.fontName;
        if (payload.fontSizePt != null) paragraphs.items[i].font.size = payload.fontSizePt;
      }
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

  function renderReport(report) {
    if (!report) {
      $('report').textContent = 'Chưa có.';
      return;
    }
    const lines = [
      'Hồ sơ ' +
        report.profileCode +
        ' v' +
        report.profileVersion +
        ': pass ' +
        report.summary.pass +
        ', fail ' +
        report.summary.fail +
        ', warning ' +
        report.summary.warning,
    ];
    report.results.forEach(function (item) {
      lines.push(
        item.severity.toUpperCase() +
          ' · ' +
          item.ruleId +
          ' · ' +
          (item.location || '—') +
          ' · ' +
          (item.explanation || ''),
      );
    });
    $('report').textContent = lines.join('\n');
  }

  async function bootstrap() {
    state.token = read(TOKEN_KEY);
    state.apiOrigin = read(API_KEY) || $('apiOrigin').value;
    if (!state.token) throw new Error('Cần đăng nhập.');
    state.bootstrap = await api('/add-in/word/bootstrap');
    renderProfiles();
    if (!state.bootstrap.entitlement || !state.bootstrap.entitlement.allowed) {
      setStatus('Workspace chưa có quyền lợi word.core.', true);
    }
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

  $('validateBtn').addEventListener('click', async function () {
    try {
      await bootstrap();
      state.profile = await api('/profiles/' + $('profileId').value);
      const snapshot = await readSnapshot();
      state.report = OrcaXWord.validateWordDocument(state.profile, snapshot);
      state.plan = OrcaXWord.planWordFormatting(state.profile, snapshot);
      renderReport(state.report);
      $('preview').textContent = state.plan.changes.length
        ? state.plan.changes
            .map(function (item) {
              return item.summary;
            })
            .join('\n')
        : 'Không có thay đổi cần áp dụng.';
      setStatus('Đã kiểm tra cục bộ. Không gửi nội dung tài liệu.');
    } catch (error) {
      setStatus(error.message, true);
    }
  });

  $('applyBtn').addEventListener('click', async function () {
    try {
      if (!state.plan || !state.report || !state.profile) {
        setStatus('Hãy kiểm tra trước.', true);
        return;
      }
      const confirmed = window.confirm('Áp dụng các thay đổi đã xem trước?');
      const result = OrcaXWord.applyWordPlan(state.plan, confirmed, function (change) {
        return applyChange(change);
      });
      if (result.cancelled) {
        setStatus(result.message);
        return;
      }
      await Promise.all(
        state.plan.changes.map(function (change) {
          return applyChange(change);
        }),
      );
      await api('/add-in/word/validation-runs', {
        method: 'POST',
        body: {
          profileId: state.profile.id,
          profileVersionId: state.profile.version.id,
          profileVersion: state.profile.currentVersion,
          profileCode: state.profile.code,
          summary: state.report.summary,
          adapter: 'office-js',
        },
      });
      setStatus(result.message);
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
      setStatus('Đăng nhập để tải hồ sơ Word.');
    }
  });
})();
