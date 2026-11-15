'use client';

import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';

export default function SupportPage() {
  return (
    <PortalScreen>
      {(workspace) => <Support key={workspace.session.workspace.id} session={workspace.session} />}
    </PortalScreen>
  );
}

const CATEGORIES = [
  { code: 'all', label: 'Tất cả' },
  { code: 'general', label: 'Chung' },
  { code: 'account', label: 'Tài khoản' },
  { code: 'billing', label: 'Thanh toán' },
  { code: 'licensing', label: 'Giấy phép' },
  { code: 'word_addin', label: 'Add-in Word' },
  { code: 'excel_addin', label: 'Add-in Excel' },
];

function statusBadge(status) {
  const map = {
    open: { label: 'Đang mở', cls: 'badge-pending' },
    in_progress: { label: 'Đang xử lý', cls: 'badge-draft' },
    resolved: { label: 'Đã giải quyết', cls: 'badge-success' },
    closed: { label: 'Đã đóng', cls: 'badge-cancelled' },
  };
  const item = map[status] || { label: status, cls: '' };
  return <span className={`badge ${item.cls}`}>{item.label}</span>;
}

function priorityLabel(priority) {
  const map = {
    low: 'Thấp',
    medium: 'Bình thường',
    high: 'Cao',
    urgent: 'Khẩn cấp',
  };
  return map[priority] || priority;
}

function Support({ session }) {
  const permissions = new Set(session.permissions ?? []);
  const canReadTickets = permissions.has('support.ticket.read');
  const canCreateTicket = permissions.has('support.ticket.create');
  const canManageTicket = permissions.has('support.ticket.manage');
  const canReadHelp = permissions.has('support.help.read');

  const [activeTab, setActiveTab] = useState('help'); // 'help' | 'tickets'
  const [articles, setArticles] = useState([]);
  const [selectedArticle, setSelectedArticle] = useState(null);
  const [selectedCategory, setSelectedCategory] = useState('all');
  const [searchQuery, setSearchQuery] = useState('');

  const [tickets, setTickets] = useState([]);
  const [selectedTicket, setSelectedTicket] = useState(null);
  const [newTitle, setNewTitle] = useState('');
  const [newCategory, setNewCategory] = useState('general');
  const [newPriority, setNewPriority] = useState('medium');
  const [newDescription, setNewDescription] = useState('');
  const [newComment, setNewComment] = useState('');

  const [error, setError] = useState('');
  const [note, setNote] = useState('');
  const [loading, setLoading] = useState(false);

  async function loadHelp() {
    try {
      const url =
        selectedCategory && selectedCategory !== 'all'
          ? `/support/help/articles?category=${selectedCategory}`
          : '/support/help/articles';
      const res = await api(url);
      setArticles(res.articles ?? []);
    } catch (err) {
      setError(err.message);
    }
  }

  async function loadArticleDetail(slug) {
    try {
      const article = await api(`/support/help/articles/${slug}`);
      setSelectedArticle(article);
    } catch (err) {
      setError(err.message);
    }
  }

  async function loadTickets() {
    if (!canReadTickets) return;
    try {
      const res = await api('/support/tickets');
      setTickets(res.tickets ?? []);
    } catch (err) {
      setError(err.message);
    }
  }

  async function loadTicketDetail(id) {
    try {
      const ticket = await api(`/support/tickets/${id}`);
      setSelectedTicket(ticket);
    } catch (err) {
      setError(err.message);
    }
  }

  useEffect(() => {
    let ignore = false;
    const calls = [];
    if (canReadHelp) {
      const url =
        selectedCategory && selectedCategory !== 'all'
          ? `/support/help/articles?category=${selectedCategory}`
          : '/support/help/articles';
      calls.push(api(url).then((res) => ({ type: 'help', res })));
    }
    if (canReadTickets) {
      calls.push(api('/support/tickets').then((res) => ({ type: 'tickets', res })));
    }
    Promise.all(calls)
      .then((results) => {
        if (ignore) return;
        for (const item of results) {
          if (item.type === 'help') {
            setArticles(item.res.articles ?? []);
          } else if (item.type === 'tickets') {
            setTickets(item.res.tickets ?? []);
          }
        }
        setError('');
      })
      .catch((err) => {
        if (!ignore) setError(err.message);
      });
    return () => {
      ignore = true;
    };
  }, [selectedCategory, canReadHelp, canReadTickets]);

  async function handleCreateTicket(event) {
    event.preventDefault();
    setError('');
    setNote('');
    setLoading(true);
    try {
      const created = await api('/support/tickets', {
        method: 'POST',
        body: {
          title: newTitle,
          category: newCategory,
          priority: newPriority,
          description: newDescription,
        },
      });
      setNote('Đã tạo yêu cầu hỗ trợ thành công.');
      setNewTitle('');
      setNewDescription('');
      await loadTickets();
      setSelectedTicket(created);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleAddComment(event) {
    event.preventDefault();
    if (!selectedTicket || !newComment.trim()) return;
    setError('');
    setLoading(true);
    try {
      await api(`/support/tickets/${selectedTicket.id}/comments`, {
        method: 'POST',
        body: { message: newComment },
      });
      setNewComment('');
      await loadTicketDetail(selectedTicket.id);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleUpdateStatus(status) {
    if (!selectedTicket) return;
    setError('');
    setLoading(true);
    try {
      await api(`/support/tickets/${selectedTicket.id}/status`, {
        method: 'PATCH',
        body: { status },
      });
      await loadTicketDetail(selectedTicket.id);
      await loadTickets();
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  const filteredArticles = articles.filter(
    (a) =>
      a.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      a.slug.toLowerCase().includes(searchQuery.toLowerCase()),
  );

  return (
    <div className="stack-vertical">
      <section className="panel">
        <p className="eyebrow">Trợ giúp & Kỹ thuật</p>
        <h1>Trung tâm hỗ trợ OrcaX Office</h1>

        <div className="tab-bar" style={{ display: 'flex', gap: '12px', margin: '16px 0' }}>
          <button
            type="button"
            className={activeTab === 'help' ? 'button' : 'button secondary'}
            onClick={() => setActiveTab('help')}
          >
            Hướng dẫn & FAQ
          </button>
          {canReadTickets ? (
            <button
              type="button"
              className={activeTab === 'tickets' ? 'button' : 'button secondary'}
              onClick={() => setActiveTab('tickets')}
            >
              Yêu cầu hỗ trợ {tickets.length ? `(${tickets.length})` : ''}
            </button>
          ) : null}
        </div>

        {note ? <p className="alert success">{note}</p> : null}
        {error ? <p className="alert error">{error}</p> : null}

        {activeTab === 'help' ? (
          <div>
            <div style={{ display: 'flex', gap: '16px', marginBottom: '16px', flexWrap: 'wrap' }}>
              <input
                type="text"
                placeholder="Tìm kiếm tài liệu hướng dẫn..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                style={{ flex: 1, minWidth: '240px' }}
              />
              <select
                value={selectedCategory}
                onChange={(e) => {
                  setSelectedCategory(e.target.value);
                  setSelectedArticle(null);
                }}
              >
                {CATEGORIES.map((c) => (
                  <option key={c.code} value={c.code}>
                    {c.label}
                  </option>
                ))}
              </select>
            </div>

            {selectedArticle ? (
              <div className="card article-detail" style={{ padding: '20px' }}>
                <button
                  type="button"
                  className="button secondary small"
                  onClick={() => setSelectedArticle(null)}
                  style={{ marginBottom: '12px' }}
                >
                  ← Quay lại danh sách
                </button>
                <h2>{selectedArticle.title}</h2>
                <p className="muted" style={{ fontSize: '13px', margin: '4px 0 16px' }}>
                  Phiên bản: {selectedArticle.productVersion} · Danh mục: {selectedArticle.category}
                </p>
                <div style={{ lineHeight: '1.6', whiteSpace: 'pre-line' }}>
                  {selectedArticle.content}
                </div>
              </div>
            ) : (
              <div className="cards-grid">
                {filteredArticles.length === 0 ? (
                  <p className="empty-hint">Không có bài viết nào phù hợp.</p>
                ) : (
                  filteredArticles.map((a) => (
                    <div
                      key={a.id}
                      className="card slot-link"
                      onClick={() => loadArticleDetail(a.slug)}
                      style={{ cursor: 'pointer' }}
                    >
                      <h3>{a.title}</h3>
                      <p className="muted">
                        Phiên bản {a.productVersion} · {a.category}
                      </p>
                    </div>
                  ))
                )}
              </div>
            )}
          </div>
        ) : (
          <div>
            {canCreateTicket ? (
              <form
                className="stack-form"
                onSubmit={handleCreateTicket}
                style={{ marginBottom: '24px' }}
              >
                <h3>Tạo yêu cầu hỗ trợ mới</h3>
                <label>
                  Tiêu đề
                  <input
                    type="text"
                    required
                    value={newTitle}
                    onChange={(e) => setNewTitle(e.target.value)}
                    placeholder="Mô tả tóm tắt vấn đề..."
                  />
                </label>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                  <label>
                    Danh mục
                    <select value={newCategory} onChange={(e) => setNewCategory(e.target.value)}>
                      <option value="general">Chung</option>
                      <option value="account">Tài khoản & Không gian</option>
                      <option value="billing">Đơn hàng & Thanh toán</option>
                      <option value="licensing">Giấy phép & Kích hoạt</option>
                      <option value="word_addin">Add-in Word</option>
                      <option value="excel_addin">Add-in Excel</option>
                    </select>
                  </label>
                  <label>
                    Mức độ ưu tiên
                    <select value={newPriority} onChange={(e) => setNewPriority(e.target.value)}>
                      <option value="low">Thấp</option>
                      <option value="medium">Bình thường</option>
                      <option value="high">Cao</option>
                      <option value="urgent">Khẩn cấp</option>
                    </select>
                  </label>
                </div>
                <label>
                  Nội dung chi tiết
                  <textarea
                    rows={4}
                    required
                    value={newDescription}
                    onChange={(e) => setNewDescription(e.target.value)}
                    placeholder="Vui lòng cung cấp chi tiết lỗi hoặc thắc mắc cần hỗ trợ..."
                  />
                </label>
                <button type="submit" className="button" disabled={loading}>
                  {loading ? 'Đang gửi...' : 'Gửi yêu cầu'}
                </button>
              </form>
            ) : null}

            {selectedTicket ? (
              <div className="card ticket-detail" style={{ padding: '20px', marginTop: '16px' }}>
                <button
                  type="button"
                  className="button secondary small"
                  onClick={() => setSelectedTicket(null)}
                  style={{ marginBottom: '12px' }}
                >
                  ← Danh sách yêu cầu
                </button>
                <div
                  style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}
                >
                  <h2>{selectedTicket.title}</h2>
                  {statusBadge(selectedTicket.status)}
                </div>
                <p className="muted" style={{ margin: '4px 0 16px' }}>
                  Danh mục: {selectedTicket.category} · Ưu tiên:{' '}
                  {priorityLabel(selectedTicket.priority)} · Tạo ngày:{' '}
                  {new Date(selectedTicket.createdAt).toLocaleString('vi-VN')}
                </p>

                {canManageTicket ? (
                  <div style={{ display: 'flex', gap: '8px', marginBottom: '16px' }}>
                    <span style={{ fontSize: '13px', alignSelf: 'center' }}>Đổi trạng thái:</span>
                    <button
                      type="button"
                      className="button secondary small"
                      onClick={() => handleUpdateStatus('in_progress')}
                    >
                      Đang xử lý
                    </button>
                    <button
                      type="button"
                      className="button secondary small"
                      onClick={() => handleUpdateStatus('resolved')}
                    >
                      Đã giải quyết
                    </button>
                    <button
                      type="button"
                      className="button secondary small"
                      onClick={() => handleUpdateStatus('closed')}
                    >
                      Đóng
                    </button>
                  </div>
                ) : null}

                <h4>Trao đổi ({selectedTicket.comments?.length || 0})</h4>
                <div
                  className="comment-list"
                  style={{
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '12px',
                    margin: '16px 0',
                  }}
                >
                  {selectedTicket.comments?.map((c) => (
                    <div
                      key={c.id}
                      style={{
                        padding: '12px',
                        background: c.isStaff
                          ? 'var(--color-bg-secondary, #f0f4f8)'
                          : 'var(--color-bg-tertiary, #fcfcfc)',
                        borderLeft: c.isStaff ? '4px solid #2b6cb0' : '4px solid #cbd5e0',
                        borderRadius: '4px',
                      }}
                    >
                      <div
                        style={{
                          display: 'flex',
                          justifyContent: 'space-between',
                          fontSize: '12px',
                          marginBottom: '4px',
                        }}
                      >
                        <strong>{c.isStaff ? 'Đội ngũ Hỗ trợ OrcaX' : 'Người gửi'}</strong>
                        <span className="muted">
                          {new Date(c.createdAt).toLocaleString('vi-VN')}
                        </span>
                      </div>
                      <div style={{ whiteSpace: 'pre-line' }}>{c.message}</div>
                    </div>
                  ))}
                </div>

                <form onSubmit={handleAddComment} style={{ marginTop: '16px' }}>
                  <textarea
                    rows={3}
                    placeholder="Nhập phản hồi..."
                    value={newComment}
                    onChange={(e) => setNewComment(e.target.value)}
                    style={{ width: '100%', marginBottom: '8px' }}
                  />
                  <button type="submit" className="button" disabled={loading || !newComment.trim()}>
                    Gửi phản hồi
                  </button>
                </form>
              </div>
            ) : (
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Mã</th>
                    <th>Tiêu đề</th>
                    <th>Danh mục</th>
                    <th>Ưu tiên</th>
                    <th>Trạng thái</th>
                    <th>Thời gian</th>
                  </tr>
                </thead>
                <tbody>
                  {tickets.length === 0 ? (
                    <tr>
                      <td colSpan={6} style={{ textAlign: 'center' }}>
                        Chưa có yêu cầu hỗ trợ nào.
                      </td>
                    </tr>
                  ) : (
                    tickets.map((t) => (
                      <tr
                        key={t.id}
                        onClick={() => loadTicketDetail(t.id)}
                        style={{ cursor: 'pointer' }}
                      >
                        <td>
                          <code>{t.id.slice(0, 8)}</code>
                        </td>
                        <td>
                          <strong>{t.title}</strong>
                        </td>
                        <td>{t.category}</td>
                        <td>{priorityLabel(t.priority)}</td>
                        <td>{statusBadge(t.status)}</td>
                        <td>{new Date(t.createdAt).toLocaleDateString('vi-VN')}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            )}
          </div>
        )}
      </section>
    </div>
  );
}
