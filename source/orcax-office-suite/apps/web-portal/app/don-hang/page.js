'use client';

import { useEffect, useState } from 'react';
import { PortalScreen } from '../../components/PortalScreen';
import { api } from '../../lib/api';

export default function OrdersPage() {
  return (
    <PortalScreen>
      {(workspace) => <Orders key={workspace.session.workspace.id} session={workspace.session} />}
    </PortalScreen>
  );
}

function formatVND(amountMinor) {
  return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(amountMinor);
}

function statusBadge(status) {
  const map = {
    draft: { label: 'Bản nháp', cls: 'badge-draft' },
    pending_payment: { label: 'Chờ thanh toán', cls: 'badge-pending' },
    paid: { label: 'Đã thanh toán', cls: 'badge-success' },
    cancelled: { label: 'Đã hủy', cls: 'badge-cancelled' },
    refunded: { label: 'Đã hoàn tiền', cls: 'badge-refunded' },
    partially_refunded: { label: 'Hoàn một phần', cls: 'badge-refunded' },
  };
  const item = map[status] || { label: status, cls: '' };
  return <span className={`badge ${item.cls}`}>{item.label}</span>;
}

function Orders({ session }) {
  const permissions = new Set(session.permissions ?? []);
  const canRead = permissions.has('commerce.order.read');
  const canCreate = permissions.has('commerce.order.create');
  const canReconcile = permissions.has('commerce.reconcile');

  if (!canRead) {
    return (
      <section className="panel">
        <p className="eyebrow">Đơn hàng</p>
        <h1>Không có quyền xem đơn hàng</h1>
      </section>
    );
  }

  return <OrdersPanel session={session} canCreate={canCreate} canReconcile={canReconcile} />;
}

function OrdersPanel({ canCreate, canReconcile }) {
  const [prices, setPrices] = useState([]);
  const [orders, setOrders] = useState([]);
  const [reconciliations, setReconciliations] = useState([]);
  const [selectedPlanId, setSelectedPlanId] = useState('');
  const [quantity, setQuantity] = useState(1);
  const [error, setError] = useState('');
  const [note, setNote] = useState('');
  const [loading, setLoading] = useState(false);

  async function reload() {
    try {
      const [pricesRes, ordersRes] = await Promise.all([api('/plan-prices'), api('/orders')]);
      setPrices(pricesRes.prices ?? []);
      setOrders(ordersRes.orders ?? []);
      if (!selectedPlanId && pricesRes.prices?.length) {
        setSelectedPlanId(pricesRes.prices[0].planId);
      }
      if (canReconcile) {
        const reconRes = await api('/reconciliation-items');
        setReconciliations(reconRes.items ?? []);
      }
    } catch (err) {
      setError(err.message);
    }
  }

  useEffect(() => {
    let ignore = false;
    const calls = [api('/plan-prices'), api('/orders')];
    if (canReconcile) {
      calls.push(api('/reconciliation-items'));
    }
    Promise.all(calls)
      .then(([pricesRes, ordersRes, reconRes]) => {
        if (ignore) return;
        setPrices(pricesRes.prices ?? []);
        setOrders(ordersRes.orders ?? []);
        setSelectedPlanId((cur) => cur || pricesRes.prices?.[0]?.planId || '');
        if (reconRes) {
          setReconciliations(reconRes.items ?? []);
        }
        setError('');
      })
      .catch((err) => {
        if (!ignore) setError(err.message);
      });
    return () => {
      ignore = true;
    };
  }, [canReconcile]);

  async function createOrder(event) {
    event.preventDefault();
    setError('');
    setNote('');
    setLoading(true);
    try {
      await api('/orders', {
        method: 'POST',
        body: { items: [{ planId: selectedPlanId, quantity: Number(quantity) }] },
      });
      setNote('Đã tạo đơn hàng mới.');
      await reload();
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  async function checkout(orderId) {
    setError('');
    setNote('');
    setLoading(true);
    try {
      const res = await api('/payments/checkout', {
        method: 'POST',
        body: { orderId },
      });
      if (res.status === 'succeeded') {
        setNote('Đơn hàng miễn phí đã được hoàn tất và kích hoạt giấy phép thành công.');
      } else {
        setNote(`Đã khởi tạo thanh toán mã ${res.paymentId}. Đơn hàng đang chờ đối soát webhook.`);
      }
      await reload();
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="stack-vertical">
      <section className="panel">
        <p className="eyebrow">Thương mại & Gói dịch vụ</p>
        <h1>Bảng giá & Đặt đơn hàng</h1>
        <p className="lead">
          Các gói bản quyền OrcaX Office Suite dành riêng cho loại không gian hiện tại.
        </p>

        {note ? <p className="alert success">{note}</p> : null}
        {error ? <p className="alert error">{error}</p> : null}

        <div className="cards-grid">
          {prices.map((p) => (
            <div key={p.id} className="card price-card">
              <h3>{p.planName}</h3>
              <p className="price-amount">{formatVND(p.amountMinor)}</p>
              <p className="price-period">
                {p.billingPeriod === 'year'
                  ? 'Theo năm'
                  : p.billingPeriod === 'month'
                    ? 'Theo tháng'
                    : 'Thanh toán một lần'}
              </p>
              {canCreate ? (
                <button
                  type="button"
                  className="button"
                  onClick={() => {
                    setSelectedPlanId(p.planId);
                    setQuantity(1);
                  }}
                >
                  Chọn gói
                </button>
              ) : null}
            </div>
          ))}
        </div>

        {canCreate && prices.length > 0 ? (
          <form className="stack-form" onSubmit={createOrder} style={{ marginTop: '24px' }}>
            <h2>Tạo đơn hàng</h2>
            <label>
              Gói đã chọn
              <select value={selectedPlanId} onChange={(e) => setSelectedPlanId(e.target.value)}>
                {prices.map((p) => (
                  <option key={p.planId} value={p.planId}>
                    {p.planName} - {formatVND(p.amountMinor)}
                  </option>
                ))}
              </select>
            </label>
            <label>
              Số lượng
              <input
                type="number"
                min="1"
                max="100"
                value={quantity}
                onChange={(e) => setQuantity(e.target.value)}
              />
            </label>
            <button type="submit" className="button" disabled={loading}>
              {loading ? 'Đang xử lý...' : 'Xác nhận tạo đơn'}
            </button>
          </form>
        ) : null}
      </section>

      <section className="panel">
        <h2>Lịch sử đơn hàng</h2>
        {orders.length === 0 ? (
          <p className="empty-hint">Chưa có đơn hàng nào trong không gian này.</p>
        ) : (
          <table className="data-table">
            <thead>
              <tr>
                <th>Mã đơn</th>
                <th>Thời gian</th>
                <th>Tổng tiền</th>
                <th>Trạng thái</th>
                <th>Thao tác</th>
              </tr>
            </thead>
            <tbody>
              {orders.map((o) => (
                <tr key={o.id}>
                  <td>
                    <code>{o.id.slice(0, 8)}</code>
                  </td>
                  <td>{new Date(o.createdAt).toLocaleString('vi-VN')}</td>
                  <td>
                    <strong>{formatVND(o.totalMinor)}</strong>
                  </td>
                  <td>{statusBadge(o.status)}</td>
                  <td>
                    {canCreate && (o.status === 'draft' || o.status === 'pending_payment') ? (
                      <button
                        type="button"
                        className="button small"
                        disabled={loading}
                        onClick={() => checkout(o.id)}
                      >
                        Thanh toán
                      </button>
                    ) : (
                      <span className="muted">—</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>

      {canReconcile && reconciliations.length > 0 ? (
        <section className="panel">
          <h2>Hàng đợi đối soát thanh toán</h2>
          <table className="data-table">
            <thead>
              <tr>
                <th>Mã đối soát</th>
                <th>Thời gian</th>
                <th>Trạng thái</th>
                <th>Ghi chú</th>
              </tr>
            </thead>
            <tbody>
              {reconciliations.map((r) => (
                <tr key={r.id}>
                  <td>
                    <code>{r.id.slice(0, 8)}</code>
                  </td>
                  <td>{new Date(r.createdAt).toLocaleString('vi-VN')}</td>
                  <td>
                    <span className="badge">{r.status}</span>
                  </td>
                  <td>{r.note}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </section>
      ) : null}
    </div>
  );
}
