import { useState } from 'react';
import { kindLabel, roleLabel } from '../lib/labels';

export function WorkspaceDialog({ workspaces, currentId, initialId, onCancel, onConfirm }) {
  const choices = workspaces.filter((item) => item.switchable);
  const [selectedId, setSelectedId] = useState(initialId || currentId || choices[0]?.id || '');
  const [confirmed, setConfirmed] = useState(false);
  const [error, setError] = useState('');
  const [pending, setPending] = useState(false);

  async function submit(event) {
    event.preventDefault();
    if (!confirmed) {
      setError('Hãy xác nhận trước khi chuyển không gian làm việc.');
      return;
    }
    if (!selectedId || selectedId === currentId) {
      setError('Hãy chọn một không gian khác không gian hiện tại.');
      return;
    }
    setPending(true);
    setError('');
    try {
      await onConfirm(selectedId);
    } catch (reason) {
      setError(reason.message);
      setPending(false);
    }
  }

  return (
    <div className="dialog-backdrop" role="presentation">
      <form
        className="dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="chuyen-tieu-de"
        onSubmit={submit}
      >
        <h2 id="chuyen-tieu-de">Chuyển không gian làm việc</h2>
        <p>
          Thao tác này chỉ chạy sau khi bạn chọn không gian và xác nhận. Danh sách bên dưới chưa gửi
          yêu cầu.
        </p>
        <fieldset className="choice-list">
          <legend>Không gian có thể chuyển</legend>
          {choices.map((item) => (
            <label key={item.id}>
              <input
                type="radio"
                name="workspace"
                value={item.id}
                checked={selectedId === item.id}
                onChange={() => {
                  setSelectedId(item.id);
                  setError('');
                }}
              />
              <span>
                {item.name} · {kindLabel(item.kind)} · {roleLabel(item.role)}
                {item.id === currentId ? ' · đang dùng' : ''}
              </span>
            </label>
          ))}
        </fieldset>
        <label className="confirm-line">
          <input
            type="checkbox"
            checked={confirmed}
            onChange={(event) => {
              setConfirmed(event.target.checked);
              setError('');
            }}
          />
          Tôi xác nhận chuyển sang không gian đã chọn.
        </label>
        {error ? <p className="form-error">{error}</p> : null}
        <div className="form-actions">
          <button type="button" className="button secondary" onClick={onCancel} disabled={pending}>
            Để sau
          </button>
          <button type="submit" className="button" disabled={pending}>
            {pending ? 'Đang chuyển' : 'Chuyển'}
          </button>
        </div>
      </form>
    </div>
  );
}
