# Quy trình Sao lưu và Phục hồi (Backup & Restore Runbook)

Tài liệu này hướng dẫn quy trình vận hành sao lưu định kỳ, kiểm tra tính toàn vẹn và phục hồi cơ sở dữ liệu PostgreSQL cho OrcaX Office Suite.

## 1. Nguyên tắc cốt lõi

- Dữ liệu `audit_events` là append-only và có trigger chặn chỉnh sửa/xóa. Bản sao lưu phải giữ nguyên các trigger và ràng buộc này.
- Phục hồi cơ sở dữ liệu phải được thử nghiệm định kỳ trên môi trường staging trước khi áp dụng trên production.
- Không lưu mật khẩu hoặc thông tin nhạy cảm dạng bản rõ trong các tập lệnh sao lưu.

## 2. Sao lưu cơ sở dữ liệu (Backup)

### 2.1 Sao lưu nhanh qua pg_dump

```powershell
# Chạy sao lưu định dạng tùy chỉnh (custom format - nén và hỗ trợ pg_restore linh hoạt)
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupFile = "backup_orcax_${timestamp}.dump"

pg_dump -h $env:DATABASE_HOST -p $env:DATABASE_PORT -U $env:DATABASE_USER -d $env:DATABASE_NAME -F c -b -v -f $backupFile
```

### 2.2 Sao lưu khi chạy qua Docker Compose

```powershell
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
docker compose -f infrastructure/deployment/docker-compose.yml exec -T db pg_dump -U orcax -d orcax_suite -F c > "backup_orcax_${timestamp}.dump"
```

## 3. Quy trình Phục hồi (Restore)

### 3.1 Khôi phục từ bản sao lưu dump

```powershell
# Tạo database mới hoặc dọn dẹp kết nối hiện tại
dropdb -h $env:DATABASE_HOST -p $env:DATABASE_PORT -U $env:DATABASE_USER orcax_suite --if-exists
createdb -h $env:DATABASE_HOST -p $env:DATABASE_PORT -U $env:DATABASE_USER orcax_suite

# Phục hồi dữ liệu
pg_restore -h $env:DATABASE_HOST -p $env:DATABASE_PORT -U $env:DATABASE_USER -d orcax_suite -v $backupFile
```

### 3.2 Kiểm tra tính toàn vẹn sau phục hồi

1. Kiểm tra trạng thái sẵn sàng của dịch vụ:
   ```bash
   curl -f http://localhost:3001/health/ready
   ```
2. Kiểm tra bảng migration:
   ```powershell
   npm run migration:show
   ```
3. Xác minh trigger bảng `audit_events`:
   Kiểm tra rằng lệnh `DELETE FROM audit_events` vẫn bị chặn bởi trigger `audit_events_immutable`.

## 4. Rollback Migration

Khi cần hoàn tác migration gần nhất:

```powershell
npm run migration:revert
```

Quy tắc: Luôn thiết kế migration hai chiều (`up` và `down`) theo nguyên tắc mở rộng trước, thu hẹp sau để lệnh `migration:revert` không gây mất mát dữ liệu không thể hoàn tác.
