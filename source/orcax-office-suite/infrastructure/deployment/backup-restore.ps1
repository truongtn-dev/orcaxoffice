param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("backup", "restore")]
    [string]$Action,

    [Parameter(Mandatory=$false)]
    [string]$File
)

$ErrorActionPreference = "Stop"

$HostName = if ($env:DATABASE_HOST) { $env:DATABASE_HOST } else { "localhost" }
$Port = if ($env:DATABASE_PORT) { $env:DATABASE_PORT } else { "5432" }
$User = if ($env:DATABASE_USER) { $env:DATABASE_USER } else { "orcax" }
$Database = if ($env:DATABASE_NAME) { $env:DATABASE_NAME } else { "orcax_suite" }

if ($Action -eq "backup") {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $targetFile = if ($File) { $File } else { "backup_${Database}_${timestamp}.dump" }
    Write-Host "Bắt đầu sao lưu cơ sở dữ liệu $Database vào $targetFile..." -ForegroundColor Cyan
    pg_dump -h $HostName -p $Port -U $User -d $Database -F c -b -v -f $targetFile
    Write-Host "Sao lưu hoàn tất: $targetFile" -ForegroundColor Green
} elseif ($Action -eq "restore") {
    if (-not $File -or -not (Test-Path $File)) {
        Write-Error "Cần chỉ định đường dẫn tệp sao lưu hợp lệ qua tham số -File."
    }
    Write-Host "Bắt đầu phục hồi cơ sở dữ liệu $Database từ $File..." -ForegroundColor Yellow
    pg_restore -h $HostName -p $Port -U $User -d $Database -c -v $File
    Write-Host "Phục hồi hoàn tất thành công." -ForegroundColor Green
}
