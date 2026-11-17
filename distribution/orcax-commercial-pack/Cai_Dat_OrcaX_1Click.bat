@echo off
chcp 65001 >nul
title Cài Đặt Add-in OrcaX Office Suite 2026 cho Microsoft Word
color 0A
cls
echo ====================================================================
echo             BỘ CÀI ĐẶT TỰ ĐỘNG ADD-IN ORCAX OFFICE SUITE 2026
echo       Chuẩn hóa thể thức văn bản: Nghị định 30, Hướng dẫn 05, NĐ 250
echo                     Bản quyền chính thức: OrcaX Office (orcax.vn)
echo ====================================================================
echo.

REM 1. Đóng Microsoft Word nếu đang mở để tránh khóa tệp tin
echo [*] Đang kiểm tra và đóng Microsoft Word...
taskkill /F /IM WINWORD.EXE >nul 2>&1
ping 127.0.0.1 -n 2 >nul

REM 2. Xác định thư mục STARTUP của Word
set "STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"
if not exist "%STARTUP_DIR%" (
    mkdir "%STARTUP_DIR%"
)

REM Dọn dẹp tệp khóa tạm thời (~$*) và add-in cũ nếu có
attrib -h -s -r "%STARTUP_DIR%\~$*.*" >nul 2>&1
del /F /A /Q "%STARTUP_DIR%\~$*.*" >nul 2>&1
if exist "%STARTUP_DIR%\VFormat_Addin.dotm" (
    del /F /Q "%STARTUP_DIR%\VFormat_Addin.dotm" >nul 2>&1
)
if exist "%STARTUP_DIR%\OrcaX_Addin.dotm" (
    del /F /Q "%STARTUP_DIR%\OrcaX_Addin.dotm" >nul 2>&1
)

REM 3. Sao chép tệp Add-in mới nhất vào STARTUP
echo [*] Đang cài đặt Add-in OrcaX vào Microsoft Word STARTUP...
REM 2b. Xoa cache MSForms.exd cu (nguyen nhan "Compile error in hidden module")
del /F /Q "%TEMP%\VBE\*.exd" >nul 2>&1
del /F /Q "%APPDATA%\Microsoft\Forms\*.exd" >nul 2>&1

copy /Y "%~dp0OrcaX_Addin.dotm" "%STARTUP_DIR%\OrcaX_Addin.dotm" >nul
if not exist "%STARTUP_DIR%\OrcaX_Addin.dotm" (
    echo [LOI] Khong the sao chep OrcaX_Addin.dotm vao STARTUP!
    echo Vui long kiem tra quyen truy cap va thu lai.
    pause
    exit /b 1
)

REM 4. Bỏ chặn hoàn toàn Mark of the Web (Zone.Identifier)
echo [*] Đang mở khóa tệp tin bảo mật Windows (Unblock Mark-of-the-Web)...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -Path '%STARTUP_DIR%' -Recurse | Unblock-File; Remove-Item -Path '%STARTUP_DIR%\OrcaX_Addin.dotm' -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue; Get-ChildItem -Path '%~dp0' -Recurse | Unblock-File; Remove-Item -Path '%~dp0OrcaX_Addin.dotm' -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue" >nul 2>&1

REM 5. Xóa danh mục lỗi bộ nhớ đệm (Resiliency Cache) và Cấu hình Trusted Locations & Macro Security
echo [*] Đang cấu hình Vùng tin cậy bảo mật (Trusted Locations) và bật Macro...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $s = \"$env:APPDATA\Microsoft\Word\STARTUP\\"; foreach ($v in @('14.0','15.0','16.0')) { Remove-Item -Path \"HKCU:\Software\Microsoft\Office\$v\Word\Resiliency\DisabledItems\" -Recurse -Force; Remove-Item -Path \"HKCU:\Software\Microsoft\Office\$v\Word\Resiliency\StartupItems\" -Recurse -Force; $sec = \"HKCU:\Software\Microsoft\Office\$v\Word\Security\"; New-Item -Path $sec -Force | Out-Null; Set-ItemProperty -Path $sec -Name VBAWarnings -Value 1 -Type DWord -Force; Set-ItemProperty -Path $sec -Name AccessVBOM -Value 1 -Type DWord -Force; Set-ItemProperty -Path $sec -Name blockcontentexecutionfrominternet -Value 0 -Type DWord -Force; $tl = \"$sec\Trusted Locations\OrcaXStartup\"; New-Item -Path $tl -Force | Out-Null; Set-ItemProperty -Path $tl -Name Path -Value $s -Type String -Force; Set-ItemProperty -Path $tl -Name AllowSubFolders -Value 1 -Type DWord -Force; Set-ItemProperty -Path $tl -Name Description -Value 'OrcaX Startup Addin' -Type String -Force; }" >nul 2>&1

REM 6. Thiết lập thư mục dữ liệu và sao chép kho biểu mẫu
set "DATA_DIR=%APPDATA%\OrcaX"
set "VFORMAT_DIR=%APPDATA%\VFormat"
if not exist "%DATA_DIR%" mkdir "%DATA_DIR%"
if not exist "%VFORMAT_DIR%" mkdir "%VFORMAT_DIR%"

if exist "%~dp0templates" (
    echo [*] Đang thiết lập kho biểu mẫu chuẩn hành chính và Đảng bộ...
    if not exist "%DATA_DIR%\templates" mkdir "%DATA_DIR%\templates"
    if not exist "%VFORMAT_DIR%\templates" mkdir "%VFORMAT_DIR%\templates"
    xcopy /E /I /Y /Q "%~dp0templates\*" "%DATA_DIR%\templates\" >nul 2>&1
    xcopy /E /I /Y /Q "%~dp0templates\*" "%VFORMAT_DIR%\templates\" >nul 2>&1
)

REM 7. Kích hoạt bản quyền VIP trọn đời
echo [*] Đang kích hoạt bản quyền VIP trọn đời OrcaX Office...
if exist "%~dp0KEY_BAN_QUYEN.reg" (
    reg import "%~dp0KEY_BAN_QUYEN.reg" >nul 2>&1
)
reg add "HKCU\Software\OrcaX\OfficeSuite\License" /f /v "LicenseKey" /t REG_SZ /d "ORCAX-VIP-BAN-QUYEN-VINH-VIEN-2026" >nul 2>&1
reg add "HKCU\Software\OrcaX\OfficeSuite\License" /f /v "Status" /t REG_SZ /d "ACTIVE" >nul 2>&1
reg add "HKCU\Software\OrcaX\OfficeSuite\License" /f /v "Type" /t REG_SZ /d "lifetime" >nul 2>&1
reg add "HKCU\Software\OrcaX\OfficeSuite\License" /f /v "CustomerName" /t REG_SZ /d "VIP User" >nul 2>&1
reg add "HKCU\Software\OrcaX\OfficeSuite\License" /f /v "ExpiresAt" /t REG_SZ /d "9999-12-31" >nul 2>&1

reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /f /v "LicenseKey" /t REG_SZ /d "ORCAX-VIP-BAN-QUYEN-VINH-VIEN-2026" >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /f /v "Status" /t REG_SZ /d "ACTIVE" >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /f /v "Type" /t REG_SZ /d "lifetime" >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /f /v "CustomerName" /t REG_SZ /d "VIP User" >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /f /v "ExpiresAt" /t REG_SZ /d "9999-12-31" >nul 2>&1

echo.
echo ====================================================================
echo [OK] CÀI ĐẶT VÀ CẤU HÌNH ORCAX OFFICE SUITE THÀNH CÔNG!
echo.
echo Mở Microsoft Word - Tab "ORCAX OFFICE" với 7 nhóm chức năng đã sẵn sàng!
echo ====================================================================
echo.
pause
