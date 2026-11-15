@echo off
chcp 65001 >nul
title Cài Đặt Add-in OrcaX cho Microsoft Word
color 0A
cls
echo ====================================================================
echo             BỘ CÀI ĐẶT TỰ ĐỘNG ADD-IN ORCAX OFFICE
echo       Chuẩn hóa thể thức văn bản Nghị định 30 và Hướng dẫn 05 VPTW
echo                     Bản quyền: OrcaX Office (orcax.vn)
echo ====================================================================
echo.

REM 1. Kiem tra va dong Microsoft Word neu dang mo
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I /N "WINWORD.EXE" >nul
if "%ERRORLEVEL%"=="0" (
    echo [*] Đang đóng Microsoft Word để cài đặt Add-in...
    taskkill /F /IM WINWORD.EXE >nul 2>&1
    timeout /t 2 /nobreak >nul
)

REM 2. Thu muc STARTUP Word
set "STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"
if not exist "%STARTUP_DIR%" (
    mkdir "%STARTUP_DIR%"
)

echo [*] Đang cài đặt Add-in OrcaX vào Microsoft Word...
copy /Y "%~dp0OrcaX_Addin.dotm" "%STARTUP_DIR%\OrcaX_Addin.dotm" >nul

REM 3. Thu muc du lieu bieu mau OrcaX
set "DATA_DIR=%APPDATA%\OrcaX"
if not exist "%DATA_DIR%" (
    mkdir "%DATA_DIR%"
)
set "VFORMAT_DIR=%APPDATA%\VFormat"
if not exist "%VFORMAT_DIR%" (
    mkdir "%VFORMAT_DIR%"
)

REM 4. Sao chep 58 bieu mau chuan
if exist "%~dp0templates" (
    echo [*] Đang thiết lập thư viện 58 biểu mẫu chuẩn Quốc gia...
    if not exist "%DATA_DIR%\templates" mkdir "%DATA_DIR%\templates"
    if not exist "%VFORMAT_DIR%\templates" mkdir "%VFORMAT_DIR%\templates"
    xcopy /E /I /Y /Q "%~dp0templates\*" "%DATA_DIR%\templates\" >nul 2>&1
    xcopy /E /I /Y /Q "%~dp0templates\*" "%VFORMAT_DIR%\templates\" >nul 2>&1
)

REM 5. Tu dong kich hoat ban quyen vinh vien
echo [*] Đang kích hoạt bản quyền vĩnh viễn OrcaX...
if exist "%~dp0KEY_BAN_QUYEN.reg" (
    reg import "%~dp0KEY_BAN_QUYEN.reg" >nul 2>&1
)
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /v "LicenseKey" /t REG_SZ /d "ORCAX-VIP-BAN-QUYEN-VINH-VIEN-2026" /f >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /v "Status" /t REG_SZ /d "ACTIVE" /f >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /v "Type" /t REG_SZ /d "lifetime" /f >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /v "CustomerName" /t REG_SZ /d "VIP User" /f >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /v "ExpiresAt" /t REG_SZ /d "9999-12-31" /f >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /v "DaysLeft" /t REG_SZ /d "9999" /f >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /v "Blocked" /t REG_SZ /d "0" /f >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /v "ActivatedAt" /t REG_SZ /d "2026-01-01 00:00:00" /f >nul 2>&1
reg add "HKCU\Software\OrcaX\OfficeSuite\License" /v "LicenseKey" /t REG_SZ /d "ORCAX-VIP-BAN-QUYEN-VINH-VIEN-2026" /f >nul 2>&1
reg add "HKCU\Software\OrcaX\OfficeSuite\License" /v "Status" /t REG_SZ /d "ACTIVE" /f >nul 2>&1

REM 6. Cau hinh Trusted Locations
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\OrcaXStartup" /v "Path" /t REG_SZ /d "%STARTUP_DIR%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\OrcaXStartup" /v "AllowSubFolders" /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\OrcaXStartup" /v "Path" /t REG_SZ /d "%STARTUP_DIR%" /f >nul 2>&1

echo.
echo ====================================================================
echo [OK] CÀI ĐẶT VÀ KÍCH HOẠT ADD-IN ORCAX THÀNH CÔNG!
echo.
echo Mở Microsoft Word - Tab ORCAX-FORMAT sẽ tự động xuất hiện!
echo ====================================================================
echo.
pause
