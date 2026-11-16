@echo off
chcp 65001 >nul
title Gỡ Cài Đặt Add-in OrcaX Office Suite
color 0C
cls
echo ====================================================================
echo                   GỠ CÀI ĐẶT ADD-IN ORCAX KHỎI WORD
echo ====================================================================
echo.

tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I /N "WINWORD.EXE" >nul
if "%ERRORLEVEL%"=="0" (
    echo [*] Đang đóng Microsoft Word...
    taskkill /F /IM WINWORD.EXE >nul 2>&1
    timeout /t 2 /nobreak >nul
)

set "STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"
attrib -h -s -r "%STARTUP_DIR%\~$*.*" >nul 2>&1
del /F /A /Q "%STARTUP_DIR%\~$*caX*.*" >nul 2>&1

if exist "%STARTUP_DIR%\OrcaX_Addin.dotm" (
    del /F /Q "%STARTUP_DIR%\OrcaX_Addin.dotm" >nul 2>&1
    echo [OK] Đã gỡ bỏ Add-in OrcaX khỏi Microsoft Word.
) else (
    echo [*] Add-in OrcaX hiện không có trong thư mục STARTUP.
)

for %%V in (14.0 15.0 16.0) do (
    reg delete "HKCU\Software\Microsoft\Office\%%V\Word\Security\Trusted Locations\OrcaXStartup" /f >nul 2>&1
)

echo.
echo ====================================================================
echo [HOÀN TẤT] ĐÃ GỠ BỎ ADD-IN ORCAX THÀNH CÔNG!
echo ====================================================================
echo.
pause
