@echo off
chcp 65001 >nul
title Gỡ Cài Đặt Add-in OrcaX
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

set "TARGET_FILE=%APPDATA%\Microsoft\Word\STARTUP\OrcaX_Addin.dotm"
if exist "%TARGET_FILE%" (
    del /F /Q "%TARGET_FILE%" >nul 2>&1
    echo [OK] Đã gỡ bỏ Add-in OrcaX khỏi Microsoft Word.
) else (
    echo [*] Add-in OrcaX hiện không có trong thư mục STARTUP.
)

echo.
echo ====================================================================
echo [HOÀN TẤT] ĐÃ GỠ BỎ ADD-IN ORCAX THÀNH CÔNG!
echo ====================================================================
echo.
pause
