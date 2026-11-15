@echo off
chcp 65001 >nul
title Gỡ Cài Đặt OrcaX Office Suite
color 0C
cls
echo ================================================================================
echo                    GỠ CÀI ĐẶT ORCAX OFFICE SUITE
echo ================================================================================
echo.

tasklist /fi "imagename eq winword.exe" 2>nul | find /i "winword.exe" >nul && taskkill /f /im winword.exe >nul 2>&1
tasklist /fi "imagename eq excel.exe" 2>nul | find /i "excel.exe" >nul && taskkill /f /im excel.exe >nul 2>&1

set "WORD_STARTUP=%APPDATA%\Microsoft\Word\STARTUP"
set "EXCEL_XLSTART=%APPDATA%\Microsoft\Excel\XLSTART"

if exist "%WORD_STARTUP%\OrcaX_Addin.dotm" (
    del /f /q "%WORD_STARTUP%\OrcaX_Addin.dotm" >nul 2>&1
    echo [+] Đã gỡ bỏ Add-in OrcaX khỏi Microsoft Word.
)

if exist "%EXCEL_XLSTART%\OrcaX_Excel.xlam" (
    del /f /q "%EXCEL_XLSTART%\OrcaX_Excel.xlam" >nul 2>&1
    echo [+] Đã gỡ bỏ Add-in OrcaX khỏi Microsoft Excel.
)

echo.
echo ================================================================================
echo                 ĐÃ GỠ CÀI ĐẶT ORCAX THÀNH CÔNG!
echo ================================================================================
echo.
pause
