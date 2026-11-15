@echo off
chcp 65001 >nul
title Cài Đặt OrcaX Office Suite - Nền Tảng Chuẩn Hóa Văn Bản & In Ấn A4
color 0B
cls
echo ================================================================================
echo             CÀI ĐẶT ORCAX OFFICE SUITE (WORD & EXCEL) - 1-CLICK
echo   Hệ thống Chuẩn hóa Thể thức Văn bản (NĐ 30, HD 05) & Chuẩn hóa In ấn Excel A4
echo                      Bản quyền: OrcaX Office Suite
echo ================================================================================
echo.

REM 1. Kiem tra va dong Word & Excel neu dang mo
set "DANG_MO=0"
tasklist /fi "imagename eq winword.exe" 2>nul | find /i "winword.exe" >nul && set "DANG_MO=1"
tasklist /fi "imagename eq excel.exe" 2>nul | find /i "excel.exe" >nul && set "DANG_MO=1"

if "%DANG_MO%"=="1" (
    echo [!] Phát hiện Microsoft Word hoặc Excel đang mở.
    echo     Đang tự động đóng để cập nhật Add-in OrcaX...
    taskkill /f /im winword.exe >nul 2>&1
    taskkill /f /im excel.exe >nul 2>&1
    timeout /t 2 >nul
)

REM 2. Thu muc STARTUP Word & XLSTART Excel
set "WORD_STARTUP=%APPDATA%\Microsoft\Word\STARTUP"
set "EXCEL_XLSTART=%APPDATA%\Microsoft\Excel\XLSTART"

if not exist "%WORD_STARTUP%" md "%WORD_STARTUP%"
if not exist "%EXCEL_XLSTART%" md "%EXCEL_XLSTART%"

REM 3. Cai dat Add-in cho Microsoft Word
echo [*] Đang cài đặt Add-in OrcaX cho Microsoft Word...
copy /y "%~dp0OrcaX_Addin.dotm" "%WORD_STARTUP%\" >nul
echo [+] Đã cài đặt xong Word Add-in: OrcaX_Addin.dotm (Tab ORCAX-FORMAT)

REM 4. Cai dat Add-in cho Microsoft Excel
echo.
echo [*] Đang cài đặt Add-in OrcaX cho Microsoft Excel...
copy /y "%~dp0OrcaX_Excel.xlam" "%EXCEL_XLSTART%\" >nul
echo [+] Đã cài đặt xong Excel Add-in: OrcaX_Excel.xlam (Tab ORCAX - IN ẤN A4)

REM 5. Dong bo thu vien bieu mau & du lieu ca nhan hoa cua OrcaX
echo.
echo [*] Đồng bộ 58 biểu mẫu chuẩn Quốc gia & Dữ liệu cơ quan OrcaX...
set "ORCAX_DATA=%APPDATA%\OrcaX"
set "VFORMAT_DATA=%APPDATA%\VFormat"
if not exist "%ORCAX_DATA%\templates" md "%ORCAX_DATA%\templates"
if not exist "%VFORMAT_DATA%\templates" md "%VFORMAT_DATA%\templates"

if exist "%~dp0templates" (
    xcopy /s /e /y /q "%~dp0templates\*" "%ORCAX_DATA%\templates\" >nul 2>&1
    xcopy /s /e /y /q "%~dp0templates\*" "%VFORMAT_DATA%\templates\" >nul 2>&1
)

if exist "%~dp0orcax_presets.json" (
    copy /y "%~dp0orcax_presets.json" "%ORCAX_DATA%\user_presets.json" >nul 2>&1
    copy /y "%~dp0orcax_presets.json" "%VFORMAT_DATA%\user_presets.json" >nul 2>&1
)

REM 6. Kich hoat Ban quyen OrcaX Office Suite vinh vien
echo.
echo [*] Kích hoạt bản quyền vĩnh viễn OrcaX Office Suite...
reg add "HKCU\Software\OrcaX\OfficeSuite\License" /v "LicenseKey" /t REG_SZ /d "ORCAX-OFFICE-SUITE-FULL-2026" /f >nul 2>&1
reg add "HKCU\Software\OrcaX\OfficeSuite\License" /v "Status" /t REG_SZ /d "ACTIVE" /f >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /v "LicenseKey" /t REG_SZ /d "ORCAX-OFFICE-SUITE-FULL-2026" /f >nul 2>&1
reg add "HKCU\Software\VB and VBA Program Settings\VFormat\License" /v "Blocked" /t REG_SZ /d "0" /f >nul 2>&1

REM 7. Cau hinh Trusted Locations
echo [*] Cấu hình vùng tin cậy (Trusted Locations) an toàn cho Office...
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\OrcaXWord" /v "Path" /t REG_SZ /d "%WORD_STARTUP%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\OrcaXWord" /v "AllowSubFolders" /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\OrcaXExcel" /v "Path" /t REG_SZ /d "%EXCEL_XLSTART%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\OrcaXExcel" /v "AllowSubFolders" /t REG_DWORD /d 1 /f >nul 2>&1

reg add "HKCU\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\OrcaXWord" /v "Path" /t REG_SZ /d "%WORD_STARTUP%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\15.0\Excel\Security\Trusted Locations\OrcaXExcel" /v "Path" /t REG_SZ /d "%EXCEL_XLSTART%" /f >nul 2>&1

echo.
echo ================================================================================
echo             CÀI ĐẶT THÀNH CÔNG TRỌN BỘ ORCAX OFFICE SUITE!
echo ================================================================================
echo.
echo  1. MICROSOFT WORD:
echo     - Mở Word, Tab "ORCAX-FORMAT" sẽ xuất hiện ngay trên thanh Ribbon.
echo     - Đầy đủ tính năng Chuẩn hóa tự động NĐ 30, HD 05, 58 biểu mẫu chuẩn Quốc gia.
echo.
echo  2. MICROSOFT EXCEL:
echo     - Mở Excel, Tab "ORCAX - IN ẤN A4" và nhóm "ORCAX IN A4" xuất hiện ở thẻ Home.
echo     - Chuẩn hóa trang in A4, co giãn chống mất chữ, tự động nhận diện A4 Dọc/Ngang.
echo.
echo  3. BẢN QUYỀN:
echo     - Bản quyền OrcaX Office Suite đã được kích hoạt vĩnh viễn (Active 100%).
echo.
echo ================================================================================
pause
