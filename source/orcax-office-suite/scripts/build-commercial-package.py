import os
import shutil
import zipfile

def build_commercial_customer_package():
    workspace_root = r'e:\OrcaXOffice'
    
    # Target distribution directory for commercial sale
    comm_dir = os.path.join(workspace_root, 'distribution', 'orcax-commercial-pack')
    os.makedirs(comm_dir, exist_ok=True)

    # 1. Use clean patched OrcaX_Addin.dotm from distribution/orcax-office-suite or current
    src_dotm = os.path.join(workspace_root, 'distribution', 'orcax-office-suite', 'OrcaX_Addin.dotm')
    dest_dotm = os.path.join(comm_dir, 'OrcaX_Addin.dotm')
    if os.path.exists(src_dotm):
        shutil.copyfile(src_dotm, dest_dotm)
    print(f'1. OrcaX_Addin.dotm verified ({os.path.getsize(dest_dotm)} bytes)')

    # 2. Write KEY_BAN_QUYEN.reg
    reg_content = '''Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\\Software\\VB and VBA Program Settings\\VFormat\\License]
"LicenseKey"="ORCAX-VIP-BAN-QUYEN-VINH-VIEN-2026"
"Status"="ACTIVE"
"Type"="lifetime"
"CustomerName"="VIP User"
"ExpiresAt"="9999-12-31"
"DaysLeft"="9999"
"Blocked"="0"
"ActivatedAt"="2026-01-01 00:00:00"

[HKEY_CURRENT_USER\\Software\\OrcaX\\OfficeSuite\\License]
"LicenseKey"="ORCAX-VIP-BAN-QUYEN-VINH-VIEN-2026"
"Status"="ACTIVE"
"Type"="lifetime"
"CustomerName"="VIP User"
"ExpiresAt"="9999-12-31"
'''
    with open(os.path.join(comm_dir, 'KEY_BAN_QUYEN.reg'), 'w', encoding='utf-8') as f:
        f.write(reg_content)
    print('2. Created KEY_BAN_QUYEN.reg')

    # 3. Write Cai_Dat_OrcaX_1Click.bat
    installer_bat = '''@echo off
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
set "STARTUP_DIR=%APPDATA%\\Microsoft\\Word\\STARTUP"
if not exist "%STARTUP_DIR%" (
    mkdir "%STARTUP_DIR%"
)

echo [*] Đang cài đặt Add-in OrcaX vào Microsoft Word...
copy /Y "%~dp0OrcaX_Addin.dotm" "%STARTUP_DIR%\\OrcaX_Addin.dotm" >nul

REM 3. Thu muc du lieu bieu mau OrcaX
set "DATA_DIR=%APPDATA%\\OrcaX"
if not exist "%DATA_DIR%" (
    mkdir "%DATA_DIR%"
)
set "VFORMAT_DIR=%APPDATA%\\VFormat"
if not exist "%VFORMAT_DIR%" (
    mkdir "%VFORMAT_DIR%"
)

REM 4. Sao chep 58 bieu mau chuan
if exist "%~dp0templates" (
    echo [*] Đang thiết lập thư viện 58 biểu mẫu chuẩn Quốc gia...
    if not exist "%DATA_DIR%\\templates" mkdir "%DATA_DIR%\\templates"
    if not exist "%VFORMAT_DIR%\\templates" mkdir "%VFORMAT_DIR%\\templates"
    xcopy /E /I /Y /Q "%~dp0templates\\*" "%DATA_DIR%\\templates\\" >nul 2>&1
    xcopy /E /I /Y /Q "%~dp0templates\\*" "%VFORMAT_DIR%\\templates\\" >nul 2>&1
)

REM 5. Tu dong kich hoat ban quyen vinh vien
echo [*] Đang kích hoạt bản quyền vĩnh viễn OrcaX...
if exist "%~dp0KEY_BAN_QUYEN.reg" (
    reg import "%~dp0KEY_BAN_QUYEN.reg" >nul 2>&1
)
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "LicenseKey" /t REG_SZ /d "ORCAX-VIP-BAN-QUYEN-VINH-VIEN-2026" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "Status" /t REG_SZ /d "ACTIVE" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "Type" /t REG_SZ /d "lifetime" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "CustomerName" /t REG_SZ /d "VIP User" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "ExpiresAt" /t REG_SZ /d "9999-12-31" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "DaysLeft" /t REG_SZ /d "9999" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "Blocked" /t REG_SZ /d "0" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "ActivatedAt" /t REG_SZ /d "2026-01-01 00:00:00" /f >nul 2>&1
reg add "HKCU\\Software\\OrcaX\\OfficeSuite\\License" /v "LicenseKey" /t REG_SZ /d "ORCAX-VIP-BAN-QUYEN-VINH-VIEN-2026" /f >nul 2>&1
reg add "HKCU\\Software\\OrcaX\\OfficeSuite\\License" /v "Status" /t REG_SZ /d "ACTIVE" /f >nul 2>&1

REM 6. Cau hinh Trusted Locations
reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\Word\\Security\\Trusted Locations\\OrcaXStartup" /v "Path" /t REG_SZ /d "%STARTUP_DIR%" /f >nul 2>&1
reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\Word\\Security\\Trusted Locations\\OrcaXStartup" /v "AllowSubFolders" /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\\Software\\Microsoft\\Office\\15.0\\Word\\Security\\Trusted Locations\\OrcaXStartup" /v "Path" /t REG_SZ /d "%STARTUP_DIR%" /f >nul 2>&1

echo.
echo ====================================================================
echo [OK] CÀI ĐẶT VÀ KÍCH HOẠT ADD-IN ORCAX THÀNH CÔNG!
echo.
echo Mở Microsoft Word - Tab ORCAX-FORMAT sẽ tự động xuất hiện!
echo ====================================================================
echo.
pause
'''
    with open(os.path.join(comm_dir, 'Cai_Dat_OrcaX_1Click.bat'), 'w', encoding='utf-8') as f:
        f.write(installer_bat)
    print('3. Created Cai_Dat_OrcaX_1Click.bat')

    # 4. Write Go_Cai_Dat_OrcaX.bat
    uninstaller_bat = '''@echo off
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

set "TARGET_FILE=%APPDATA%\\Microsoft\\Word\\STARTUP\\OrcaX_Addin.dotm"
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
'''
    with open(os.path.join(comm_dir, 'Go_Cai_Dat_OrcaX.bat'), 'w', encoding='utf-8') as f:
        f.write(uninstaller_bat)
    print('4. Created Go_Cai_Dat_OrcaX.bat')

    # 5. Write Huong_Dan_Cai_Dat.txt
    guide_content = '''================================================================================
                    HƯỚNG DẪN CÀI ĐẶT & KÍCH HOẠT ORCAX (v1.0)
               Hệ thống Chuẩn Hóa Thể Thức Văn Bản Việt Nam
                           Bản quyền: OrcaX Office (orcax.vn)
================================================================================

1. CÁCH 1 - CÀI ĐẶT TỰ ĐỘNG 1-CLICK (Khuyên dùng):
   - Bước 1: Giải nén toàn bộ file zip này ra thư mục.
   - Bước 2: Nhấp đúp chuột vào file "Cai_Dat_OrcaX_1Click.bat".
   - Bước 3: Mở Microsoft Word, bạn sẽ thấy ngay Tab "ORCAX-FORMAT" trên thanh Ribbon!
   * Lưu ý: File cài đặt 1-Click đã tự động kích hoạt Bản quyền vĩnh viễn và đồng bộ 58 biểu mẫu.

2. CÁCH 2 - CÀI ĐẶT THỦ CÔNG:
   - Bước 1: Nhấn tổ hợp phím Windows + R, nhập: %appdata%\\Microsoft\\Word\\STARTUP và nhấn Enter.
   - Bước 2: Copy file "OrcaX_Addin.dotm" vào thư mục STARTUP vừa mở.
   - Bước 3: Nhấp đúp chuột vào file "KEY_BAN_QUYEN.reg" -> chọn "Yes" -> "OK".
   - Bước 4: Khởi động lại Microsoft Word.

3. KÍCH HOẠT BẢN QUYỀN VĨNH VIỄN:
   - Nhấp đúp chuột vào file "KEY_BAN_QUYEN.reg" -> chọn "Yes" -> "OK".
   - Sau khi mở Word, chọn Tab "ORCAX-FORMAT" -> bấm "Thông tin bản quyền" để xem trạng thái kích hoạt.
   - Trải nghiệm 100% tính năng: Chuẩn hóa Nghị định 30, Hướng dẫn 05 Đảng, Công văn 5512 Giáo án,
     Căn lề A4 chuẩn, Đánh số trang tự động và thư viện 58 biểu mẫu chuẩn Quốc gia.

4. GỠ CÀI ĐẶT:
   - Nhấp đúp chuột vào file "Go_Cai_Dat_OrcaX.bat".

================================================================================
Hỗ trợ khách hàng: OrcaX Office Suite (orcax.vn) - Email: support@orcax.vn
================================================================================
'''
    with open(os.path.join(comm_dir, 'Huong_Dan_Cai_Dat.txt'), 'w', encoding='utf-8') as f:
        f.write(guide_content)
    print('5. Created Huong_Dan_Cai_Dat.txt')

    # 6. Copy templates if not present
    tpl_dest = os.path.join(comm_dir, 'templates')
    if not os.path.exists(tpl_dest):
        tpl_src = os.path.join(workspace_root, 'templates')
        shutil.copytree(tpl_src, tpl_dest)
        print(f'6. Copied templates ({len(os.listdir(tpl_dest))} files)')
    else:
        print(f'6. Templates verified ({len(os.listdir(tpl_dest))} files)')

    # 7. Create ZIP archive for commercial distribution
    zip_target = os.path.join(workspace_root, 'OrcaX_Setup_v1.0.zip')
    with zipfile.ZipFile(zip_target, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(comm_dir):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, comm_dir)
                zf.write(full_path, rel_path)
                
    print(f'\\nSUCCESS: Commercial package created at {comm_dir}')
    print(f'ZIP package ready for customer download: {zip_target} ({os.path.getsize(zip_target)} bytes)')

if __name__ == '__main__':
    build_commercial_customer_package()
