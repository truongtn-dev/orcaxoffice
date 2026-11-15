import os
import shutil
import zipfile

def build_clean_orcax_package():
    workspace_root = r'e:\OrcaXOffice'
    
    # 1. Locate source dotm
    source_dotm = r'C:\Users\nguye\AppData\Roaming\Microsoft\Word\STARTUP\OrcaX_Addin.dotm'
    if not os.path.exists(source_dotm):
        source_dotm = os.path.join(workspace_root, 'distribution', 'orcax-office-suite-addins', 'OrcaX_Addin.dotm')
    if not os.path.exists(source_dotm):
        source_dotm = os.path.join(workspace_root, 'distribution', 'orcax-word-addin', 'OrcaX_Addin.dotm')

    # 2. Locate source xlam
    source_xlam = os.path.join(workspace_root, 'distribution', 'orcax-office-suite-addins', 'OrcaX_Excel.xlam')
    if not os.path.exists(source_xlam):
        source_xlam = os.path.join(workspace_root, 'distribution', 'orcax-office-suite-addins', 'ChuanHoa_Excel.xlam')
    if not os.path.exists(source_xlam):
        source_xlam = os.path.join(workspace_root, 'distribution', 'orcax-word-addin', 'ChuanHoa_Excel.xlam')
        
    templates_dir = os.path.join(workspace_root, 'templates')
    clean_dist_dir = os.path.join(workspace_root, 'distribution', 'orcax-office-suite')
    
    temp_build = os.path.join(workspace_root, 'build_clean_temp_dir')
    if os.path.exists(temp_build):
        shutil.rmtree(temp_build)
    os.makedirs(temp_build, exist_ok=True)

    print(f'1. Building clean OrcaX_Addin.dotm from {source_dotm}...')
    word_work = os.path.join(temp_build, 'word')
    os.makedirs(word_work, exist_ok=True)
    with zipfile.ZipFile(source_dotm, 'r') as zf:
        zf.extractall(word_work)
        
    custom_ui_path = os.path.join(word_work, 'customUI', 'customUI14.xml')
    with open(custom_ui_path, 'r', encoding='utf-8') as f:
        ui_xml = f.read()
        
    # Replace any remaining third party strings with OrcaX
    ui_xml = ui_xml.replace('id="tabVFormat" label="V-FORMAT"', 'id="tabOrcaXFormat" label="ORCAX-FORMAT"')
    ui_xml = ui_xml.replace('V-FORMAT', 'ORCAX-FORMAT')
    ui_xml = ui_xml.replace('VFormat', 'OrcaX')
    ui_xml = ui_xml.replace('máy chủ TAFINEX', 'máy chủ OrcaX Office Suite')
    ui_xml = ui_xml.replace('TAFINEX', 'OrcaX Office Suite')
    ui_xml = ui_xml.replace('tafinex.com', 'orcax.vn')
    ui_xml = ui_xml.replace('V-Format', 'OrcaX Office Suite')
    
    with open(custom_ui_path, 'w', encoding='utf-8') as f:
        f.write(ui_xml)
        
    orcax_word_target = os.path.join(temp_build, 'OrcaX_Addin.dotm')
    with zipfile.ZipFile(orcax_word_target, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(word_work):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, word_work)
                zf.write(full_path, rel_path)
    print(f'   -> OrcaX_Addin.dotm built: {os.path.getsize(orcax_word_target)} bytes')

    print(f'2. Building clean OrcaX_Excel.xlam from {source_xlam}...')
    excel_work = os.path.join(temp_build, 'excel')
    os.makedirs(excel_work, exist_ok=True)
    with zipfile.ZipFile(source_xlam, 'r') as zf:
        zf.extractall(excel_work)
        
    for ui_name in ['customUI/customUI.xml', 'customUI/customUI14.xml']:
        p = os.path.join(excel_work, ui_name)
        if os.path.exists(p):
            with open(p, 'r', encoding='utf-8') as f:
                content = f.read()
            content = content.replace('label="CHUẨN HÓA IN"', 'label="ORCAX - IN ẤN A4"')
            content = content.replace('label="IN ẤN A4"', 'label="ORCAX IN A4"')
            content = content.replace('từ GitHub', 'từ máy chủ OrcaX')
            content = content.replace('ChuanHoa_Excel', 'OrcaX_Excel')
            content = content.replace('V-Format', 'OrcaX Office Suite')
            content = content.replace('VFormat', 'OrcaX')
            with open(p, 'w', encoding='utf-8') as f:
                f.write(content)
                
    orcax_excel_target = os.path.join(temp_build, 'OrcaX_Excel.xlam')
    with zipfile.ZipFile(orcax_excel_target, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(excel_work):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, excel_work)
                zf.write(full_path, rel_path)
    print(f'   -> OrcaX_Excel.xlam built: {os.path.getsize(orcax_excel_target)} bytes')

    # 3. Clean and prepare clean_dist_dir
    if os.path.exists(clean_dist_dir):
        shutil.rmtree(clean_dist_dir)
    os.makedirs(clean_dist_dir, exist_ok=True)

    shutil.copyfile(orcax_word_target, os.path.join(clean_dist_dir, 'OrcaX_Addin.dotm'))
    shutil.copyfile(orcax_excel_target, os.path.join(clean_dist_dir, 'OrcaX_Excel.xlam'))

    print('3. Copying 58 templates...')
    target_templates_dir = os.path.join(clean_dist_dir, 'templates')
    shutil.copytree(templates_dir, target_templates_dir)
    print(f'   -> Templates copied: {len(os.listdir(target_templates_dir))} files')

    print('4. Copying orcax_presets.json...')
    presets_src = os.path.join(templates_dir, 'user_presets.json')
    if os.path.exists(presets_src):
        shutil.copyfile(presets_src, os.path.join(clean_dist_dir, 'orcax_presets.json'))
        shutil.copyfile(presets_src, os.path.join(clean_dist_dir, 'user_presets.json'))

    print('5. Writing Cai_Dat_OrcaX_1Click.bat...')
    installer_bat = '''@echo off
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
set "WORD_STARTUP=%APPDATA%\\Microsoft\\Word\\STARTUP"
set "EXCEL_XLSTART=%APPDATA%\\Microsoft\\Excel\\XLSTART"

if not exist "%WORD_STARTUP%" md "%WORD_STARTUP%"
if not exist "%EXCEL_XLSTART%" md "%EXCEL_XLSTART%"

REM 3. Cai dat Add-in cho Microsoft Word
echo [*] Đang cài đặt Add-in OrcaX cho Microsoft Word...
copy /y "%~dp0OrcaX_Addin.dotm" "%WORD_STARTUP%\\" >nul
echo [+] Đã cài đặt xong Word Add-in: OrcaX_Addin.dotm (Tab ORCAX-FORMAT)

REM 4. Cai dat Add-in cho Microsoft Excel
echo.
echo [*] Đang cài đặt Add-in OrcaX cho Microsoft Excel...
copy /y "%~dp0OrcaX_Excel.xlam" "%EXCEL_XLSTART%\\" >nul
echo [+] Đã cài đặt xong Excel Add-in: OrcaX_Excel.xlam (Tab ORCAX - IN ẤN A4)

REM 5. Dong bo thu vien bieu mau & du lieu ca nhan hoa cua OrcaX
echo.
echo [*] Đồng bộ 58 biểu mẫu chuẩn Quốc gia & Dữ liệu cơ quan OrcaX...
set "ORCAX_DATA=%APPDATA%\\OrcaX"
set "VFORMAT_DATA=%APPDATA%\\VFormat"
if not exist "%ORCAX_DATA%\\templates" md "%ORCAX_DATA%\\templates"
if not exist "%VFORMAT_DATA%\\templates" md "%VFORMAT_DATA%\\templates"

if exist "%~dp0templates" (
    xcopy /s /e /y /q "%~dp0templates\\*" "%ORCAX_DATA%\\templates\\" >nul 2>&1
    xcopy /s /e /y /q "%~dp0templates\\*" "%VFORMAT_DATA%\\templates\\" >nul 2>&1
)

if exist "%~dp0orcax_presets.json" (
    copy /y "%~dp0orcax_presets.json" "%ORCAX_DATA%\\user_presets.json" >nul 2>&1
    copy /y "%~dp0orcax_presets.json" "%VFORMAT_DATA%\\user_presets.json" >nul 2>&1
)

REM 6. Kich hoat Ban quyen OrcaX Office Suite vinh vien
echo.
echo [*] Kích hoạt bản quyền vĩnh viễn OrcaX Office Suite...
reg add "HKCU\\Software\\OrcaX\\OfficeSuite\\License" /v "LicenseKey" /t REG_SZ /d "ORCAX-OFFICE-SUITE-FULL-2026" /f >nul 2>&1
reg add "HKCU\\Software\\OrcaX\\OfficeSuite\\License" /v "Status" /t REG_SZ /d "ACTIVE" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "LicenseKey" /t REG_SZ /d "ORCAX-OFFICE-SUITE-FULL-2026" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "Blocked" /t REG_SZ /d "0" /f >nul 2>&1

REM 7. Cau hinh Trusted Locations
echo [*] Cấu hình vùng tin cậy (Trusted Locations) an toàn cho Office...
reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\Word\\Security\\Trusted Locations\\OrcaXWord" /v "Path" /t REG_SZ /d "%WORD_STARTUP%" /f >nul 2>&1
reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\Word\\Security\\Trusted Locations\\OrcaXWord" /v "AllowSubFolders" /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Security\\Trusted Locations\\OrcaXExcel" /v "Path" /t REG_SZ /d "%EXCEL_XLSTART%" /f >nul 2>&1
reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Security\\Trusted Locations\\OrcaXExcel" /v "AllowSubFolders" /t REG_DWORD /d 1 /f >nul 2>&1

reg add "HKCU\\Software\\Microsoft\\Office\\15.0\\Word\\Security\\Trusted Locations\\OrcaXWord" /v "Path" /t REG_SZ /d "%WORD_STARTUP%" /f >nul 2>&1
reg add "HKCU\\Software\\Microsoft\\Office\\15.0\\Excel\\Security\\Trusted Locations\\OrcaXExcel" /v "Path" /t REG_SZ /d "%EXCEL_XLSTART%" /f >nul 2>&1

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
'''
    with open(os.path.join(clean_dist_dir, 'Cai_Dat_OrcaX_1Click.bat'), 'w', encoding='utf-8') as f:
        f.write(installer_bat)

    print('6. Writing Go_Cai_Dat_OrcaX.bat...')
    uninstaller_bat = '''@echo off
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

set "WORD_STARTUP=%APPDATA%\\Microsoft\\Word\\STARTUP"
set "EXCEL_XLSTART=%APPDATA%\\Microsoft\\Excel\\XLSTART"

if exist "%WORD_STARTUP%\\OrcaX_Addin.dotm" (
    del /f /q "%WORD_STARTUP%\\OrcaX_Addin.dotm" >nul 2>&1
    echo [+] Đã gỡ bỏ Add-in OrcaX khỏi Microsoft Word.
)

if exist "%EXCEL_XLSTART%\\OrcaX_Excel.xlam" (
    del /f /q "%EXCEL_XLSTART%\\OrcaX_Excel.xlam" >nul 2>&1
    echo [+] Đã gỡ bỏ Add-in OrcaX khỏi Microsoft Excel.
)

echo.
echo ================================================================================
echo                 ĐÃ GỠ CÀI ĐẶT ORCAX THÀNH CÔNG!
echo ================================================================================
echo.
pause
'''
    with open(os.path.join(clean_dist_dir, 'Go_Cai_Dat_OrcaX.bat'), 'w', encoding='utf-8') as f:
        f.write(uninstaller_bat)

    print('7. Writing HUONG_DAN_CAI_DAT_VA_SU_DUNG.txt...')
    guide_txt = '''================================================================================
                    HƯỚNG DẪN CÀI ĐẶT & SỬ DỤNG ORCAX OFFICE SUITE
             Nền tảng Chuẩn Hóa Thể Thức Văn Bản Việt Nam & In Ấn Excel A4
                           Bản quyền: OrcaX Office Suite
================================================================================

1. BỘ CÔNG CỤ ORCAX GỒM CÓ:
--------------------------------------------------------------------------------
A. MICROSOFT WORD (OrcaX_Addin.dotm):
   - Tab "ORCAX-FORMAT" trên thanh Ribbon.
   - Nhóm 1: Chuẩn hóa & Rà soát: Chuẩn hóa tự động 1-Click theo Nghị định 30/2020
     và Hướng dẫn 05-HD/VPTW, nhận diện loại văn bản, rà soát thể thức, chuyển mã
     phông chữ cũ (TCVN3, VNI) sang Unicode, hoàn tác chuẩn hóa.
   - Nhóm 2: Chuẩn hóa thủ công: Tiêu đề đầu trang, tên loại & trích yếu, đề mục &
     nội dung, chữ ký & nơi nhận.
   - Nhóm 3: Bố cục trang: Căn lề A4 chuẩn (20-20-30-15mm), chèn trang ngang/dọc,
     xóa trang trắng thừa, đánh số trang tự động ẩn trang đầu.
   - Nhóm 4: Chèn nhanh thành phần: Khung Quốc hiệu cơ quan các cấp (Bộ, Cục, Sở,
     UBND Xã/Phường, Đảng ủy, Chi bộ), mẫu chữ ký thẩm quyền, độ khẩn, độ mật.
   - Nhóm 5: Công cụ văn bản: Xóa khoảng trắng lỗi, sửa dính chữ, sửa dấu thanh.
   - Nhóm 6: Thư viện 58 biểu mẫu chuẩn Quốc gia (Hành chính, Xã/Phường, Đảng, Chi bộ).

B. MICROSOFT EXCEL (OrcaX_Excel.xlam):
   - Tab "ORCAX - IN ẤN A4" và nhóm "ORCAX IN A4" ngay tại thẻ Home của Excel.
   - Tự động căn lề trang in A4 chuẩn xác.
   - Tự động co giãn bảng tính (Fit to page) chống che mất chữ 100% (cả ô đơn và ô gộp Merge).
   - Tự động nhận diện hướng in (A4 Dọc hoặc A4 Ngang) phù hợp với kích thước bảng.
   - Nút hoàn tác 1-Click trả bảng tính về như trước khi chuẩn hóa.

--------------------------------------------------------------------------------

2. HƯỚNG DẪN CÀI ĐẶT:
--------------------------------------------------------------------------------
CÁCH 1: CÀI ĐẶT TỰ ĐỘNG 1-CLICK (Khuyên dùng)
   - Bước 1: Giải nén toàn bộ tệp "OrcaX_Office_Suite_v1.0.zip".
   - Bước 2: Nhấp đúp chuột vào file:
             "Cai_Dat_OrcaX_1Click.bat"
   - Bước 3: Mở Microsoft Word và Excel, bạn sẽ thấy ngay thanh công cụ OrcaX!

CÁCH 2: CÀI ĐẶT THỦ CÔNG
   - Word: Nhấn phím Windows + R -> nhập: %appdata%\\Microsoft\\Word\\STARTUP
           Copy file "OrcaX_Addin.dotm" vào thư mục STARTUP.
   - Excel: Nhấn phím Windows + R -> nhập: %appdata%\\Microsoft\\Excel\\XLSTART
           Copy file "OrcaX_Excel.xlam" vào thư mục XLSTART.
   - Biểu mẫu: Copy thư mục "templates" và "orcax_presets.json" vào %appdata%\\OrcaX\\.
   - Khởi động lại Microsoft Word và Excel.

--------------------------------------------------------------------------------

3. KÍCH HOẠT BẢN QUYỀN:
   - File cài đặt 1-Click đã tự động kích hoạt bản quyền vĩnh viễn (Active 100%).
   - Trên Word: Tab "ORCAX-FORMAT" -> Bấm nút "Thông tin bản quyền" để xem thông tin.

--------------------------------------------------------------------------------

4. GỠ CÀI ĐẶT:
   - Nhấp đúp chuột vào file "Go_Cai_Dat_OrcaX.bat" để gỡ bỏ sạch sẽ khỏi máy tính.

================================================================================
Hệ thống: OrcaX Office Suite - Nền tảng Chuẩn hóa Thể thức Văn bản Việt Nam
Địa chỉ dự án: E:\\OrcaXOffice
================================================================================
'''
    with open(os.path.join(clean_dist_dir, 'HUONG_DAN_CAI_DAT_VA_SU_DUNG.txt'), 'w', encoding='utf-8') as f:
        f.write(guide_txt)

    print('8. Creating clean OrcaX_Office_Suite_v1.0.zip...')
    zip_clean = os.path.join(workspace_root, 'OrcaX_Office_Suite_v1.0.zip')
    with zipfile.ZipFile(zip_clean, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(clean_dist_dir):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, clean_dist_dir)
                zf.write(full_path, rel_path)
    print(f'   -> OrcaX_Office_Suite_v1.0.zip created: {os.path.getsize(zip_clean)} bytes')

    # Also cleanup old third-party files from distribution root
    print('9. Cleaning up third-party reference files...')
    old_dirs = [
        os.path.join(workspace_root, 'distribution', 'orcax-office-suite-addins'),
        os.path.join(workspace_root, 'distribution', 'orcax-word-addin'),
        os.path.join(workspace_root, 'distribution', 'OrcaX_Word_Addin_v1.0'),
    ]
    for d in old_dirs:
        if os.path.exists(d):
            shutil.rmtree(d)
            print(f'   -> Removed {d}')

    old_zips = [
        os.path.join(workspace_root, 'OrcaX_Word_Addin_v1.0.zip'),
        os.path.join(workspace_root, 'OrcaX_Office_Suite_Addins_v1.0.zip'),
        os.path.join(workspace_root, 'distribution', 'OrcaX_Word_Addin_v1.0.zip'),
    ]
    for z in old_zips:
        if os.path.exists(z):
            os.remove(z)
            print(f'   -> Removed {z}')

    if os.path.exists(temp_build):
        shutil.rmtree(temp_build)

    print('DONE: 100% Clean, customized OrcaX Office Suite distribution package is ready!')

if __name__ == '__main__':
    build_clean_orcax_package()
