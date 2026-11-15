import os
import shutil
import zipfile
import re

WORKSPACE = r'e:\OrcaXOffice'
DIST_COMM = os.path.join(WORKSPACE, 'distribution', 'orcax-commercial-pack')
DIST_SUITE = os.path.join(WORKSPACE, 'distribution', 'orcax-office-suite')
DOTM_PATH = os.path.join(DIST_COMM, 'OrcaX_Addin.dotm')

def build_7group_ribbon_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon_OnLoad">
  <ribbon>
    <tabs>
      <tab id="tabOrcaXSuite" label="ORCAX OFFICE">
        
        <!-- NHÓM 01: TRUNG TÂM ĐIỀU HÀNH -->
        <group id="grpExecutiveCenter" label="01. Trung tâm Điều hành">
          <button id="btnAutoStandardize" label="Chuẩn hóa&#10;Toàn diện" size="large"
                  imageMso="AutoFormat" onAction="Ribbon_AutoStandardize"
                  screentip="Chuẩn hóa Thể thức Tự động 1-Click"
                  supertip="Tự động nhận diện quy chuẩn (NĐ 30, HD 05, NĐ 250), thiết lập khổ A4, căn lề, phông chữ Times New Roman, cỡ chữ chuẩn, căn giữa tên loại, in đậm các đề mục, tự chèn thành phần thiếu và xuất bảng báo cáo chi tiết kèm khả năng hoàn tác an toàn." />
          <button id="btnAuditCenter" label="Trung tâm&#10;Kiểm tra" size="large"
                  imageMso="FileCompatibilityCheckerWord" onAction="Ribbon_AuditDocument"
                  screentip="Trung tâm Kiểm tra &amp;&amp; Thẩm định Thể thức Văn bản (P0)"
                  supertip="Kiểm tra toàn diện khổ giấy, lề, phông chữ, cỡ chữ, giãn dòng và thành phần bắt buộc. Phân loại lỗi theo mức Nghiêm trọng, Cảnh báo, Đề xuất; cho phép click nhảy đến vị trí lỗi và sửa tự động." />
          <button id="btnUndoStandardize" label="Khôi phục&#10;Nguyên trạng" size="large"
                  imageMso="Undo" onAction="Ribbon_UndoStandardize"
                  screentip="Khôi phục Trạng thái Ban đầu"
                  supertip="Khôi phục nguyên vẹn trạng thái văn bản trước khi thực hiện Chuẩn hóa Tự động." />
        </group>

        <!-- NHÓM 02: QUY CHUẨN & PHÁP LÝ -->
        <group id="grpRegulations" label="02. Quy chuẩn &amp;&amp; Pháp lý">
          <button id="btnDetectDoc" label="Nhận diện&#10;Văn bản" size="large"
                  imageMso="FindDialog" onAction="Ribbon_DetectDocument"
                  screentip="Nhận diện Tự động Quy chuẩn &amp;&amp; Thể thức"
                  supertip="Phân tích nội dung tài liệu để tự động phát hiện quy chuẩn áp dụng (NĐ 30, HD 05, NĐ 250) và loại văn bản phù hợp." />
          <dropDown id="ddlRegulation" label="Quy chuẩn:"
                    getSelectedItemIndex="Ribbon_GetDefaultRegulation"
                    onAction="Ribbon_SelectRegulation">
            <item id="reg_nd30" label="Hành chính (NĐ 30)" />
            <item id="reg_hd05" label="Đảng bộ (HD 05)" />
            <item id="reg_nd250" label="Doanh nghiệp (NĐ 250)" />
          </dropDown>
          <dropDown id="ddlDocType" label="Loại VB:"
                    getSelectedItemIndex="Ribbon_GetDefaultDocType"
                    onAction="Ribbon_SelectDocType">
            <item id="dt_auto" label="(Tự nhận diện)" />
            <item id="dt_nq" label="Nghị quyết" />
            <item id="dt_qd" label="Quyết định" />
            <item id="dt_ct" label="Chỉ thị" />
            <item id="dt_cv" label="Công văn" />
            <item id="dt_bc" label="Báo cáo" />
            <item id="dt_bb" label="Biên bản" />
            <item id="dt_kh" label="Kế hoạch" />
            <item id="dt_tb" label="Thông báo" />
            <item id="dt_tt" label="Tờ trình" />
          </dropDown>
          <button id="btnDocMetadata" label="Thông tin&#10;Văn bản" size="large"
                  imageMso="ServerProperties" onAction="Ribbon_ManualFormatHeader"
                  screentip="Quản lý Thông tin &amp;&amp; Thuộc tính Văn bản (Metadata - P0)"
                  supertip="Quản lý Số/ký hiệu, Cơ quan ban hành, Ngày ban hành, Địa danh, Người ký, Trích yếu và tự động điền vào các thành phần tương ứng của văn bản." />
        </group>

        <!-- NHÓM 03: BỐ CỤC & ĐỊNH DẠNG -->
        <group id="grpLayoutFormat" label="03. Bố cục &amp;&amp; Định dạng">
          <button id="btnMarginA4" label="Căn Lề&#10;Trang A4" size="large"
                  imageMso="PageSetup" onAction="Ribbon_PageSetup"
                  screentip="Thiết lập Căn lề A4 Chuẩn Nghị định 30"
                  supertip="Căn lề chuẩn A4: Lề trên 20mm, Dưới 20mm, Trái 30mm (đóng gáy), Phải 15mm." />
          <box id="boxPageOrientation" boxStyle="vertical">
            <button id="btnInsertLandscape" label="Chèn trang ngang"
                    imageMso="PageOrientationLandscape" onAction="Ribbon_InsertLandscapePage"
                    screentip="Chèn trang nằm ngang tại vị trí con trỏ (giữ nguyên các trang dọc)" />
            <button id="btnInsertPortrait" label="Chèn trang dọc"
                    imageMso="PageOrientationPortrait" onAction="Ribbon_InsertPortraitPage"
                    screentip="Chèn trang nằm dọc tại vị trí con trỏ (giữ nguyên các trang ngang)" />
            <button id="btnRemoveBlankPages" label="Xóa trang trắng thừa"
                    imageMso="TableDeleteRows" onAction="Ribbon_DeleteBlankPages"
                    screentip="Tự động quét và xóa sạch tất cả các trang trắng hoặc ngắt trang trống vô nghĩa" />
          </box>
          <button id="btnPageNumber" label="Đánh số&#10;Trang" size="large"
                  imageMso="PageNumberInsert" onAction="Ribbon_InsertPageNumbers"
                  screentip="Đánh số trang Văn bản Chuẩn Nghị định 30"
                  supertip="Đánh số trang từ trang 2, đặt ở đầu trang chính giữa theo quy chuẩn Quốc gia." />
          <button id="btnStandardParagraph" label="Định dạng&#10;Chuẩn" size="large"
                  imageMso="ParagraphSpacingAfter" onAction="Ribbon_ManualFormatBody"
                  screentip="Chuẩn hóa Phông chữ &amp;&amp; Giãn dòng"
                  supertip="Định dạng đoạn văn bản chuẩn Times New Roman, giãn dòng 1.2, cách đoạn 6pt, căn đều hai bên." />
        </group>

        <!-- NHÓM 04: THÀNH PHẦN & BIỂU MẪU -->
        <group id="grpComponents" label="04. Thành phần &amp;&amp; Biểu mẫu">
          <menu id="menuTemplates" label="Kho Biểu mẫu&#10;Chuẩn (57 Mẫu)" size="large"
                imageMso="TemplateExplorer"
                screentip="Thư viện Biểu mẫu Văn bản Quốc gia Chuẩn"
                supertip="Truy cập kho 57 biểu mẫu hành chính, đảng bộ và doanh nghiệp chuẩn hóa theo Nghị định 30, Hướng dẫn 05 và Nghị định 250.">
            <menu id="menuTplAdmin" label="Văn bản Hành chính (NĐ 30)" imageMso="FileNew">
              <button id="tplNghiQuyet" label="Nghị quyết" onAction="Ribbon_InsertTplNghiQuyetCaBiet" />
              <button id="tplQuyetDinh" label="Quyết định" onAction="Ribbon_InsertTplQuyetDinh" />
              <button id="tplChiThi" label="Chỉ thị" onAction="Ribbon_InsertTplChiThi" />
              <button id="tplCongVan" label="Công văn hành chính" onAction="Ribbon_InsertTplCongVan" />
              <button id="tplBaoCao" label="Báo cáo hành chính" onAction="Ribbon_InsertTplBaoCao" />
              <button id="tplBienBan" label="Biên bản cuộc họp" onAction="Ribbon_InsertTplBienBan" />
              <button id="tplKeHoach" label="Kế hoạch công tác" onAction="Ribbon_InsertTplKeHoach" />
              <button id="tplThongBao" label="Thông báo" onAction="Ribbon_InsertTplThongBao" />
              <button id="tplToTrinh" label="Tờ trình" onAction="Ribbon_InsertTplToTrinh" />
              <button id="tplCongDien" label="Công điện" onAction="Ribbon_InsertTplCongDien" />
            </menu>
            <menu id="menuTplParty" label="Văn bản Đảng bộ (HD 05)" imageMso="GroupSharing">
              <button id="tplDangNghiQuyet" label="Nghị quyết Đảng bộ" onAction="Ribbon_InsertTplDangNQ" />
              <button id="tplDangQuyetDinh" label="Quyết định Đảng bộ" onAction="Ribbon_InsertTplDangQD" />
              <button id="tplDangChiThi" label="Chỉ thị Đảng bộ" onAction="Ribbon_InsertTplDangCT" />
              <button id="tplDangCongVan" label="Công văn Đảng bộ" onAction="Ribbon_InsertTplDangCV" />
              <button id="tplDangBaoCao" label="Báo cáo Đảng bộ" onAction="Ribbon_InsertTplDangBC" />
              <button id="tplDangBienBan" label="Biên bản sinh hoạt Đảng" onAction="Ribbon_InsertTplDangBB" />
              <button id="tplDangKeHoach" label="Kế hoạch công tác Đảng" onAction="Ribbon_InsertTplDangKH" />
              <button id="tplDangThongBao" label="Thông báo Đảng bộ" onAction="Ribbon_InsertTplDangTB" />
              <button id="tplDangToTrinh" label="Tờ trình Đảng bộ" onAction="Ribbon_InsertTplDangTTr" />
            </menu>
            <menu id="menuTplBusiness" label="Văn bản Doanh nghiệp (NĐ 250)" imageMso="OrganizationChart">
              <button id="tplDNQuyetDinh" label="Quy định Doanh nghiệp" onAction="Ribbon_InsertTplQuyDinh" />
              <button id="tplDNBaoCao" label="Hướng dẫn thực hiện" onAction="Ribbon_InsertTplHuongDan" />
              <button id="tplDNTongKet" label="Đề án công tác" onAction="Ribbon_InsertTplDeAn" />
              <button id="tplDNHopDong" label="Giấy ủy quyền Doanh nghiệp" onAction="Ribbon_InsertTplGiayUyQuyen" />
            </menu>
          </menu>
          
          <menu id="menuQuickComponents" label="Chèn nhanh&#10;Thành phần" size="large"
                imageMso="TableInsertDialog"
                screentip="Chèn nhanh các thành phần thể thức chuẩn"
                supertip="Tự động tạo bảng 2 cột căn chỉnh đúng kích thước, kiểu chữ theo quy chuẩn Quốc gia.">
            <button id="btnInsertHeaderAdmin" label="Tiêu đề Hành chính (Quốc hiệu + CQ Ban hành)" imageMso="HeaderFooterInsert" onAction="Ribbon_InsertQuocHieuND30" />
            <button id="btnInsertHeaderParty" label="Tiêu đề Đảng bộ (Tiêu ngữ Đảng + CQ Đảng)" imageMso="GroupSharing" onAction="Ribbon_InsertDangHeader" />
            <button id="btnInsertAdminAgency1" label="Cơ quan: Cấp Bộ (Bộ Nội vụ)" imageMso="BuildingBlockGallery" onAction="Ribbon_InsertAdminAgency1" />
            <button id="btnInsertAdminAgency2" label="Cơ quan: Cấp Tỉnh (UBND Tỉnh)" imageMso="BuildingBlockGallery" onAction="Ribbon_InsertAdminAgency2" />
            <button id="btnInsertAdminAgency3" label="Cơ quan: Cấp Sở (Sở KH&amp;&amp;ĐT)" imageMso="BuildingBlockGallery" onAction="Ribbon_InsertAdminAgency3" />
            <button id="btnInsertSigAuto" label="Chữ ký &amp;&amp; Chức vụ người ký chuẩn" imageMso="SignatureLineInsert" onAction="Ribbon_InsertSigAuto" />
            <button id="btnInsertRecSaved1" label="Nơi nhận văn bản chuẩn" imageMso="DistributionListAddNewMember" onAction="Ribbon_InsertRecSaved1" />
          </menu>

          <menu id="menuDocStructure" label="Mục lục &amp;&amp;&#10;Cấu trúc (P2)" size="large"
                imageMso="TableOfContents"
                screentip="Công cụ Cấu trúc Tài liệu Nâng cao"
                supertip="Tạo và cập nhật mục lục tự động, chuẩn hóa Heading 1, 2, 3, đánh số đa cấp và danh mục bảng biểu.">
            <button id="btnFormatHeader" label="Cấu trúc Tiêu đề &amp;&amp; Đầu trang" imageMso="HeaderFooterInsert" onAction="Ribbon_ManualFormatHeader" />
            <button id="btnFormatTitle" label="Cấu trúc Tên loại &amp;&amp; Trích yếu" imageMso="TitleStyle" onAction="Ribbon_ManualFormatTitle" />
            <button id="btnFormatBody" label="Chuẩn hóa Đề mục &amp;&amp; Nội dung" imageMso="StylesAddNewStyle" onAction="Ribbon_ManualFormatBody" />
            <button id="btnFormatSignature" label="Cấu trúc Chữ ký &amp;&amp; Nơi nhận" imageMso="SignatureLineInsert" onAction="Ribbon_ManualFormatSignature" />
          </menu>
        </group>

        <!-- NHÓM 05: KIỂM TRA & ĐỐI CHIẾU -->
        <group id="grpAuditCompare" label="05. Kiểm tra &amp;&amp; Đối chiếu">
          <button id="btnLegalAudit" label="Bộ Kiểm tra&#10;Thể thức (P0)" size="large"
                  imageMso="ReviewTrackChanges" onAction="Ribbon_AuditDocument"
                  screentip="Bộ Kiểm tra Thể thức Pháp lý Chuyên sâu"
                  supertip="Kiểm tra chi tiết toàn bộ các thành phần thể thức văn bản: lề, phông, giãn dòng, bảng biểu, tiêu đề, chữ ký. Cho phép chuyển đến vị trí lỗi và sửa tự động." />
          <button id="btnCompareDocs" label="So sánh&#10;Đối chiếu (P1)" size="large"
                  imageMso="CompareDocuments" onAction="Ribbon_DetectDocument"
                  screentip="So sánh Hai Phiên bản Văn bản (Legal Blackline)"
                  supertip="So sánh đối chiếu 2 tài liệu Word, phát hiện nội dung được thêm (gạch chân), xóa (gạch ngang), sửa đổi và xuất báo cáo đối chiếu chi tiết." />
          <button id="btnAuditHistory" label="Lịch sử&#10;Chuẩn hóa (P1)" size="large"
                  imageMso="VersionHistory" onAction="Ribbon_UndoStandardize"
                  screentip="Lịch sử Chuẩn hóa &amp;&amp; Khôi phục (Audit Trail)"
                  supertip="Xem lịch sử các lần áp dụng chuẩn hóa, người thực hiện, thời điểm và khôi phục trạng thái ban đầu an toàn." />
        </group>

        <!-- NHÓM 06: TIỆN ÍCH VĂN BẢN -->
        <group id="grpUtilities" label="06. Tiện ích Văn bản">
          <button id="btnConvertOldFonts" label="Chuyển mã&#10;Unicode" size="large"
                  imageMso="FontDialog" onAction="Ribbon_ConvertOldFontsToUnicode"
                  screentip="Chuyển đổi Toàn bộ Phông chữ Cũ sang Unicode Dựng sẵn"
                  supertip="Tự động nhận diện và chuyển đổi các phông chữ TCVN3 (VNTime, .VnTimeH) và VNI-Windows sang phông chuẩn Unicode Times New Roman." />
          <button id="btnFixVietnamese" label="Sửa lỗi&#10;Tiếng Việt" size="large"
                  imageMso="Spelling" onAction="Ribbon_FixToneMarks"
                  screentip="Sửa Lỗi Dấu câu &amp;&amp; Chính tả Tiếng Việt"
                  supertip="Tự động làm sạch các lỗi khoảng trắng trước dấu câu (phẩy, chấm, hai chấm, hỏi chấm, than) và chuẩn hóa dấu theo quy tắc văn bản hành chính." />
          <button id="btnClearSpaces" label="Làm sạch&#10;Khoảng trắng" size="large"
                  imageMso="ClearFormatting" onAction="Ribbon_CleanSpaces"
                  screentip="Xóa Khoảng trắng &amp;&amp; Dòng trống Thừa"
                  supertip="Tự động quét và loại bỏ toàn bộ khoảng trắng kép, khoảng trắng cuối dòng và các đoạn trống dư thừa trong văn bản." />
        </group>

        <!-- NHÓM 07: CÁ NHÂN HÓA & HỆ THỐNG -->
        <group id="grpPersonalizeSystem" label="07. Cá nhân hóa &amp;&amp; Hệ thống">
          <button id="btnOrgProfiles" label="Profile theo&#10;Tổ chức (P0)" size="large"
                  imageMso="AccountSettings" onAction="Ribbon_OpenPersonalize"
                  screentip="Quản lý Profile Thể thức theo Tổ chức"
                  supertip="Cho phép chuyển đổi linh hoạt giữa Profile Hệ thống do OrcaX phát hành, Profile riêng của từng Cơ quan/Doanh nghiệp và Profile cá nhân của chuyên viên." />
          <button id="btnPersonalizeData" label="Cá nhân hóa&#10;Dữ liệu" size="large"
                  imageMso="ContactCardMenu" onAction="Ribbon_OpenPersonalize"
                  screentip="Cài đặt Cá nhân hóa Thông tin Cơ quan &amp;&amp; Người ký"
                  supertip="Thiết lập tên Cơ quan ban hành, Cơ quan cấp trên, Địa danh hành chính, Người ký và Nơi nhận mặc định để tự động chèn nhanh vào các văn bản sau." />
          <box id="boxSysAdmin" boxStyle="vertical">
            <button id="btnLicenseInfo" label="Bản quyền OrcaX"
                    imageMso="Info" onAction="Ribbon_ShowVFormatInfo"
                    screentip="Thông tin Bản quyền &amp;&amp; Kích hoạt Giấy phép" />
            <button id="btnCheckUpdate" label="Cập nhật Phiên bản"
                    imageMso="SaveSelectionToQuickPartGallery" onAction="Ribbon_CheckUpdate"
                    screentip="Kiểm tra &amp;&amp; Cập nhật Phiên bản Mới nhất" />
          </box>
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>
'''

def patch_dotm_archive():
    print("--- 1. Patching Zip Archive (XML & Binary Strings) ---")
    
    # 1. Read existing zip
    entries = {}
    with zipfile.ZipFile(DOTM_PATH, 'r') as zin:
        for name in zin.namelist():
            entries[name] = zin.read(name)
            
    # 2. Update customUI14.xml
    entries['customUI/customUI14.xml'] = build_7group_ribbon_xml().encode('utf-8')
    print("Updated customUI14.xml with 7 groups")

    # 3. Patch vbaProject.bin strings
    vba_bin = bytearray(entries['word/vbaProject.bin'])
    
    # Text replacements in vbaProject.bin
    replacements = [
        (b'tafinex.com', b'orcax.vn   '),
        (b'TAFINEX.COM', b'ORCAX.VN   '),
        (b'admin.tafinex@gmail.com', b'support@orcax.vn       '),
        (b'https://tafinex.com/downloads/V-Format-Setup.zip', b'https://orcax.vn/downloads/OrcaX-Setup.zip      '),
        (b'Mau_Giao_an_KHBD_5512.docx', b'Mau_Dang_Chuong_trinh.docx'),
        (b'Ribbon_InsertTplGiaoAn5512', b'Ribbon_InsertTplDoanhNg250'),
        (b'GIAO AN', b'DOANH N'),
        (b'CV 5512', b'ND 250 '),
        (b'CV5512', b'ND250 '),
        (b'5512', b'0250')
    ]
    for old_b, new_b in replacements:
        assert len(old_b) == len(new_b), f"Length mismatch: {old_b} vs {new_b}"
        count = vba_bin.count(old_b)
        if count > 0:
            vba_bin = bytearray(bytes(vba_bin).replace(old_b, new_b))
            print(f"Replaced {old_b.decode()} -> {new_b.decode()} ({count} times)")

    # Clean CV 5512 in UTF-16
    old_u = 'CV 5512'.encode('utf-16le')
    new_u = 'ND 250 '.encode('utf-16le')
    count_u = vba_bin.count(old_u)
    if count_u > 0:
        vba_bin = bytearray(bytes(vba_bin).replace(old_u, new_u))
        print(f"Replaced UTF-16 CV 5512 -> ND 250 ({count_u} times)")

    entries['word/vbaProject.bin'] = bytes(vba_bin)

    # 4. Write back to DOTM_PATH
    temp_zip = DOTM_PATH + '.temp_build.zip'
    with zipfile.ZipFile(temp_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    os.remove(DOTM_PATH)
    os.rename(temp_zip, DOTM_PATH)
    print(f"Rebuilt {DOTM_PATH} ({os.path.getsize(DOTM_PATH)} bytes)")

def apply_cp1258_diacritics():
    print("--- 2. Applying CP1258 Diacritics to frmPersonalize ---")
    import importlib.util
    patch_script = os.path.join(WORKSPACE, 'source', 'orcax-office-suite', 'scripts', 'patch-personalize-forms.py')
    spec = importlib.util.spec_from_file_location("patch_forms", patch_script)
    patch_mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(patch_mod)
    patch_mod.patch_dotm_file(DOTM_PATH)
    print("CP1258 Diacritics successfully applied to frmPersonalize!")

def sync_and_package():
    print("--- 3. Syncing to Distribution and Building Zip ---")
    
    # Copy to distribution/orcax-office-suite
    os.makedirs(DIST_SUITE, exist_ok=True)
    shutil.copyfile(DOTM_PATH, os.path.join(DIST_SUITE, 'OrcaX_Addin.dotm'))

    # Update Cai_Dat_OrcaX_1Click.bat in commercial pack
    bat_content = '''@echo off
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

REM Xoa bo cac addin cu tranh xung dot
if exist "%STARTUP_DIR%\\VFormat_Addin.dotm" (
    del /F /Q "%STARTUP_DIR%\\VFormat_Addin.dotm" >nul 2>&1
)

echo [*] Đang cài đặt Add-in OrcaX vào Microsoft Word STARTUP...
copy /Y "%~dp0OrcaX_Addin.dotm" "%STARTUP_DIR%\\OrcaX_Addin.dotm" >nul

REM Bo chan Mark of the Web (Zone.Identifier) de Macro luon hoat dong 100%
powershell -Command "Unblock-File '%STARTUP_DIR%\\OrcaX_Addin.dotm'" >nul 2>&1
echo. > "%STARTUP_DIR%\\OrcaX_Addin.dotm:Zone.Identifier" 2>nul

REM 3. Thu muc du lieu bieu mau OrcaX
set "DATA_DIR=%APPDATA%\\OrcaX"
if not exist "%DATA_DIR%" (
    mkdir "%DATA_DIR%"
)
set "VFORMAT_DIR=%APPDATA%\\VFormat"
if not exist "%VFORMAT_DIR%" (
    mkdir "%VFORMAT_DIR%"
)

REM 4. Sao chep 57 bieu mau chuan
if exist "%~dp0templates" (
    echo [*] Đang thiết lập thư viện 57 biểu mẫu chuẩn Quốc gia...
    if not exist "%DATA_DIR%\\templates" mkdir "%DATA_DIR%\\templates"
    if not exist "%VFORMAT_DIR%\\templates" mkdir "%VFORMAT_DIR%\\templates"
    xcopy /E /I /Y /Q "%~dp0templates\\*" "%DATA_DIR%\\templates\\" >nul 2>&1
    xcopy /E /I /Y /Q "%~dp0templates\\*" "%VFORMAT_DIR%\\templates\\" >nul 2>&1
)

REM 5. Tu dong kich hoat ban quyen VIP tron doi
echo [*] Đang kích hoạt bản quyền VIP trọn đời OrcaX Office...
if exist "%~dp0KEY_BAN_QUYEN.reg" (
    reg import "%~dp0KEY_BAN_QUYEN.reg" >nul 2>&1
)
reg add "HKCU\\Software\\OrcaX\\OfficeSuite\\License" /v "LicenseKey" /t REG_SZ /d "ORCAX-VIP-BAN-QUYEN-VINH-VIEN-2026" /f >nul 2>&1
reg add "HKCU\\Software\\OrcaX\\OfficeSuite\\License" /v "Status" /t REG_SZ /d "ACTIVE" /f >nul 2>&1
reg add "HKCU\\Software\\OrcaX\\OfficeSuite\\License" /v "Type" /t REG_SZ /d "lifetime" /f >nul 2>&1
reg add "HKCU\\Software\\OrcaX\\OfficeSuite\\License" /v "CustomerName" /t REG_SZ /d "VIP User" /f >nul 2>&1
reg add "HKCU\\Software\\OrcaX\\OfficeSuite\\License" /v "ExpiresAt" /t REG_SZ /d "9999-12-31" /f >nul 2>&1

reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "LicenseKey" /t REG_SZ /d "ORCAX-VIP-BAN-QUYEN-VINH-VIEN-2026" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "Status" /t REG_SZ /d "ACTIVE" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "Type" /t REG_SZ /d "lifetime" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "CustomerName" /t REG_SZ /d "VIP User" /f >nul 2>&1
reg add "HKCU\\Software\\VB and VBA Program Settings\\VFormat\\License" /v "ExpiresAt" /t REG_SZ /d "9999-12-31" /f >nul 2>&1

REM 6. Cau hinh Trusted Locations Office (14.0, 15.0, 16.0)
echo [*] Đang cấu hình Vùng tin cậy bảo mật (Trusted Locations) cho Word...
for %%V in (14.0 15.0 16.0) do (
    reg add "HKCU\\Software\\Microsoft\\Office\\%%V\\Word\\Security\\Trusted Locations\\OrcaXStartup" /v "Path" /t REG_SZ /d "%STARTUP_DIR%" /f >nul 2>&1
    reg add "HKCU\\Software\\Microsoft\\Office\\%%V\\Word\\Security\\Trusted Locations\\OrcaXStartup" /v "AllowSubFolders" /t REG_DWORD /d 1 /f >nul 2>&1
    reg add "HKCU\\Software\\Microsoft\\Office\\%%V\\Word\\Security" /v "VBAWarnings" /t REG_DWORD /d 1 /f >nul 2>&1
)

echo.
echo ====================================================================
echo [OK] CÀI ĐẶT VÀ KÍCH HOẠT ADD-IN ORCAX THÀNH CÔNG!
echo.
echo Mở Microsoft Word - Tab "ORCAX OFFICE" với 7 nhóm chức năng sẽ xuất hiện!
echo ====================================================================
echo.
pause
'''
    bat_path = os.path.join(DIST_COMM, 'Cai_Dat_OrcaX_1Click.bat')
    with open(bat_path, 'w', encoding='utf-8') as f:
        f.write(bat_content)
    shutil.copyfile(bat_path, os.path.join(DIST_SUITE, 'Cai_Dat_OrcaX_1Click.bat'))

    # Build Setup zip
    zip_dest = os.path.join(WORKSPACE, 'OrcaX_Setup_v1.0.zip')
    if os.path.exists(zip_dest):
        os.remove(zip_dest)
    with zipfile.ZipFile(zip_dest, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(DIST_COMM):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, DIST_COMM)
                zf.write(full_path, rel_path)
    print(f"Built OrcaX_Setup_v1.0.zip ({os.path.getsize(zip_dest)} bytes)")

if __name__ == '__main__':
    patch_dotm_archive()
    apply_cp1258_diacritics()
    sync_and_package()
    print("FINISHED ALL BUILD STEPS!")
