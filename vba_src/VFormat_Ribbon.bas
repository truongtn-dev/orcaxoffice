Attribute VB_Name = "VFormat_Ribbon"

' ==============================================================================

' MODULE: VFormat_Ribbon

' MO TA: CAC HAM CALLBACKS CHO THANH RIBBON ORCAX    & CA NHAN HOA

' 100% PURE ASCII FOR ZERO COMPILE ERRORS

' ==============================================================================

Option Explicit



Private m_RibbonUI As IRibbonUI



' ==============================================================================

' 5 BREAKTHROUGH ENTERPRISE CALLBACKS (ORCAX 2026)

' ==============================================================================

Public Sub Ribbon_AuditWritingStyle(control As IRibbonControl)

    VFormat_Advanced.AuditAdministrativeStyle

End Sub



Public Sub Ribbon_InsertQRCode(control As IRibbonControl)

    VFormat_Advanced.InsertDocumentQRCode

End Sub



Public Sub Ribbon_InsertDigitalSignature(control As IRibbonControl)

    VFormat_Advanced.InsertDigitalSignatureBox

End Sub



Public Sub Ribbon_CompareDrafts(control As IRibbonControl)

    VFormat_Advanced.CompareDraftsAndGenerateReport

End Sub



Public Sub Ribbon_ExportArchivePDF(control As IRibbonControl)

    VFormat_Advanced.ExportStandardArchivePDF

End Sub





Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)

    Set m_RibbonUI = ribbon

    ' OrcaX v2.0: cache ribbon cho dynamic menus (Soan nhanh doc presets)
    OrcaX_Dynamic.CacheRibbon ribbon

    ' VFormat_License.InitLicenseOnStartup - Da tat theo yeu cau

End Sub



Private Function DecodeUTF8(ByVal b64Str As String) As String

    DecodeUTF8 = VFormat_Core.DecodeUTF8(b64Str)

End Function



Private Sub ShowUniMsg(ByVal msgB64OrPlain As String, ByVal titleB64OrPlain As String, Optional ByVal iconType As Long = 64)

    VFormat_Core.ShowUniMsg msgB64OrPlain, titleB64OrPlain, iconType

End Sub



' ==============================================================================

' 1. NHOM 1: NHAN DIEN & RA SOAT

' ==============================================================================

Public Sub Ribbon_AutoStandardize(control As IRibbonControl)

    VFormat_Core.AutoStandardizeDocument

    If Not m_RibbonUI Is Nothing Then

        m_RibbonUI.InvalidateControl "ddRegulation"

        m_RibbonUI.InvalidateControl "ddDocType"

    End If

End Sub



Public Sub Ribbon_DetectDocument(control As IRibbonControl)

    VFormat_Core.ResetRegulationAutoDetect

    VFormat_Core.PerformDocumentDetection

    If Not m_RibbonUI Is Nothing Then

        m_RibbonUI.InvalidateControl "ddRegulation"

        m_RibbonUI.InvalidateControl "ddDocType"

    End If

End Sub



Public Sub Ribbon_SelectRegulation(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    VFormat_Core.SetSelectedRegulation selectedIndex

End Sub



Public Sub Ribbon_GetDefaultRegulation(control As IRibbonControl, ByRef returnedVal)

    returnedVal = VFormat_Core.GetEffectiveRegulation()

End Sub



Public Sub Ribbon_SelectDocType(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    VFormat_Core.g_SelectedDocType = selectedIndex

End Sub



Public Sub Ribbon_GetDefaultDocType(control As IRibbonControl, ByRef returnedVal)

    returnedVal = VFormat_Core.g_SelectedDocType

    If returnedVal = 0 Then

        returnedVal = VFormat_Core.DetectDocType()

        VFormat_Core.g_SelectedDocType = returnedVal

    End If

End Sub



Public Sub Ribbon_AuditDocument(control As IRibbonControl)

    VFormat_Core.AuditDocumentCompliance

End Sub



Public Sub Ribbon_ConvertOldFontsToUnicode(control As IRibbonControl)

    VFormat_Core.ConvertOldFontsToUnicode

End Sub



Public Sub Ribbon_UndoStandardize(control As IRibbonControl)

    VFormat_Core.UndoLastStandardize

End Sub



' ==============================================================================

' 1B. CHUAN HOA THU CONG (MANUAL ACTIONS)

' ==============================================================================

Public Sub Ribbon_ManualFormatHeader(control As IRibbonControl)

    VFormat_Core.ManualFormatHeader

End Sub



Public Sub Ribbon_ManualFormatTitle(control As IRibbonControl)

    VFormat_Core.ManualFormatTitle

End Sub



Public Sub Ribbon_ManualFormatBody(control As IRibbonControl)

    VFormat_Core.ManualFormatBody

End Sub



Public Sub Ribbon_ManualFormatSignature(control As IRibbonControl)

    VFormat_Core.ManualFormatSignature

End Sub



' ==============================================================================

' 2. NHOM 2: BO CUC TRANG

' ==============================================================================

Public Sub Ribbon_PageSetup(control As IRibbonControl)

    VFormat_Core.ApplyPageSetupA4

End Sub



Public Sub Ribbon_InsertLandscapePage(control As IRibbonControl)

    VFormat_Core.InsertLandscapeSection

End Sub



Public Sub Ribbon_InsertPortraitPage(control As IRibbonControl)

    VFormat_Core.InsertPortraitSection

End Sub



Public Sub Ribbon_DeleteBlankPages(control As IRibbonControl)

    VFormat_Core.DeleteBlankPages

End Sub



Public Sub Ribbon_InsertPageNumbers(control As IRibbonControl)

    VFormat_Core.ShowPageNumberDialog

End Sub



' ==============================================================================

' 3. NHOM 3: CHEN THANH PHAN THE THUC

' ==============================================================================

Public Sub Ribbon_InsertQuocHieuAuto(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuAutoAtCursor

End Sub



Public Sub Ribbon_InsertQuocHieuCustom(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuAutoAtCursor

End Sub



Public Sub Ribbon_InsertAdminAgency1(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuAdminIndex 1

End Sub



Public Sub Ribbon_InsertAdminAgency2(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuAdminIndex 2

End Sub



Public Sub Ribbon_InsertAdminAgency3(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuAdminIndex 3

End Sub



Public Sub Ribbon_InsertAdminAgency4(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuAdminIndex 4

End Sub



Public Sub Ribbon_InsertAdminAgency5(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuAdminIndex 5

End Sub



Public Sub Ribbon_InsertAdminAgency6(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuAdminIndex 6

End Sub



Public Sub Ribbon_InsertPartyAgency1(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuPartyIndex 1

End Sub



Public Sub Ribbon_InsertPartyAgency2(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuPartyIndex 2

End Sub



Public Sub Ribbon_InsertPartyAgency3(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuPartyIndex 3

End Sub



Public Sub Ribbon_InsertPartyAgency4(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuPartyIndex 4

End Sub



Public Sub Ribbon_InsertQuocHieuND30(control As IRibbonControl)

    VFormat_Core.InsertQuocHieuND30AtCursor

End Sub



Public Sub Ribbon_InsertPhuLucHeader(control As IRibbonControl)

    VFormat_Core.InsertPhuLucHeaderAtCursor

End Sub



Public Sub Ribbon_InsertBanSaoHeader(control As IRibbonControl)

    VFormat_Core.InsertBanSaoHeaderAtCursor "SAO Y"

End Sub



Public Sub Ribbon_InsertDangHeader(control As IRibbonControl)

    VFormat_Core.InsertDangHeaderAtCursor

End Sub



Public Sub Ribbon_InsertTitleCongVan(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("Vv ..."), DecodeUTF8("S8OtbmggZ+G7rWk6IC4uLg==")

End Sub



Public Sub Ribbon_InsertTitleQuyetDinh(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("UVVZ4bq+VCDEkOG7ik5I"), DecodeUTF8("VuG7gSB2aeG7h2MgLi4u")

End Sub



Public Sub Ribbon_InsertTitleQuyetDinhQuyChe(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("UVVZ4bq+VCDEkOG7ik5I"), DecodeUTF8("QmFuIGjDoG5oIChQaMOqIGR1eeG7h3QpIC4uLg==")

End Sub



Public Sub Ribbon_InsertTitleNghiQuyet(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("TkdI4buKIFFVWeG6vlQ="), DecodeUTF8("VuG7gSB2aeG7h2MgLi4u")

End Sub



Public Sub Ribbon_InsertTitleToTrinh(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("VOG7nCBUUsOMTkg="), DecodeUTF8("VuG7gSB2aeG7h2MgLi4u")

End Sub



Public Sub Ribbon_InsertTitleKeHoach(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("S+G6viBIT+G6oENI"), DecodeUTF8("VGjhu7FjIGhp4buHbiAuLi4=")

End Sub



Public Sub Ribbon_InsertTitleBaoCao(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("QsOBTyBDw4FP"), DecodeUTF8("S+G6v3QgcXXhuqMgLi4u")

End Sub



Public Sub Ribbon_InsertTitleThongBao(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("VEjDlE5HIELDgU8="), DecodeUTF8("VuG7gSB2aeG7h2MgLi4u")

End Sub



Public Sub Ribbon_InsertTitleBienBan(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("QknDik4gQuG6ok4="), DecodeUTF8("SG/huqF0IMSR4buZbmcgLi4u")

End Sub



Public Sub Ribbon_InsertTitleCongDien(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("Q8OUTkcgxJBJ4buGTg=="), DecodeUTF8("VuG7gSB2aeG7h2MgLi4u")

End Sub



Public Sub Ribbon_InsertTitleGiayMoi(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("R0nhuqRZIE3hu5xJ"), DecodeUTF8("ROG7sSBjdeG7mWMgaOG7jXAgLi4u")

End Sub



Public Sub Ribbon_InsertTitleGiayGioiThieu(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("R0nhuqRZIEdJ4buaSSBUSEnhu4ZV"), ""

End Sub



Public Sub Ribbon_InsertTitleGiayNghiPhep(control As IRibbonControl)

    VFormat_Core.InsertDocTitleAtCursor DecodeUTF8("R0nhuqRZIE5HSOG7iCBQSMOJUA=="), ""

End Sub



Public Sub Ribbon_InsertSigAuto(control As IRibbonControl)

    VFormat_Core.InsertSigAutoAtCursor

End Sub



Public Sub Ribbon_InsertSigCustom(control As IRibbonControl)

    VFormat_Core.InsertSigAutoAtCursor

End Sub



Public Sub Ribbon_InsertSigSaved1(control As IRibbonControl)

    VFormat_Core.InsertSigSavedIndex 1, 1

End Sub



Public Sub Ribbon_InsertSigSaved2(control As IRibbonControl)

    VFormat_Core.InsertSigSavedIndex 2, 1

End Sub



Public Sub Ribbon_InsertSigSaved3(control As IRibbonControl)

    VFormat_Core.InsertSigSavedIndex 3, 1

End Sub



Public Sub Ribbon_InsertSigSaved4(control As IRibbonControl)

    VFormat_Core.InsertSigSavedIndex 4, 1

End Sub



Public Sub Ribbon_InsertSigSaved5(control As IRibbonControl)

    VFormat_Core.InsertSigSavedIndex 5, 1

End Sub



Public Sub Ribbon_InsertSigSaved6(control As IRibbonControl)

    VFormat_Core.InsertSigSavedIndex 6, 1

End Sub



Public Sub Ribbon_InsertSigSaved7(control As IRibbonControl)

    VFormat_Core.InsertSigSavedIndex 7, 1

End Sub



Public Sub Ribbon_InsertSigSaved8(control As IRibbonControl)

    VFormat_Core.InsertSigSavedIndex 8, 1

End Sub



Public Sub Ribbon_InsertRecSaved1(control As IRibbonControl)

    VFormat_Core.InsertRecSavedIndex 1

End Sub



Public Sub Ribbon_InsertRecSaved2(control As IRibbonControl)

    VFormat_Core.InsertRecSavedIndex 2

End Sub



Public Sub Ribbon_InsertRecSaved3(control As IRibbonControl)

    VFormat_Core.InsertRecSavedIndex 3

End Sub



Public Sub Ribbon_InsertRecSaved4(control As IRibbonControl)

    VFormat_Core.InsertRecSavedIndex 4

End Sub



Public Sub Ribbon_InsertSigBCH(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSBOaMawIGvDrW5oIGfhu61pOw0KLSBMxrB1IFbEg24gcGjDsm5nLg=="), _

        DecodeUTF8("VE0uIEJBTiBDSOG6pFAgSMOATkgNCkLDjSBUSMav"), _

        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

End Sub



Public Sub Ribbon_InsertSigBTV(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSBOaMawIGvDrW5oIGfhu61pOw0KLSBMxrB1IFbEg24gcGjDsm5nLg=="), _

        DecodeUTF8("VE0uIEJBTiBUSMav4bucTkcgVuG7pA0KUEjDkyBCw40gVEjGrw=="), _

        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBC")

End Sub



Public Sub Ribbon_InsertSigUBND(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSBOaMawIHRyw6puOw0KLSBDaOG7pyB04buLY2gsIFBDVCBVQk5EOw0KLSBMxrB1OiBWVCwgVEgu"), _

        DecodeUTF8("VE0uIOG7plkgQkFOIE5Iw4JOIETDgk4NCkNI4bumIFThu4pDSA=="), _

        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

End Sub



Public Sub Ribbon_InsertSigKT(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSBOaMawIHRyw6puOw0KLSBMxrB1OiBWVCwgSEMu"), _

        DecodeUTF8("S1QuIENI4bumIFThu4pDSA0KUEjDkyBDSOG7piBU4buKQ0g="), _

        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBD")

End Sub



Public Sub Ribbon_InsertSigTL(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSBOaMawIHRyw6puOw0KLSBMxrB1OiBWVC4="), _

        DecodeUTF8("VEwuIENI4bumIFThu4pDSA0KQ0jDgU5IIFbEgk4gUEjDkk5H"), _

        DecodeUTF8("VHLhuqduIFbEg24gRA==")

End Sub



Public Sub Ribbon_InsertSigTUQ(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSBOaMawIHRyw6puOw0KLSBMxrB1OiBWVC4="), _

        DecodeUTF8("VFVRLiBDSOG7piBU4buKQ0gNClRSxq/hu55ORyBQSMOSTkc="), _

        DecodeUTF8("TMOqIFbEg24gRQ==")

End Sub



Public Sub Ribbon_InsertSigQuyen(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSBOaMawIHRyw6puOw0KLSBMxrB1OiBWVC4="), _

        DecodeUTF8("US4gQ0jhu6YgVOG7ikNI"), _

        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBG")

End Sub



Public Sub Ribbon_InsertSigThuTruong(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSBOaMawIHRyw6puOw0KLSBMxrB1OiBWVC4="), _

        DecodeUTF8("R0nDgU0gxJDhu5BD"), _

        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

End Sub



' --- NHOM 4: DO KHAN / DO MAT / CHAN TRANG ---

Public Sub Ribbon_InsertDoKhanHoaToc(control As IRibbonControl)

    VFormat_Core.InsertDoKhanAtCursor DecodeUTF8("SOG7jkEgVOG7kEM=")

End Sub



Public Sub Ribbon_InsertDoKhanThuongKhan(control As IRibbonControl)

    VFormat_Core.InsertDoKhanAtCursor DecodeUTF8("VEjGr+G7ok5HIEtI4bqoTg==")

End Sub



Public Sub Ribbon_InsertDoKhanKhan(control As IRibbonControl)

    VFormat_Core.InsertDoKhanAtCursor DecodeUTF8("S0jhuqhO")

End Sub



Public Sub Ribbon_InsertDoMatMat(control As IRibbonControl)

    VFormat_Core.InsertDoMatAtCursor DecodeUTF8("TeG6rFQ=")

End Sub



Public Sub Ribbon_InsertDoMatToiMat(control As IRibbonControl)

    VFormat_Core.InsertDoMatAtCursor DecodeUTF8("VOG7kEkgTeG6rFQ=")

End Sub



Public Sub Ribbon_InsertDoMatTuyetMat(control As IRibbonControl)

    VFormat_Core.InsertDoMatAtCursor DecodeUTF8("VFVZ4buGVCBN4bqsVA==")

End Sub



Public Sub Ribbon_InsertPhamViLuuHanh(control As IRibbonControl)

    VFormat_Core.InsertPhamViLuuHanhAtCursor DecodeUTF8("TMavVSBIw4BOSCBO4buYSSBC4buY")

End Sub



Public Sub Ribbon_InsertPhamViXemXongTraLai(control As IRibbonControl)

    VFormat_Core.InsertPhamViLuuHanhAtCursor DecodeUTF8("WEVNIFhPTkcgVFLhuqIgTOG6oEk=")

End Sub



Public Sub Ribbon_InsertFooterAgencyInfo(control As IRibbonControl)

    VFormat_Core.InsertFooterAgencyInfoAtCursor

End Sub



' ==============================================================================

' 4. NHOM 4: CONG CU TIENG VIET

' ==============================================================================

Public Sub Ribbon_RemoveToneSelection(control As IRibbonControl)

    VFormat_Core.RemoveTonesFromSelection

End Sub



Public Sub Ribbon_FixToneMarks(control As IRibbonControl)

    VFormat_Core.StandardizeToneMarksAll

End Sub



Public Sub Ribbon_CleanSpaces(control As IRibbonControl)

    VFormat_Core.CleanSpacingErrors

End Sub



' ==============================================================================

' 5. NHOM 5: BIEU MAU, CA NHAN HOA & THONG TIN

' ==============================================================================

Public Sub Ribbon_InsertTplCongVan(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Cong_van_hanh_chinh.docx"

End Sub



Public Sub Ribbon_InsertTplQuyetDinh(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Quyet_dinh.docx"

End Sub



Public Sub Ribbon_InsertTplQuyetDinhQuyChe(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Quyet_dinh_quy_che.docx"

End Sub



Public Sub Ribbon_InsertTplToTrinh(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_To_trinh.docx"

End Sub



Public Sub Ribbon_InsertTplKeHoach(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Ke_hoach.docx"

End Sub



Public Sub Ribbon_InsertTplBaoCao(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Bao_cao.docx"

End Sub



Public Sub Ribbon_InsertTplThongBao(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Thong_bao.docx"

End Sub



Public Sub Ribbon_InsertTplBienBan(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Bien_ban.docx"

End Sub



Public Sub Ribbon_InsertTplCongDien(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Cong_dien.docx"

End Sub



Public Sub Ribbon_InsertTplGiayMoi(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Giay_moi.docx"

End Sub



Public Sub Ribbon_InsertTplGiayGioiThieu(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Giay_gioi_thieu.docx"

End Sub



Public Sub Ribbon_InsertTplGiayNghiPhep(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Giay_nghi_phep.docx"

End Sub



Public Sub Ribbon_InsertTplPhuLuc(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Phu_luc_van_ban.docx"

End Sub



Public Sub Ribbon_InsertTplBanSao(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Ban_sao_van_ban.docx"

End Sub



Public Sub Ribbon_InsertTplDangNQ(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Nghi_quyet.docx"

End Sub



Public Sub Ribbon_InsertTplDangCV(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Cong_van.docx"

End Sub



Public Sub Ribbon_InsertTplVBHN(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Van_ban_hop_nhat_VBHN.docx"

End Sub



Public Sub Ribbon_InsertTplGiaoAn5512(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Giao_an_KHBD_5512.docx"

End Sub



Public Sub Ribbon_OpenPersonalize(control As IRibbonControl)

    VFormat_Core.OpenPersonalizeWindow

End Sub



Public Sub Ribbon_ShowVFormatInfo(control As IRibbonControl)

    VFormat_License.ShowVFormatAppInfo

End Sub



Public Sub Ribbon_CheckUpdate(control As IRibbonControl)

    VFormat_License.CheckUpdateOnline False

End Sub



Public Sub Ribbon_InsertTplDangQD(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Quyet_dinh.docx"

End Sub



Public Sub Ribbon_InsertTplDangKH(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Ke_hoach.docx"

End Sub



Public Sub Ribbon_InsertTplDangBC(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Bao_cao.docx"

End Sub



Public Sub Ribbon_InsertTplDangTB(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Thong_bao.docx"

End Sub



Public Sub Ribbon_InsertTplDangTTr(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_To_trinh.docx"

End Sub



Public Sub Ribbon_SelectOrgLevel(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    VFormat_Core.g_SelectedOrgLevel = selectedIndex

End Sub



Public Sub Ribbon_GetDefaultOrgLevel(control As IRibbonControl, ByRef returnedVal)

    returnedVal = VFormat_Core.g_SelectedOrgLevel

End Sub



Public Sub Ribbon_InsertSigChiBo(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSDEkOG6o25nIOG7p3kgeMOjL3BoxrDhu51uZy4uLiAoxJHhu4MgYi9jKTsNCi0gQ8OhYyDEkeG7k25nIGNow60gxJHhuqNuZyB2acOqbiB0cm9uZyBDaGkgYuG7mTsNCi0gTMawdTogQ2hpYuG7mS4="), _

        DecodeUTF8("VC9NIENISSBC4buYDQpCw41fVEjGrw=="), _

        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

End Sub



Public Sub Ribbon_InsertSigChiBoPho(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSDEkOG6o25nIOG7p3kgeMOjL3BoxrDhu51uZy4uLiAoxJHhu4MgYi9jKTsNCi0gQ8OhYyDEkeG7k25nIGNow60gxJHhuqNuZyB2acOqbiB0cm9uZyBDaGkgYuG7mTsNCi0gTMawdTogQ2hpYuG7mS4="), _

        DecodeUTF8("VC9NIENISSBC4buYDQpQSMOTIELDjSBUSMav"), _

        DecodeUTF8("VHLhuqduIFbEg24gQg==")

End Sub



Public Sub Ribbon_InsertSigDangUyXa(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSBOaMawIGvDrW5oIGfhu61pOw0KLSBUaMaw4budbmcgdHLhu7FjIMSQ4bqjbmcgcGjGsOG7nW5nOw0KLSBMxrB1IFbEg24gcGjDsm5nLg=="), _

        DecodeUTF8("VE0uIEJBTiBDSOG6pFAgSMOATkgNCkLDjSBUSMav"), _

        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

End Sub



Public Sub Ribbon_InsertSigUBNDXa(control As IRibbonControl)

    VFormat_Core.InsertSignatureBlockAtCursor _

        DecodeUTF8("LSBOaMawIHRyw6puOw0KLSBUaMaw4budbmcgdHLhu7FjIMSQ4bqjbmcgcGjGsOG7nW5nLCBIxJBORDsNCi0gTMawdTogVlQu"), _

        DecodeUTF8("VE0uIOG7plkgQkFOIE5Iw4JOIETDgk4NCkNI4bumIFThu4pDSA=="), _

        DecodeUTF8("VHLhuqduIFbEg24gQg==")

End Sub



Public Sub Ribbon_InsertTplChiBoNQ(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Chi_bo_Nghi_quyet.docx"

End Sub



Public Sub Ribbon_InsertTplChiBoBC(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Chi_bo_Bao_cao.docx"

End Sub



Public Sub Ribbon_InsertTplChiBoKH(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Chi_bo_Ke_hoach.docx"

End Sub



Public Sub Ribbon_InsertTplChiBoBB(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Chi_bo_Bien_ban_sinh_hoat.docx"

End Sub



Public Sub Ribbon_InsertTplChiBoCV(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Chi_bo_Cong_van.docx"

End Sub



Public Sub Ribbon_InsertTplXaCV(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Hanh_chinh_Xa_Phuong_Cong_van.docx"

End Sub



Public Sub Ribbon_InsertTplXaQD(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Hanh_chinh_Xa_Phuong_Quyet_dinh.docx"

End Sub



Public Sub Ribbon_InsertTplXaTT(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Hanh_chinh_Xa_Phuong_To_trinh.docx"

End Sub



Public Sub Ribbon_InsertTplXaBC(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Hanh_chinh_Xa_Phuong_Bao_cao.docx"

End Sub



Public Sub Ribbon_InsertTplXaKH(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Hanh_chinh_Xa_Phuong_Ke_hoach.docx"

End Sub



Public Sub Ribbon_InsertTplXaTB(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Hanh_chinh_Xa_Phuong_Thong_bao.docx"

End Sub



Public Sub Ribbon_InsertTplXaBB(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Hanh_chinh_Xa_Phuong_Bien_ban.docx"

End Sub



Public Sub Ribbon_InsertTplXaGM(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Hanh_chinh_Xa_Phuong_Giay_moi.docx"

End Sub



Public Sub Ribbon_InsertTplXaGGT(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Hanh_chinh_Xa_Phuong_Giay_gioi_thieu.docx"

End Sub



Public Sub Ribbon_InsertTplNghiQuyetCaBiet(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Nghi_quyet_ca_biet.docx"

End Sub



Public Sub Ribbon_InsertTplChiThi(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Chi_thi.docx"

End Sub



Public Sub Ribbon_InsertTplHuongDan(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Huong_dan.docx"

End Sub



Public Sub Ribbon_InsertTplQuyDinh(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Quy_dinh.docx"

End Sub



Public Sub Ribbon_InsertTplDeAn(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_De_an.docx"

End Sub



Public Sub Ribbon_InsertTplGiayUyQuyen(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Giay_uy_quyen.docx"

End Sub



Public Sub Ribbon_InsertTplPhieuChuyen(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Phieu_chuyen.docx"

End Sub



Public Sub Ribbon_InsertTplDangCT(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Chi_thi.docx"

End Sub



Public Sub Ribbon_InsertTplDangKL(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Ket_luan.docx"

End Sub



Public Sub Ribbon_InsertTplDangQDi(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Quy_dinh.docx"

End Sub



Public Sub Ribbon_InsertTplDangCTr(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Chuong_trinh.docx"

End Sub



Public Sub Ribbon_InsertTplDangTT(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Thong_tri.docx"

End Sub



Public Sub Ribbon_InsertTplDangBB(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Bien_ban.docx"

End Sub



Public Sub Ribbon_InsertTplDangGM(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Giay_moi.docx"

End Sub



Public Sub Ribbon_InsertTplDangGGT(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Giay_gioi_thieu.docx"

End Sub



Public Sub Ribbon_InsertTplDangPhuLuc(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Phu_luc.docx"

End Sub



Public Sub Ribbon_InsertTplDangBanSao(control As IRibbonControl)

    VFormat_Core.OpenTemplateDoc "Mau_Dang_Ban_sao.docx"

End Sub







