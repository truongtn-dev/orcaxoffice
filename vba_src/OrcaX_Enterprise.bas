Attribute VB_Name = "OrcaX_Enterprise"
Option Explicit

' ==============================================================================
' MODULE: OrcaX_Enterprise
' CHUC NANG:
'   - P0: Trung tam kiem tra van ban (Audit Center / CollectAuditFindings)
'   - P0: Quan ly thong tin van ban (frmDocMetadata)
'   - P0: Profile theo to chuc (redirect -> frmPersonalize)
'   - P1: So sanh van ban (Legal Blackline / Document Diff)
'   - P1: Chuan hoa hang loat (Batch Folder Processing -> ExecuteAutoStandardize)
'   - P1: Lich su chuan hoa & Khoi phuc (Audit History)
'   - P2: Cau truc tai lieu nang cao (Muc luc, Heading 1/2/3, Da cap, Danh muc bang)
' QUY UOC: Tieng Viet trong thong bao gui nguoi dung qua ShowUniMsg(b64).
' ==============================================================================

' ------- Audit severity constants (dung chung voi frmAuditHub) -----------------
Public Const FX_PASS As Long = 1
Public Const FX_PROPOSE As Long = 2
Public Const FX_WARN As Long = 3
Public Const FX_CRIT As Long = 4

' ------- Trang thai audit moi nhat (frmAuditHub doc truc tiep) ----------------
Public g_Findings As Collection        ' moi item: "<sev>|<message>" (message ASCII)
Public g_CritCount As Long
Public g_WarnCount As Long
Public g_PropCount As Long
Public g_PassCount As Long
Public g_Score As Long                 ' 0-100
Public g_AuditDocName As String
Public g_Audited As Boolean

' Thong bao tieng Viet (base64 UTF-8, hien thi qua ShowUniMsg)
Private Const B64_NO_DOC As String = "VnVpIGzDsm5nIG3hu58gbeG7mXQgdsSDbiBi4bqjbiBXb3JkIMSR4buDIGtp4buDbSB0cmEgdGjhu4MgdGjhu6ljIQ=="
Private Const B64_TITLE_AUDIT As String = "T3JjYVggT2ZmaWNlIC0gUsOgIHNvw6F0IHRo4buDIHRo4bupYw=="
Private Const B64_META_DONE As String = "xJDDoyDEkWnhu4FuIHRow7RuZyB0aW4gdsSDbiBi4bqjbiB2w6BvIGtodW5nIHRpw6p1IMSR4buBLg=="
Private Const B64_META_NO_HEADER As String = "Q2jGsGEgdMOsbSB0aOG6pXkga2h1bmcgdGnDqnUgxJHhu4EgxJHhu4MgxJFp4buBbi4gSMOjeSBjaOG6oXkgQ2h14bqpbiBow7NhIFRvw6BuIGRp4buHbiB0csaw4bubYy4="
Private Const B64_PICK_FOLDER As String = "Q2jhu41uIHRoxrAgbeG7pWMgY2jhu6lhIGZpbGUgV29yZCBj4bqnbiBjaHXhuqluIGjDs2E="
Private Const B64_BATCH_DONE As String = "Q2h14bqpbiBow7NhIGjDoG5nIGxv4bqhdCBob8OgbiB04bqldCE="
Private Const B64_BATCH_OK As String = "VOG7lW5nIHPhu5EgZmlsZSDEkcOjIGNodeG6qW4gaMOzYTog"
Private Const B64_BATCH_DIR As String = "VGjGsCBt4bulYyBr4bq/dCBxdeG6ozog"
Private Const B64_BATCH_ERR As String = "U+G7kSBmaWxlIGzhu5dpIGLhu48gcXVhOiA="
Private Const B64_CMP_ORIG As String = "T3JjYVggLSBDaOG7jW4gVsSDbiBi4bqjbiBH4buQQw=="
Private Const B64_CMP_REV As String = "T3JjYVggLSBDaOG7jW4gVsSDbiBi4bqjbiBT4busQSDEkOG7lEk="
Private Const B64_PROCESSING As String = "xJBhbmcgY2h14bqpbiBow7NhOiA="
Private Const B64_SEV_CRIT As String = "TmdoacOqbSB0cuG7jW5n"
Private Const B64_SEV_WARN As String = "Q+G6o25oIGLDoW8="
Private Const B64_SEV_PROP As String = "xJDhu4EgeHXhuqV0"
Private Const B64_SEV_PASS As String = "xJDhuqF0"
Private Const B64_SCORE As String = "xJBp4buDbSB0dcOibiB0aOG7pw=="

' Chuoi Unicode dung de nhan dien noi dung khung tieu de (ma hoa base64)
Private Const B64_QH As String = "Q+G7mE5HIEjDkkEgWMODIEjhu5hJIENI4bumIE5HSMSoQSBWSeG7hlQgTkFN"
Private Const B64_DANG As String = "xJDhuqJORyBD4buYTkcgU+G6ok4gVknhu4ZUIE5BTQ=="
Private Const B64_SO As String = "U+G7kDo="
Private Const B64_NGAY As String = "TkfDgFk="
Private Const B64_THANG As String = "VEjDgU5H"
Private Const B64_NAM As String = "TsSCTQ=="
Private Const B64_DOC_LAP As String = "xJDhu5hDIEzhuqxQ"

' ------------------------------------------------------------------------------
' 1. P0: TRUNG TAM KIEM TRA THE THUC VAN BAN (AUDIT CENTER)
' ------------------------------------------------------------------------------
Public Sub Ribbon_OpenAuditHub(Optional control As IRibbonControl = Nothing)
    On Error Resume Next
    frmAuditHub.Show vbModeless
    If Err.Number <> 0 Then
        Err.Clear
        RunQuickAuditReport
    End If
End Sub

Public Sub Ribbon_AuditDocument(control As IRibbonControl)
    Ribbon_OpenAuditHub control
End Sub

Public Sub Ribbon_InspectDoc(control As IRibbonControl)
    Ribbon_OpenAuditHub control
End Sub

' Chay toan bo bo kiem tra, dien vao trang thai g_* cho frmAuditHub & bao cao
Public Function CollectAuditFindings() As Boolean
    On Error GoTo ErrHandler
    If Documents.Count = 0 Then
        g_Audited = False
        CollectAuditFindings = False
        Exit Function
    End If

    Dim doc As Document
    Set doc = ActiveDocument
    Set g_Findings = New Collection
    g_CritCount = 0
    g_WarnCount = 0
    g_PropCount = 0
    g_PassCount = 0
    g_AuditDocName = doc.Name

    ' 1. Kiem tra kho giay
    Dim sec As Section
    Set sec = doc.Sections(1)
    Dim pgW As Single, pgH As Single
    pgW = sec.PageSetup.PageWidth
    pgH = sec.PageSetup.PageHeight
    If Abs(pgW - 595.3) > 15 Or Abs(pgH - 841.9) > 15 Then
        AddFinding FX_CRIT, "Kho giay chua phai kho A4 chuan (210 x 297 mm)"
    Else
        AddFinding FX_PASS, "Kho giay A4 chuan (210 x 297 mm)"
    End If

    ' 2. Kiem tra can le
    Dim topM As Single, botM As Single, lftM As Single, rgtM As Single
    topM = PointsToMillimeters(sec.PageSetup.TopMargin)
    botM = PointsToMillimeters(sec.PageSetup.BottomMargin)
    lftM = PointsToMillimeters(sec.PageSetup.LeftMargin)
    rgtM = PointsToMillimeters(sec.PageSetup.RightMargin)

    If lftM < 29 Or lftM > 36 Then
        AddFinding FX_CRIT, "Le trai (" & Format(lftM, "0.0") & " mm) sai quy chuan ND 30 (30 - 35 mm)"
    Else
        AddFinding FX_PASS, "Le trai: " & Format(lftM, "0.0") & " mm (chuan 30-35 mm)"
    End If

    If topM < 19 Or topM > 26 Then
        AddFinding FX_WARN, "Le tren (" & Format(topM, "0.0") & " mm) lech quy chuan ND 30 (20 - 25 mm)"
    Else
        AddFinding FX_PASS, "Le tren: " & Format(topM, "0.0") & " mm (chuan 20-25 mm)"
    End If

    If botM < 19 Or botM > 26 Then
        AddFinding FX_WARN, "Le duoi (" & Format(botM, "0.0") & " mm) lech quy chuan ND 30 (20 - 25 mm)"
    Else
        AddFinding FX_PASS, "Le duoi: " & Format(botM, "0.0") & " mm (chuan 20-25 mm)"
    End If

    If rgtM < 14 Or rgtM > 21 Then
        AddFinding FX_WARN, "Le phai (" & Format(rgtM, "0.0") & " mm) lech quy chuan ND 30 (15 - 20 mm)"
    Else
        AddFinding FX_PASS, "Le phai: " & Format(rgtM, "0.0") & " mm (chuan 15-20 mm)"
    End If

    ' 3. Kiem tra Font chu
    Dim p As Paragraph
    Dim nonTNRCount As Long
    Dim pCount As Long
    pCount = doc.Paragraphs.Count
    Dim sampleLimit As Long
    sampleLimit = pCount
    If sampleLimit > 200 Then sampleLimit = 200
    Dim i As Long
    For i = 1 To sampleLimit
        Set p = doc.Paragraphs(i)
        If Len(Trim(p.Range.Text)) > 2 Then
            If InStr(1, p.Range.Font.Name & "", "Times New Roman", vbTextCompare) = 0 Then
                nonTNRCount = nonTNRCount + 1
            End If
        End If
    Next i

    If nonTNRCount > 0 Then
        AddFinding FX_CRIT, "Phat hien " & nonTNRCount & " doan van chua dung phong chu Times New Roman"
    Else
        AddFinding FX_PASS, "Toan bo van ban su dung phong Times New Roman"
    End If

    ' 4. Kiem tra Quoc hieu & Tieu ngu (ca ASCII lan Unicode co dau)
    Dim fullTxt As String
    fullTxt = UCase$(doc.Content.Text)
    Dim quocHieuU As String
    quocHieuU = UCase$(VFormat_Core.DecodeUTF8(B64_QH))
    If InStr(fullTxt, "CONG HOA XA HOI CHU NGHIA VIET NAM") = 0 _
       And InStr(fullTxt, quocHieuU) = 0 Then
        AddFinding FX_PROPOSE, "Chua phat hien Quoc hieu dau trang (Khoi quoc hieu / Tieu ngu)"
    Else
        AddFinding FX_PASS, "Da co khoi Quoc hieu / Tieu ngu dau trang"
    End If

    ' 5. Kiem tra so trang (goi y danh so trang)
    If doc.ComputeStatistics(wdStatisticPages) > 1 Then
        AddFinding FX_PROPOSE, "Van ban nhieu hon 1 trang - nen danh so trang theo quy chuan"
    End If

    ' Tinh diem tuan thu
    Dim total As Long
    total = g_CritCount + g_WarnCount + g_PropCount + g_PassCount
    If total = 0 Then total = 1
    g_Score = CLng(100# * (g_PassCount + 0.5 * g_PropCount) / total)
    If g_Score > 100 Then g_Score = 100
    If g_CritCount = 0 And g_WarnCount = 0 Then
        If g_Score < 90 Then g_Score = 90
    End If
    g_Audited = True
    CollectAuditFindings = True
    Exit Function
ErrHandler:
    g_Audited = False
    CollectAuditFindings = False
End Function

Private Sub AddFinding(ByVal sev As Long, ByVal msg As String)
    g_Findings.Add CStr(sev) & "|" & msg
    Select Case sev
        Case FX_CRIT: g_CritCount = g_CritCount + 1
        Case FX_WARN: g_WarnCount = g_WarnCount + 1
        Case FX_PROPOSE: g_PropCount = g_PropCount + 1
        Case Else: g_PassCount = g_PassCount + 1
    End Select
End Sub

' Tra label tieng Viet cua muc do theo b64 (cho form hien thi)
Public Function FindingSeverityLabelB64(ByVal sev As Long) As String
    Select Case sev
        Case FX_CRIT: FindingSeverityLabelB64 = B64_SEV_CRIT
        Case FX_WARN: FindingSeverityLabelB64 = B64_SEV_WARN
        Case FX_PROPOSE: FindingSeverityLabelB64 = B64_SEV_PROP
        Case Else: FindingSeverityLabelB64 = B64_SEV_PASS
    End Select
End Function

Public Function FindingIconChar(ByVal sev As Long) As String
    Select Case sev
        Case FX_CRIT: FindingIconChar = "[!!] "
        Case FX_WARN: FindingIconChar = "[!]  "
        Case FX_PROPOSE: FindingIconChar = "[i]  "
        Case Else: FindingIconChar = "[OK] "
    End Select
End Function

' Bao cao nhanh dang MsgBox (giu hanh vi cu cho nut "Ra soat nhanh")
Public Sub RunQuickAuditReport()
    If Not CollectAuditFindings() Then
        VFormat_Core.ShowUniMsg B64_NO_DOC, B64_TITLE_AUDIT, 48
        Exit Sub
    End If

    Dim doc As Document
    Set doc = ActiveDocument
    Dim rpt As String
    rpt = "====================================================================" & vbCrLf
    rpt = rpt & "         KET QUA THAM DINH THE THUC VAN BAN - ORCAX AUDIT" & vbCrLf
    rpt = rpt & "====================================================================" & vbCrLf
    rpt = rpt & "Ten van ban: " & doc.Name & vbCrLf
    rpt = rpt & "Thoi gian kiem tra: " & Now & vbCrLf & vbCrLf

    Dim f As Variant
    For Each f In g_Findings
        rpt = rpt & FindingIconChar(CLng(Split(CStr(f), "|")(0))) & Split(CStr(f), "|")(1) & vbCrLf
    Next f

    rpt = rpt & vbCrLf & "--------------------------------------------------------------------" & vbCrLf
    rpt = rpt & "TONG HOP LOI: " & g_CritCount & " Loi nghiem trong | " & g_WarnCount & _
          " Canh bao | " & g_PropCount & " De xuat" & vbCrLf
    rpt = rpt & "DIEM TUAN THU: " & g_Score & "/100" & vbCrLf
    rpt = rpt & "Ban co the bam 'Chuan hoa Toan dien' de tu dong sua cac loi tren!" & vbCrLf
    rpt = rpt & "===================================================================="
    MsgBox rpt, vbInformation, "OrcaX Office - Trung tam Tham dinh The thuc"
End Sub

' ------------------------------------------------------------------------------
' 2. P0: QUAN LY THONG TIN VAN BAN (METADATA HUB -> frmDocMetadata)
' ------------------------------------------------------------------------------
Public Sub Ribbon_OpenDocMetadata(Optional control As IRibbonControl = Nothing)
    On Error Resume Next
    frmDocMetadata.Show vbModeless
    If Err.Number <> 0 Then
        Err.Clear
        VFormat_Core.ShowUniMsg B64_META_NO_HEADER, B64_TITLE_AUDIT, 48
    End If
End Sub

' Duoc frmDocMetadata goi: dien thong tin vao khung tieu de hien co
' Tra ve True neu xu ly duoc it nhat 1 thanh phan
Public Function ApplyDocMetadata(ByVal docNumber As String, ByVal docDate As String, _
                                 ByVal location As String, ByVal agencyName As String) As Boolean
    On Error GoTo Fail
    If Documents.Count = 0 Then Exit Function
    Dim doc As Document
    Set doc = ActiveDocument

    ' Chuoi nhan dien Unicode (co dau) va ASCII
    Dim soU As String, ngayU As String, namU As String, thangU As String
    soU = UCase$(VFormat_Core.DecodeUTF8(B64_SO))
    ngayU = UCase$(VFormat_Core.DecodeUTF8(B64_NGAY))
    namU = UCase$(VFormat_Core.DecodeUTF8(B64_NAM))
    thangU = UCase$(VFormat_Core.DecodeUTF8(B64_THANG))
    Dim qhU As String, dangU As String, docLapU As String
    qhU = UCase$(VFormat_Core.DecodeUTF8(B64_QH))
    dangU = UCase$(VFormat_Core.DecodeUTF8(B64_DANG))
    docLapU = UCase$(VFormat_Core.DecodeUTF8(B64_DOC_LAP))

    Dim applied As Boolean
    applied = False

    ' Tim khung tieu de: 3 bang dau tien, chua dau hieu so / ngay / quoc hieu
    Dim tbl As Table, tblFound As Table
    Dim maxCheck As Long
    If doc.Tables.Count < 3 Then maxCheck = doc.Tables.Count Else maxCheck = 3
    Dim i As Long
    For i = 1 To maxCheck
        Set tbl = doc.Tables(i)
        Dim cellTxt As String
        cellTxt = UCase$(tbl.Range.Text)
        If InStr(cellTxt, "SO:") > 0 Or InStr(cellTxt, soU) > 0 _
           Or InStr(cellTxt, "CONG HOA") > 0 Or InStr(cellTxt, qhU) > 0 _
           Or InStr(cellTxt, "NGAY") > 0 Or InStr(cellTxt, ngayU) > 0 Then
            Set tblFound = tbl
            Exit For
        End If
    Next i

    If tblFound Is Nothing Then
        ApplyDocMetadata = False
        Exit Function
    End If

    ' Duyet tung cell va dien theo mau noi dung
    Dim c As Cell, txt As String, uTxt As String, newTxt As String
    For Each c In tblFound.Range.Cells
        txt = GetCellText(c)
        uTxt = UCase$(txt)
        newTxt = txt

        ' 1) So & ky hieu: "So: ..." -> giu prefix, thay so moi
        If Len(docNumber) > 0 And (InStr(uTxt, "SO:") > 0 Or InStr(uTxt, soU) > 0) Then
            Dim colonPos As Long
            colonPos = InStr(uTxt, "SO:")
            If colonPos = 0 Then colonPos = InStr(uTxt, soU)
            If colonPos > 0 Then
                newTxt = RTrim$(Left$(txt, colonPos + Len(soU) - 1)) & " " & docNumber
                SetCellText c, newTxt
                applied = True
            End If

        ' 2) Dong dia danh & ngay thang: "..., ngay ... thang ... nam ..."
        ElseIf Len(docDate) > 0 And (InStr(uTxt, "NGAY") > 0 Or InStr(uTxt, ngayU) > 0) _
               And (InStr(uTxt, "NAM") > 0 Or InStr(uTxt, namU) > 0) Then
            Dim locPart As String
            If Len(location) > 0 Then
                locPart = location
            Else
                locPart = ExtractLocationFrom(txt, ngayU)
            End If
            newTxt = locPart & ", " & FormatVietnameseDate(docDate, ngayU, thangU, namU)
            SetCellText c, newTxt
            applied = True

        ' 3) Ten co quan: dong IN HOA dai, khong phai quoc hieu / so / ngay
        ElseIf Len(agencyName) > 0 And Len(Trim$(txt)) >= 8 _
               And uTxt = Trim$(txt) _
               And InStr(uTxt, "SO:") = 0 And InStr(uTxt, soU) = 0 _
               And InStr(uTxt, "NGAY") = 0 And InStr(uTxt, ngayU) = 0 _
               And InStr(uTxt, qhU) = 0 And InStr(uTxt, dangU) = 0 _
               And InStr(uTxt, docLapU) = 0 Then
            If IsAgencyLine(uTxt) Then
                SetCellText c, UCase$(agencyName)
                applied = True
            End If
        End If
    Next c

    ApplyDocMetadata = applied
    Exit Function
Fail:
    ApplyDocMetadata = False
End Function

Private Function GetCellText(ByVal c As Cell) As String
    On Error Resume Next
    Dim t As String
    t = c.Range.Text
    ' bo ky hieu ket thuc cell (Chr13 & Chr7)
    t = Replace(t, Chr(13) & Chr(7), "")
    t = Replace(t, Chr(13), "")
    GetCellText = Trim$(t)
End Function

Private Sub SetCellText(ByVal c As Cell, ByVal newTxt As String)
    On Error Resume Next
    c.Range.Text = newTxt
End Sub

' Nhan dien dong ten co quan theo cac tien to hanh chinh (da UCase, ho tro Unicode)
Private Function IsAgencyLine(ByVal uTxt As String) As Boolean
    On Error Resume Next
    Static loaded As Boolean
    Static prefixes As Collection
    If Not loaded Then
        Set prefixes = New Collection
        prefixes.Add UCase$(VFormat_Core.DecodeUTF8("QuG7mCA="))    ' BO
        prefixes.Add UCase$(VFormat_Core.DecodeUTF8("U+G7niA="))    ' SO
        prefixes.Add UCase$(VFormat_Core.DecodeUTF8("4bumWSBCQU4=")) ' UY BAN
        prefixes.Add "UBND"
        prefixes.Add UCase$(VFormat_Core.DecodeUTF8("VsSCTiBQSMOSTkc=")) ' VAN PHONG
        prefixes.Add UCase$(VFormat_Core.DecodeUTF8("Q+G7pEMg"))    ' CUC
        prefixes.Add UCase$(VFormat_Core.DecodeUTF8("UEjDkk5HIA==")) ' PHONG
        prefixes.Add UCase$(VFormat_Core.DecodeUTF8("VFLGr+G7nE5HIA==")) ' TRUONG
        prefixes.Add UCase$(VFormat_Core.DecodeUTF8("SOG7mEkg"))    ' HOI
        prefixes.Add UCase$(VFormat_Core.DecodeUTF8("QkFOIA=="))    ' BAN
        loaded = True
    End If
    Dim p As Variant
    For Each p In prefixes
        If InStr(uTxt, CStr(p)) = 1 Then
            IsAgencyLine = True
            Exit Function
        End If
    Next p
    IsAgencyLine = False
End Function

Private Function ExtractLocationFrom(ByVal t As String, ByVal ngayU As String) As String
    ' Trich dia danh truoc ", ngay" (ca ASCII lan Unicode co dau)
    Dim p As Long
    p = InStr(UCase$(t), ", NGAY")
    If p < 1 Then p = InStr(UCase$(t), ", " & UCase$(ngayU))
    If p > 1 Then
        ExtractLocationFrom = Trim$(Left$(t, p - 1))
    Else
        ExtractLocationFrom = ""
    End If
End Function

' "16/11/2026" -> "ngay 16 thang 11 nam 2026" (Unicode neu co, ASCII neu khong)
Public Function FormatVietnameseDate(ByVal d As String, ByVal ngayU As String, _
                                     ByVal thangU As String, ByVal namU As String) As String
    On Error Resume Next
    Dim parts() As String
    parts = Split(Trim$(d), "/")
    If UBound(parts) = 2 Then
        If Len(ngayU) > 5 Then
            FormatVietnameseDate = ngayU & " " & parts(0) & " " & thangU & " " & parts(1) & " " & namU & " " & parts(2)
        Else
            FormatVietnameseDate = "ngay " & parts(0) & " thang " & parts(1) & " nam " & parts(2)
        End If
    Else
        FormatVietnameseDate = Trim$(d)
    End If
End Function

' ------------------------------------------------------------------------------
' 3. P0: PROFILE THEO TO CHUC (redirect -> Cai dat Ca nhan hoa)
' ------------------------------------------------------------------------------
Public Sub Ribbon_OpenOrgProfiles(control As IRibbonControl)
    On Error Resume Next
    frmPersonalize.Show vbModeless
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Profile he thong: Nghi dinh 30/2020/ND-CP, Huong dan 05-HD/VPTW, Nghi dinh 250-CP.", _
               vbInformation, "OrcaX Office - Profile Engine"
    End If
End Sub

' ------------------------------------------------------------------------------
' 4. P1: SO SANH VAN BAN (LEGAL BLACKLINE / DOCUMENT DIFF)
' ------------------------------------------------------------------------------
Public Sub Ribbon_CompareDocuments(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim fd As FileDialog
    Dim origPath As String, revPath As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = VFormat_Core.DecodeUTF8(B64_CMP_ORIG)
    fd.Filters.Clear
    fd.Filters.Add "Word Documents", "*.docx;*.doc;*.dotx;*.dotm"
    fd.AllowMultiSelect = False

    If fd.Show <> -1 Then Exit Sub
    origPath = fd.SelectedItems(1)

    fd.Title = VFormat_Core.DecodeUTF8(B64_CMP_REV)
    If fd.Show <> -1 Then Exit Sub
    revPath = fd.SelectedItems(1)

    Dim docOrig As Document
    Set docOrig = Documents.Open(FileName:=origPath, ReadOnly:=True, Visible:=False)
    Dim docRev As Document
    Set docRev = Documents.Open(FileName:=revPath, ReadOnly:=True, Visible:=False)

    Dim docCompared As Document
    Set docCompared = Application.CompareDocuments( _
        OriginalDocument:=docOrig, _
        RevisedDocument:=docRev, _
        Destination:=wdCompareDestinationNew, _
        Granularity:=wdGranularityWordLevel, _
        CompareFormatting:=True, _
        ShowChangesIn:=wdShowSourceDocumentsBoth)

    docOrig.Close SaveChanges:=False
    docRev.Close SaveChanges:=False

    docCompared.Activate
    docCompared.ShowRevisions = True

    Dim revCount As Long
    revCount = docCompared.Revisions.Count
    MsgBox "So sanh van ban hoan tat!" & vbCrLf & _
           "- Van ban goc: " & Dir(origPath) & vbCrLf & _
           "- Van ban sua doi: " & Dir(revPath) & vbCrLf & _
           "- Phat hien: " & revCount & " vi tri thay doi (Them, Xoa, Sua, Dinh dang)." & vbCrLf & _
           "Van ban doi chieu Tracked Changes da duoc tao thanh cong!", vbInformation, _
           "OrcaX Office - So Sanh Van Ban"
    Exit Sub
ErrHandler:
    MsgBox "Loi trong qua trinh so sanh van ban: " & Err.Description, vbCritical, "OrcaX Office"
End Sub

' ------------------------------------------------------------------------------
' 5. P1: CHUAN HOA HANG LOAT (BATCH FOLDER PROCESSING -> frmBatchFolder)
' ------------------------------------------------------------------------------
Public Sub Ribbon_OpenBatchFolder(Optional control As IRibbonControl = Nothing)
    On Error Resume Next
    frmBatchFolder.Show vbModeless
    If Err.Number <> 0 Then
        Err.Clear
        BatchProcessFolderInteractive
    End If
End Sub

' Xu ly 1 file don le cho batch: tra ve True neu thanh cong
' Luu y: regIdx theo he so 0-based cua Core (0 = ND 30, 1 = HD 05, 2 = ND 250)
Public Function BatchProcessOneFile(ByVal fullPath As String, ByVal outFolder As String, _
                                    ByVal regIdx As Long) As Boolean
    On Error GoTo Fail
    Dim doc As Document
    Set doc = Documents.Open(FileName:=fullPath, Visible:=False)
    If doc Is Nothing Then Exit Function

    VFormat_Core.SetSelectedRegulation regIdx
    VFormat_Core.ExecuteAutoStandardize doc, True
    doc.SaveAs2 FileName:=outFolder & "\" & Dir$(fullPath)
    doc.Close SaveChanges:=False
    Set doc = Nothing
    BatchProcessOneFile = True
    Exit Function
Fail:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    BatchProcessOneFile = False
End Function

' Xay danh sach file *.doc* trong thu muc (co the de quy)
Public Function CollectDocFiles(ByVal folderPath As String, ByVal recursive As Boolean) As Collection
    Dim result As Collection
    Set result = New Collection
    CollectDocFilesInto folderPath, recursive, result
    Set CollectDocFiles = result
End Function

Private Sub CollectDocFilesInto(ByVal folderPath As String, ByVal recursive As Boolean, ByVal result As Collection)
    On Error Resume Next
    Dim fName As String
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    fName = Dir$(folderPath & "*.doc*")
    Do While Len(fName) > 0
        If Left$(fName, 2) <> "~$" Then result.Add folderPath & fName
        fName = Dir$()
    Loop
    If recursive Then
        ' Thu lai toan bo thu muc con TRUOC, tranh pha trang thai Dir cua vong file
        Dim subDirs As Collection
        Set subDirs = New Collection
        Dim subDir As String
        subDir = Dir$(folderPath & "*", vbDirectory)
        Do While Len(subDir) > 0
            If subDir <> "." And subDir <> ".." Then
                If (GetAttr(folderPath & subDir) And vbDirectory) = vbDirectory Then
                    subDirs.Add folderPath & subDir
                End If
            End If
            subDir = Dir$()
        Loop
        Dim sd As Variant
        For Each sd In subDirs
            CollectDocFilesInto CStr(sd) & "\", True, result
        Next sd
    End If
End Sub

' Luong cu (fallback neu form khong co): chon thu muc -> chuan hoa toan bo
Public Sub BatchProcessFolderInteractive()
    On Error GoTo ErrHandler
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = VFormat_Core.DecodeUTF8(B64_PICK_FOLDER)
    If fd.Show <> -1 Then Exit Sub

    Dim folderPath As String
    folderPath = fd.SelectedItems(1)
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim outFolder As String
    outFolder = folderPath & "ChuanHoa_OrcaX"
    If Dir(outFolder, vbDirectory) = "" Then MkDir outFolder

    Dim files As Collection
    Set files = CollectDocFiles(folderPath, False)

    Dim procCount As Long, errCount As Long, i As Long
    Application.ScreenUpdating = False
    VFormat_Core.StartProgress "OrcaX Batch", "Chuan hoa hang loat...", "BatchProcessFolderInteractive"
    For i = 1 To files.Count
        VFormat_Core.SetProgress CLng(100 * (i - 1) / files.Count), _
            "File " & i & "/" & files.Count, "BatchProcessFolderInteractive"
        If BatchProcessOneFile(CStr(files(i)), outFolder, 0) Then
            procCount = procCount + 1
        Else
            errCount = errCount + 1
        End If
    Next i
    VFormat_Core.FinishProgress "BatchProcessFolderInteractive"
    Application.ScreenUpdating = True

    VFormat_Core.ShowUniMsg _
        VFormat_Core.DecodeUTF8(B64_BATCH_DONE) & vbCrLf & _
        VFormat_Core.DecodeUTF8(B64_BATCH_OK) & procCount & vbCrLf & _
        VFormat_Core.DecodeUTF8(B64_BATCH_DIR) & outFolder & vbCrLf & _
        VFormat_Core.DecodeUTF8(B64_BATCH_ERR) & errCount, _
        "OrcaX Office - Chuan Hoa Hang Loat", 64
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Loi chuan hoa hang loat: " & Err.Description, vbCritical, "OrcaX Office"
End Sub

' ------------------------------------------------------------------------------
' 6. P1: LICH SU CHUAN HOA & KHOI PHUC (AUDIT HISTORY)
' ------------------------------------------------------------------------------
Public Sub Ribbon_ShowAuditHistory(control As IRibbonControl)
    On Error Resume Next
    Dim msg As String
    msg = "LICH SU CHUAN HOA & AUDIT TRAIL - ORCAX OFFICE" & vbCrLf & _
          "------------------------------------------------" & vbCrLf & _
          "- Nguoi thuc hien: " & Application.UserName & vbCrLf & _
          "- Phien ban OrcaX: Workbench v2.0" & vbCrLf & _
          "- Quy chuan ap dung gan nhat: " & RegulationName() & vbCrLf & _
          "- Diem tuan thu lan kiem tra gan nhat: " & IIf(g_Audited, g_Score & "/100", "(chua kiem tra)") & vbCrLf & _
          "- Trang thai ban luu phuc hoi: " & IIf(g_Audited, "San sang", "Chua co") & vbCrLf & vbCrLf & _
          "Bam nut 'Khoi phuc nguyen trang' tren Ribbon de quay lai trang thai truoc khi chuan hoa."
    MsgBox msg, vbInformation, "OrcaX Office - Lich Su Chuan Hoa"
End Sub

Private Function RegulationName() As String
    On Error Resume Next
    ' Chi so quy chuan 0-based: 0 = ND 30, 1 = HD 05, 2 = ND 250, 3 = KHBD
    Select Case VFormat_Core.GetEffectiveRegulation()
        Case 0: RegulationName = "Han chinh (ND 30/2020)"
        Case 1: RegulationName = "Van ban Dang (HD 05)"
        Case 2: RegulationName = "Hop nhat (ND 250)"
        Case 3: RegulationName = "Ke hoach bai day (CV 5512)"
        Case Else: RegulationName = "Tu dong nhan dien"
    End Select
End Function

Public Sub Ribbon_ExportAuditReport(control As IRibbonControl)
    RunQuickAuditReport
End Sub

' ------------------------------------------------------------------------------
' 7. P2: CONG CU CAU TRUC TAI LIEU NANG CAO (TOC, HEADINGS, MULTILEVEL)
' ------------------------------------------------------------------------------
Public Sub Ribbon_InsertAutoTOC(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub

    Dim rng As Range
    If doc.TablesOfContents.Count > 0 Then
        doc.TablesOfContents(1).Update
        MsgBox "Da cap nhat tu dong Muc luc van ban!", vbInformation, "OrcaX Office"
    Else
        Set rng = Selection.Range
        doc.TablesOfContents.Add Range:=rng, _
            RightAlignPageNumbers:=True, _
            UseHeadingStyles:=True, _
            UpperHeadingLevel:=1, _
            LowerHeadingLevel:=3, _
            IncludePageNumbers:=True
        MsgBox "Da tao Muc luc tu dong chuan cap 1-3!", vbInformation, "OrcaX Office"
    End If
    Exit Sub
ErrHandler:
    MsgBox "Loi tao muc luc tu dong: " & Err.Description, vbCritical, "OrcaX Office"
End Sub

Public Sub Ribbon_FormatHeading1(control As IRibbonControl)
    On Error Resume Next
    With Selection.Paragraphs(1).Range
        .Style = ActiveDocument.Styles(wdStyleHeading1)
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .ParagraphFormat.SpaceBefore = 6
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
    End With
End Sub

Public Sub Ribbon_FormatHeading2(control As IRibbonControl)
    On Error Resume Next
    With Selection.Paragraphs(1).Range
        .Style = ActiveDocument.Styles(wdStyleHeading2)
        .Font.Name = "Times New Roman"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Italic = True
        .ParagraphFormat.SpaceBefore = 4
        .ParagraphFormat.SpaceAfter = 4
    End With
End Sub

Public Sub Ribbon_FormatHeading3(control As IRibbonControl)
    On Error Resume Next
    With Selection.Paragraphs(1).Range
        .Style = ActiveDocument.Styles(wdStyleHeading3)
        .Font.Name = "Times New Roman"
        .Font.Size = 13
        .Font.Bold = False
        .Font.Italic = True
        .ParagraphFormat.SpaceBefore = 2
        .ParagraphFormat.SpaceAfter = 2
    End With
End Sub

Public Sub Ribbon_FormatMultilevelList(control As IRibbonControl)
    On Error Resume Next
    Selection.Range.ListFormat.ApplyNumberDefault
    MsgBox "Da ap dung dinh dang danh sach danh so phan cap hanh chinh!", vbInformation, "OrcaX Office"
End Sub

Public Sub Ribbon_InsertTableOfFigures(control As IRibbonControl)
    On Error Resume Next
    ActiveDocument.TablesOfFigures.Add Range:=Selection.Range, Caption:=wdCaptionTable
    MsgBox "Da chen Danh muc Bang bieu / Hinh anh tu dong!", vbInformation, "OrcaX Office"
End Sub

' ------------------------------------------------------------------------------
' 8. TIEN ICH DINH DANG & LAM SACH VAN BAN
' ------------------------------------------------------------------------------
Public Sub Ribbon_SetStandardMarginsA4(control As IRibbonControl)
    On Error Resume Next
    With ActiveDocument.PageSetup
        .PaperSize = wdPaperA4
        .TopMargin = MillimetersToPoints(20)
        .BottomMargin = MillimetersToPoints(20)
        .LeftMargin = MillimetersToPoints(30)
        .RightMargin = MillimetersToPoints(15)
    End With
    MsgBox "Da thiet lap Can Le Chuan A4 Nghi dinh 30:" & vbCrLf & _
           "- Le trai: 30 mm (de dong gay)" & vbCrLf & _
           "- Le phai: 15 mm" & vbCrLf & _
           "- Le tren: 20 mm" & vbCrLf & _
           "- Le duoi: 20 mm", vbInformation, "OrcaX Office"
End Sub

Public Sub Ribbon_StandardizeParagraphAndFont(control As IRibbonControl)
    On Error Resume Next
    With Selection.Range
        .Font.Name = "Times New Roman"
        .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
        .ParagraphFormat.LineSpacing = LinesToPoints(1.2)
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
    End With
    MsgBox "Da dinh dang chuan doan van ban: Times New Roman, gian dong 1.2, cach doan duoi 6pt, can deu 2 ben!", _
           vbInformation, "OrcaX Office"
End Sub

Public Sub Ribbon_FixVietnameseSpellingAndPunctuation(control As IRibbonControl)
    On Error Resume Next
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = " ,"
        .Replacement.Text = ","
        .Execute Replace:=wdReplaceAll
        .Text = " ."
        .Replacement.Text = "."
        .Execute Replace:=wdReplaceAll
        .Text = " :"
        .Replacement.Text = ":"
        .Execute Replace:=wdReplaceAll
        .Text = " ;"
        .Replacement.Text = ";"
        .Execute Replace:=wdReplaceAll
        .Text = " ?"
        .Replacement.Text = "?"
        .Execute Replace:=wdReplaceAll
        .Text = " !"
        .Replacement.Text = "!"
        .Execute Replace:=wdReplaceAll
    End With
    MsgBox "Da lam sach va chuan hoa dau cau tieng Viet theo quy chuan!", vbInformation, "OrcaX Office"
End Sub

Public Sub Ribbon_CleanExtraSpacesAndLines(control As IRibbonControl)
    On Error Resume Next
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "  "
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        .Execute Replace:=wdReplaceAll
        .Text = "^p^p^p"
        .Replacement.Text = "^p^p"
        .Execute Replace:=wdReplaceAll
    End With
    MsgBox "Da xoa toan bo khoang trang loi va dong trong thua trong tai lieu!", vbInformation, "OrcaX Office"
End Sub
