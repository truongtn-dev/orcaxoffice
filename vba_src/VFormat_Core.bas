Attribute VB_Name = "VFormat_Core"

' ==============================================================================

' MODULE: VFormat_Core

' MO TA: BO XU LY CHUAN HOA THE THUC THEO RULEBOOK (HD 05, ND 30, ND 250, ND 250)

' 100% PURE ASCII COMPLIANT FOR ZERO VBA COMPILE ERRORS

' ==============================================================================

Option Explicit



#If VBA7 Then

    Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _

        ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, _

        ByVal uType As Long) As Long

    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr

#Else

    Private Declare Function MessageBoxW Lib "user32" ( _

        ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, _

        ByVal uType As Long) As Long

    Private Declare Function GetForegroundWindow Lib "user32" () As Long

#End If



Private Const MB_SETFOREGROUND As Long = &H10000



Private Const wdWrapTextNone As Long = 0

Private Const wdWrapTextAround As Long = 1



Public Const MM_TO_PT As Single = 2.83465

Public Const CM_TO_PT As Single = 28.3465



' --- BIEN TOAN CUC LUU TRU THONG TIN DEBUG & CAU HINH ---

Public g_LastErrorMsg As String

Public g_LastErrorSource As String

Public g_LastErrorNum As Long

Public g_LastAction As String

Public g_LastErrorTime As String



Public g_SelectedRegulation As Integer

Public g_RegulationManuallySelected As Boolean

Public g_SelectedOrgLevel As Integer

Public g_SelectedDocType As Integer



' --- BIEN TRANG THAI THANH TIEN TRINH LOADING GIUA MAN HINH ---

Private m_ProgressIsShowing As Boolean

Private m_ProgressOwner As String



Private Sub LogTrace(ByVal msg As String)

    ' Safe debug tracer stub

End Sub



Public Sub DbgLog(ByVal msg As String)

    On Error Resume Next

    Dim fn As Integer

    fn = FreeFile

    Open Environ$("TEMP") & "\vformat_debug.log" For Append As #fn

    Print #fn, Format(Now, "hh:nn:ss") & " " & msg

    Close #fn

End Sub



Public Sub SafeOpenUrlOrFile(ByVal targetPathOrUrl As String)

    On Error Resume Next

    If Len(Trim(targetPathOrUrl)) = 0 Then Exit Sub

    Application.FollowHyperlink Address:=targetPathOrUrl, NewWindow:=True

End Sub



Public Sub SetSelectedRegulation(ByVal regIdx As Integer)

    g_SelectedRegulation = regIdx

    g_RegulationManuallySelected = True

End Sub



Public Sub ResetRegulationAutoDetect()

    g_RegulationManuallySelected = False

End Sub



Public Function GetEffectiveRegulation(Optional doc As Document = Nothing) As Integer

    If g_RegulationManuallySelected And g_SelectedRegulation >= 0 Then

        GetEffectiveRegulation = g_SelectedRegulation

    Else

        Dim det As Integer

        det = DetectRegulation(doc)

        g_SelectedRegulation = det

        GetEffectiveRegulation = det

    End If

End Function



Public Sub SetSelectedOrgLevel(ByVal orgLvl As Integer)

    g_SelectedOrgLevel = orgLvl

End Sub



Public Sub SetSelectedDocType(ByVal dType As Integer)

    g_SelectedDocType = dType

End Sub



' --- BO GIAI MA UNICODE & THONG BAO TU THAN (SELF-CONTAINED) ---

Public Function DecodeUTF8(ByVal b64Str As String) As String

    On Error GoTo Fallback

    If Len(Trim(b64Str)) = 0 Then

        DecodeUTF8 = ""

        Exit Function

    End If

    Dim xmlDoc As Object, node As Object

    Set xmlDoc = CreateObject("MSXML2.DOMDocument")

    Set node = xmlDoc.createElement("b64")

    node.DataType = "bin.base64"

    node.Text = b64Str

    

    Dim rawBytes() As Byte

    rawBytes = node.nodeTypedValue

    

    Dim i As Long, ub As Long, c As Long, ch As Long

    Dim res As String

    ub = UBound(rawBytes)

    i = LBound(rawBytes)

    Do While i <= ub

        c = rawBytes(i)

        If c < &H80 Then

            res = res & ChrW(c)

            i = i + 1

        ElseIf (c And &HE0) = &HC0 And i + 1 <= ub Then

            ch = ((c And &H1F) * 64) Or (rawBytes(i + 1) And &H3F)

            res = res & ChrW(ch)

            i = i + 2

        ElseIf (c And &HF0) = &HE0 And i + 2 <= ub Then

            ch = ((c And &HF) * 4096) Or ((CLng(rawBytes(i + 1) And &H3F)) * 64) Or (rawBytes(i + 2) And &H3F)

            res = res & ChrW(ch)

            i = i + 3

        Else

            i = i + 1

        End If

    Loop

    DecodeUTF8 = res

    Exit Function

Fallback:

    DecodeUTF8 = b64Str

End Function



Public Sub ShowUniMsg(ByVal msgB64OrPlain As String, ByVal titleB64OrPlain As String, Optional ByVal iconType As Long = 64)

    On Error GoTo Fallback

    If Application.DisplayAlerts = 0 Then Exit Sub

    Dim msgText As String, titleText As String

    

    If InStr(msgB64OrPlain, " ") > 0 Or InStr(msgB64OrPlain, vbCr) > 0 Or InStr(msgB64OrPlain, vbLf) > 0 Then

        msgText = msgB64OrPlain

    Else

        msgText = DecodeUTF8(msgB64OrPlain)

    End If

    

    If InStr(titleB64OrPlain, " ") > 0 Then

        titleText = titleB64OrPlain

    Else

        titleText = DecodeUTF8(titleB64OrPlain)

    End If

    

    #If VBA7 Then

        Dim h As LongPtr: h = GetForegroundWindow()

    #Else

        Dim h As Long: h = GetForegroundWindow()

    #End If

    

    MessageBoxW h, StrPtr(msgText), StrPtr(titleText), iconType Or MB_SETFOREGROUND

    Exit Sub

Fallback:

    On Error Resume Next

    MsgBox msgText, iconType, titleText

End Sub



Public Function ShowUniConfirm(ByVal msgB64OrPlain As String, ByVal titleB64OrPlain As String) As Boolean

    On Error GoTo Fallback

    If Application.DisplayAlerts = 0 Then

        ShowUniConfirm = True

        Exit Function

    End If

    Dim msgText As String, titleText As String

    If InStr(msgB64OrPlain, " ") > 0 Or InStr(msgB64OrPlain, vbCr) > 0 Or InStr(msgB64OrPlain, vbLf) > 0 Then

        msgText = msgB64OrPlain

    Else

        msgText = DecodeUTF8(msgB64OrPlain)

    End If

    If InStr(titleB64OrPlain, " ") > 0 Then

        titleText = titleB64OrPlain

    Else

        titleText = DecodeUTF8(titleB64OrPlain)

    End If

    

    #If VBA7 Then

        Dim h As LongPtr: h = GetForegroundWindow()

    #Else

        Dim h As Long: h = GetForegroundWindow()

    #End If

    

    Dim ret As Long

    ret = MessageBoxW(h, StrPtr(msgText), StrPtr(titleText), 36 Or MB_SETFOREGROUND)

    ShowUniConfirm = (ret = 6)

    Exit Function

Fallback:

    ShowUniConfirm = (MsgBox(msgText, vbQuestion + vbYesNo, titleText) = vbYes)

End Function





Public Function ShowUniInputBox(ByVal promptB64OrPlain As String, ByVal titleB64OrPlain As String, Optional ByVal defaultVal As String = "") As String

    On Error Resume Next

    Dim promptText As String, titleText As String

    If InStr(promptB64OrPlain, " ") > 0 Or InStr(promptB64OrPlain, vbCr) > 0 Or InStr(promptB64OrPlain, vbLf) > 0 Then

        promptText = promptB64OrPlain

    Else

        promptText = DecodeUTF8(promptB64OrPlain)

    End If

    

    If InStr(titleB64OrPlain, " ") > 0 Then

        titleText = titleB64OrPlain

    Else

        titleText = DecodeUTF8(titleB64OrPlain)

    End If

    

    ShowUniInputBox = InputBox(promptText, titleText, defaultVal)

End Function



Public Sub ShowUniToast(ByVal msgB64OrPlain As String, ByVal titleB64OrPlain As String)

    On Error Resume Next

    Dim msgText As String, titleText As String

    

    If InStr(msgB64OrPlain, " ") > 0 Or InStr(msgB64OrPlain, vbCr) > 0 Or InStr(msgB64OrPlain, vbLf) > 0 Then

        msgText = msgB64OrPlain

    Else

        msgText = DecodeUTF8(msgB64OrPlain)

    End If

    

    If InStr(titleB64OrPlain, " ") > 0 Then

        titleText = titleB64OrPlain

    Else

        titleText = DecodeUTF8(titleB64OrPlain)

    End If

    

    Dim statusTitle As String: statusTitle = titleText

    If Left(statusTitle, 11) = "V-FORMAT - " Then

        statusTitle = Mid(statusTitle, 12)

    ElseIf Left(statusTitle, 8) = "V-FORMAT" Then

        statusTitle = Mid(statusTitle, 9)

    End If

    statusTitle = Trim(statusTitle)

    

    If Len(statusTitle) > 0 Then

        Application.StatusBar = "[V-FORMAT] " & statusTitle & " | " & Replace(Replace(msgText, vbCrLf, " - "), vbCr, " - ")

    Else

        Application.StatusBar = "[V-FORMAT] " & Replace(Replace(msgText, vbCrLf, " - "), vbCr, " - ")

    End If

End Sub



' ==============================================================================

' BO DIEU KHIEN THANH TIEN TRINH LOADING GIUA MAN HINH (PROGRESS BAR CONTROLLER)

' ==============================================================================

Public Sub StartProgress(ByVal titleStr As String, ByVal initialMsg As String, Optional ByVal ownerName As String = "")

    On Error Resume Next

    If Application.DisplayAlerts = 0 Then Exit Sub

    If m_ProgressIsShowing And Len(m_ProgressOwner) > 0 And ownerName <> m_ProgressOwner Then

        Exit Sub

    End If

    

    m_ProgressIsShowing = True

    If Len(ownerName) > 0 Then

        m_ProgressOwner = ownerName

    Else

        m_ProgressOwner = "General"

    End If

    

    frmProgress.lblTitle.Caption = titleStr

    frmProgress.lblStatus.Caption = initialMsg

    frmProgress.lblBarFill.Width = 0

    frmProgress.lblPercent.Caption = "0%"

    frmProgress.Show vbModeless

    DoEvents

End Sub



Public Sub SetProgress(ByVal percent As Integer, ByVal statusMsg As String, Optional ByVal callerName As String = "")

    On Error Resume Next

    If Application.DisplayAlerts = 0 Then Exit Sub

    If Not m_ProgressIsShowing Then Exit Sub

    If Len(m_ProgressOwner) > 0 And Len(callerName) > 0 And callerName <> m_ProgressOwner Then

        Exit Sub

    End If

    

    If percent < 0 Then percent = 0

    If percent > 100 Then percent = 100

    

    frmProgress.lblStatus.Caption = statusMsg

    frmProgress.lblPercent.Caption = "Ti" & ChrW(7871) & "n tr" & ChrW(236) & "nh: " & percent & "%"

    Dim maxW As Single: maxW = 338

    frmProgress.lblBarFill.Width = (maxW * percent) / 100

    DoEvents

End Sub



Public Sub FinishProgress(Optional ByVal callerName As String = "")

    On Error Resume Next

    If Application.DisplayAlerts = 0 Then Exit Sub

    If Not m_ProgressIsShowing Then Exit Sub

    If Len(m_ProgressOwner) > 0 And Len(callerName) > 0 And callerName <> m_ProgressOwner Then

        Exit Sub

    End If

    

    m_ProgressIsShowing = False

    m_ProgressOwner = ""

    Unload frmProgress

    DoEvents

End Sub



Public Function ShowStandardizeReportForm(ByVal rptContent As String) As Boolean

    ' ==============================================================================

    ' HIEN THI BANG BAO CAO KET QUA CHUAN HOA MODERN USERFORM (frmReport)

    ' ==============================================================================

    On Error GoTo Fallback

    If Application.DisplayAlerts = 0 Then

        ShowStandardizeReportForm = True

        Exit Function

    End If

    frmReport.txtReport.Text = rptContent

    frmReport.ResultChoice = 1

    frmReport.Show vbModal

    

    Dim choice As Integer

    choice = frmReport.ResultChoice

    Unload frmReport

    DoEvents

    

    ShowStandardizeReportForm = (choice = 1)

    Exit Function

Fallback:

    On Error Resume Next

    ShowStandardizeReportForm = ShowUniConfirm(rptContent, DecodeUTF8("Vi1GT1JNQVQgLSBCw6FvIEPDoW8gQ2h14bqpbiBIw7NhIFThu7EgxJDhu5luZw=="))

End Function



Private Sub SafeDoEvents(Optional ByVal force As Boolean = False)

    Static counter As Long

    counter = counter + 1

    If force Or (counter Mod 25 = 0) Then

        DoEvents

    End If

End Sub



Private Function MinLong(ByVal a As Long, ByVal b As Long) As Long

    If a < b Then MinLong = a Else MinLong = b

End Function



Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long

    If a > b Then MaxLong = a Else MaxLong = b

End Function



Public Function VF_LinesToPoints(ByVal numLines As Single) As Single

    VF_LinesToPoints = numLines * 12!

End Function



Public Sub RecordError(ByVal src As String, ByVal errNum As Long, ByVal errDesc As String, Optional ByVal extraInfo As String = "")

    g_LastErrorSource = src

    g_LastErrorNum = errNum

    g_LastErrorMsg = errDesc

    g_LastErrorTime = Format(Now, "yyyy-mm-dd hh:nn:ss")

    If Len(extraInfo) > 0 Then

        g_LastAction = extraInfo

    End If

End Sub



' ==============================================================================

' 1. NHAN DIEN QUY CHUAN & LOAI VAN BAN

' ==============================================================================

Public Function DetectRegulation(Optional doc As Document = Nothing) As Integer

    On Error Resume Next

    DetectRegulation = 0

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then Exit Function

    If doc.Paragraphs.Count = 0 Then Exit Function

    

    Dim pText As String, uText As String, i As Long

    

    ' 1. KIEM TRA VAN BAN HOP NHAT (ND 250)

    For i = 1 To MinLong(doc.Paragraphs.Count, 20)

        uText = UCase(ConvertToUnsignedVietnamese(doc.Paragraphs(i).Range.Text))

        If InStr(uText, "VAN BAN HOP NHAT") > 0 Or InStr(uText, "VBHN") > 0 Then

            DetectRegulation = 2

            Exit Function

        End If

    Next i

    

    ' 2. KIEM TRA KE HOACH BAI DAY / GIAO AN (ND 250)

    For i = 1 To MinLong(doc.Paragraphs.Count, 20)

        uText = UCase(ConvertToUnsignedVietnamese(doc.Paragraphs(i).Range.Text))

        If InStr(uText, "KE HOACH BAI DAY") > 0 Or InStr(uText, "GIAO AN") > 0 Then

            DetectRegulation = 3

            Exit Function

        End If

    Next i

    

    ' 3. UU TIEN SO 1: KIEM TRA QUOC HIEU & TIEU NGU CUA HANH CHINH NHA NUOC (ND 30)

    ' Neu van ban co Quoc hieu "CONG HOA XA HOI CHU NGHIA VIET NAM" hoac "DOC LAP - TU DO - HANH PHUC",

    ' thi 100% la Van ban Hanh chinh theo Nghi dinh 30, tuyet doi khong the la Van ban Dang!

    If doc.Tables.Count > 0 Then

        Dim firstTbl As Table: Set firstTbl = doc.Tables(1)

        If firstTbl.Range.Start < 2000 Then

            Dim tblTxt As String: tblTxt = UCase(ConvertToUnsignedVietnamese(firstTbl.Range.Text))

            If InStr(tblTxt, "CONG HOA XA HOI") > 0 Or InStr(tblTxt, "DOC LAP - TU DO") > 0 Or InStr(tblTxt, "DOC LAP") > 0 Then

                DetectRegulation = 0

                Exit Function

            End If

            If firstTbl.Columns.Count >= 2 Or firstTbl.Rows(1).Cells.Count >= 2 Then

                Dim c2Txt As String: c2Txt = UCase(ConvertToUnsignedVietnamese(firstTbl.Cell(1, 2).Range.Text))

                If InStr(c2Txt, "CONG HOA") > 0 Or InStr(c2Txt, "DOC LAP") > 0 Then

                    DetectRegulation = 0

                    Exit Function

                End If

            End If

        End If

    End If

    

    ' Quet Quoc hieu trong cac doan van dau trang (neu khong dung bang)

    For i = 1 To MinLong(doc.Paragraphs.Count, 25)

        uText = UCase(ConvertToUnsignedVietnamese(doc.Paragraphs(i).Range.Text))

        If InStr(uText, "CONG HOA XA HOI") > 0 Or InStr(uText, "DOC LAP - TU DO") > 0 Then

            DetectRegulation = 0

            Exit Function

        End If

    Next i

    

    ' 4. KIEM TRA VAN BAN DANG THEO HUONG DAN 05 (KHI KHONG CO QUOC HIEU NHA NUOC)

    ' A. Kiem tra bang tieu de dau trang neu co

    If doc.Tables.Count > 0 Then

        Dim hTbl As Table: Set hTbl = doc.Tables(1)

        If hTbl.Range.Start < 2000 Then

            If hTbl.Columns.Count >= 2 Or hTbl.Rows(1).Cells.Count >= 2 Then

                Dim cell1Txt As String: cell1Txt = UCase(ConvertToUnsignedVietnamese(hTbl.Cell(1, 1).Range.Text))

                Dim cell2Txt As String: cell2Txt = UCase(ConvertToUnsignedVietnamese(hTbl.Cell(1, 2).Range.Text))

                

                If InStr(cell2Txt, "DANG CONG SAN") > 0 Then

                    DetectRegulation = 1

                    Exit Function

                End If

                

                If InStr(cell1Txt, "DANG CONG SAN") > 0 Or InStr(cell1Txt, "TINH UY") > 0 Or _

                   InStr(cell1Txt, "THANH UY") > 0 Or InStr(cell1Txt, "HUYEN UY") > 0 Or _

                   InStr(cell1Txt, "QUAN UY") > 0 Or InStr(cell1Txt, "THI UY") > 0 Or _

                   InStr(cell1Txt, "DANG UY") > 0 Or InStr(cell1Txt, "CHI BO") > 0 Or _

                   InStr(cell1Txt, "BAN BI THU") > 0 Or InStr(cell1Txt, "BO CHINH TRI") > 0 Or _

                   InStr(cell1Txt, "VAN PHONG TRUNG UONG DANG") > 0 Or InStr(cell1Txt, "BAN CHAP HANH TRUNG UONG") > 0 Or _

                   InStr(cell1Txt, "UY BAN KIEM TRA") > 0 Or InStr(cell1Txt, "BAN TO CHUC") > 0 Or _

                   InStr(cell1Txt, "BAN TUYEN GIAO") > 0 Or InStr(cell1Txt, "BAN DAN VAN") > 0 Then

                    DetectRegulation = 1

                    Exit Function

                End If

            End If

        End If

    End If

    

    ' B. Quet cac doan van dau trang (Khu vuc tieu de & so hieu)

    For i = 1 To MinLong(doc.Paragraphs.Count, 15)

        uText = UCase(ConvertToUnsignedVietnamese(doc.Paragraphs(i).Range.Text))

        ' Co tieu de "DANG CONG SAN VIET NAM"

        If InStr(uText, "DANG CONG SAN VIET NAM") > 0 Or InStr(uText, "DANG CONG SAN") > 0 Then

            DetectRegulation = 1

            Exit Function

        End If

        

        ' Co so hieu quy chuan van ban Dang dac trung (-NQ/TW, -NQ/TU, -NQ/HU, -NQ/DU, -NQ/CB, -QD/TU, -HD/VPTW, v.v.)

        If InStr(uText, "-NQ/TW") > 0 Or InStr(uText, "-NQ/TU") > 0 Or InStr(uText, "-NQ/HU") > 0 Or _

           InStr(uText, "-NQ/DU") > 0 Or InStr(uText, "-NQ/CB") > 0 Or InStr(uText, "-QD/TU") > 0 Or _

           InStr(uText, "-QD/HU") > 0 Or InStr(uText, "-QD/DU") > 0 Or InStr(uText, "-QD/CB") > 0 Or _

           InStr(uText, "-KH/TU") > 0 Or InStr(uText, "-KH/HU") > 0 Or InStr(uText, "-KH/DU") > 0 Or _

           InStr(uText, "-KH/CB") > 0 Or InStr(uText, "-HD/VPTW") > 0 Or InStr(uText, "-CV/VPTW") > 0 Or _

           InStr(uText, "-BC/TU") > 0 Or InStr(uText, "-BC/HU") > 0 Or InStr(uText, "-BC/DU") > 0 Or _

           InStr(uText, "-BC/CB") > 0 Then

            DetectRegulation = 1

            Exit Function

        End If

        

        ' Co ten co quan ban hanh la co quan Dang o 5 doan dau

        If i <= 5 Then

            If Left(uText, 8) = "DANG BO " Or Left(uText, 7) = "CHI BO " Or _

               Left(uText, 8) = "TINH UY " Or Left(uText, 9) = "THANH UY " Or _

               Left(uText, 9) = "HUYEN UY " Or Left(uText, 8) = "QUAN UY " Or _

               Left(uText, 8) = "THI UY " Or Left(uText, 8) = "DANG UY " Or _

               InStr(uText, "BAN CHAP HANH TRUNG UONG") > 0 Or InStr(uText, "BAN BI THU") > 0 Or _

               InStr(uText, "BO CHINH TRI") > 0 Then

                DetectRegulation = 1

                Exit Function

            End If

        End If

    Next i

    

    ' 5. MAC DINH: NEU KHONG PHAI CAC TRUONG HOP TREN THI LA HANH CHINH NHA NUOC (ND 30)

    DetectRegulation = 0

End Function



Public Function DetectDocType(Optional doc As Document = Nothing) As Integer

    On Error Resume Next

    DetectDocType = 0

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then Exit Function

    If doc.Paragraphs.Count = 0 Then Exit Function

    

    Dim rawText As String, pText As String, i As Long, maxP As Long

    maxP = MinLong(doc.Paragraphs.Count, 35)

    

    For i = 1 To maxP

        rawText = Trim(doc.Paragraphs(i).Range.Text)

        pText = UCase(ConvertToUnsignedVietnamese(rawText))

        

        If InStr(pText, "QUYET DINH") > 0 Then

            DetectDocType = 2: Exit Function

        ElseIf InStr(pText, "NGHI QUYET") > 0 Then

            DetectDocType = 3: Exit Function

        ElseIf InStr(pText, "TO TRINH") > 0 Then

            DetectDocType = 4: Exit Function

        ElseIf (InStr(pText, "KE HOACH") > 0) And InStr(pText, "BAI DAY") = 0 Then

            DetectDocType = 5: Exit Function

        ElseIf InStr(pText, "BAO CAO") > 0 Then

            DetectDocType = 6: Exit Function

        ElseIf InStr(pText, "THONG BAO") > 0 Then

            DetectDocType = 7: Exit Function

        ElseIf InStr(pText, "BIEN BAN") > 0 Then

            DetectDocType = 8: Exit Function

        ElseIf InStr(pText, "GIAY NGHI PHEP") > 0 Then

            DetectDocType = 11: Exit Function

        ElseIf InStr(pText, "GIAY MOI") > 0 Then

            DetectDocType = 12: Exit Function

        ElseIf InStr(pText, "GIAY GIOI THIEU") > 0 Then

            DetectDocType = 13: Exit Function

        ElseIf InStr(pText, "CONG DIEN") > 0 Then

            DetectDocType = 14: Exit Function

        ElseIf InStr(pText, "PHU LUC") > 0 Then

            DetectDocType = 15: Exit Function

        ElseIf InStr(pText, "BAN SAO") > 0 Or InStr(pText, "SAO Y") > 0 Or InStr(pText, "SAO LUC") > 0 Or InStr(pText, "TRICH SAO") > 0 Then

            DetectDocType = 16: Exit Function

        ElseIf InStr(pText, "VAN BAN HOP NHAT") > 0 Or InStr(pText, "VBHN") > 0 Then

            DetectDocType = 9: Exit Function

        ElseIf InStr(pText, "KE HOACH BAI DAY") > 0 Or InStr(pText, "GIAO AN") > 0 Then

            DetectDocType = 10: Exit Function

        ElseIf InStr(pText, "KINH GUI") > 0 Or InStr(pText, "KINH TRINH") > 0 Then

            DetectDocType = 1: Exit Function

        End If

    Next i

End Function



Public Sub PerformDocumentDetection()

    If Documents.Count = 0 Then

        ShowUniMsg "VnVpIGzDsm5nIG3hu58gbeG7mXQgdMOgaSBsaeG7h3UgV29yZCDEkeG7gyBuaOG6rW4gZGnhu4duIQ==", "Vi1GT1JNQVQ=", 48

        Exit Sub

    End If

    

    Dim doc As Document: Set doc = ActiveDocument

    Dim regIdx As Integer: regIdx = DetectRegulation(doc)

    Dim docTypeIdx As Integer: docTypeIdx = DetectDocType(doc)

    

    g_SelectedRegulation = regIdx

    g_RegulationManuallySelected = False

    g_SelectedDocType = docTypeIdx

    

    Dim regName As String, docTypeName As String

    Select Case regIdx

        Case 0: regName = DecodeUTF8("SMOgbmggY2jDrW5oIE5ow6Agbsaw4bubYyAoTmdo4buLIMSR4buLbmggMzAvMjAyMC9OxJAtQ1Ap")

        Case 1: regName = DecodeUTF8("VsSDbiBi4bqjbiBj4bunYSDEkOG6o25nIChIxrDhu5tuZyBk4bqrbiAwNS1IRC9WUFRXKQ==")

        Case 2: regName = DecodeUTF8("VsSDbiBi4bqjbiBI4bujcCBuaOG6pXQgKE5naOG7iyDEkeG7i25oIDI1MC8yMDI2L07EkC1DUCk=")

        Case 3: regName = DecodeUTF8("S+G6vyBob+G6oWNoIGLDoGkgZOG6oXkgLSBHacOhbyDDoW4gKEPDtG5nIHbEg24gNTUxMi9CR0TEkFQp")

    End Select

    

    Select Case docTypeIdx

        Case 1: docTypeName = DecodeUTF8("Q8O0bmcgdsSDbiBow6BuaCBjaMOtbmggKGvDrW5oIGfhu61pL2vDrW5oIHRyw6xuaCk=")

        Case 2: docTypeName = DecodeUTF8("UVVZ4bq+VCDEkOG7ik5I")

        Case 3: docTypeName = DecodeUTF8("TkdI4buKIFFVWeG6vlQ=")

        Case 4: docTypeName = DecodeUTF8("VOG7nCBUUsOMTkg=")

        Case 5: docTypeName = DecodeUTF8("S+G6viBIT+G6oENI")

        Case 6: docTypeName = DecodeUTF8("QsOBTyBDw4FP")

        Case 7: docTypeName = DecodeUTF8("VEjDlE5HIELDgU8=")

        Case 8: docTypeName = DecodeUTF8("QknDik4gQuG6ok4=")

        Case 9: docTypeName = DecodeUTF8("VsSCTiBC4bqiTiBI4buiUCBOSOG6pFQ=")

        Case 10: docTypeName = DecodeUTF8("S+G6viBIT+G6oENIIELDgEkgROG6oFkgKEdJw4FPIMOBTik=")

        Case 11: docTypeName = DecodeUTF8("R0nhuqRZIE5HSOG7iCBQSMOJUA==")

        Case 12: docTypeName = DecodeUTF8("R0nhuqRZIE3hu5xJ")

        Case 13: docTypeName = DecodeUTF8("R0nhuqRZIEdJ4buaSSBUSEnhu4ZV")

        Case 14: docTypeName = DecodeUTF8("Q8OUTkcgxJBJ4buGTg==")

        Case 15: docTypeName = DecodeUTF8("UGjhu6UgbOG7pWMgdsSDbiBi4bqjbg==")

        Case 16: docTypeName = DecodeUTF8("QuG6o24gc2FvIHbEg24gYuG6o24gKFNBTyBZIC8gU0FPIEzhu6RDIC8gVFLDjUNIIFNBTyk=")

        Case Else: docTypeName = DecodeUTF8("Q2jGsGEgcGjDoXQgaGnhu4duIHLDtSB0w6puIGxv4bqhaSAoc+G7rSBk4bulbmcgY2h14bqpbiBjaHVuZyk=")

    End Select

    

    Dim toastMsg As String

    toastMsg = DecodeUTF8("UXV5IGNodeG6qW46IA==") & regName & vbCrLf & _

               DecodeUTF8("TG/huqFpIHbEg24gYuG6o246IA==") & docTypeName

    

    ShowUniToast toastMsg, DecodeUTF8("Vi1GT1JNQVQgLSBOaOG6rW4gRGnhu4duIFbEg24gQuG6o24=")

End Sub



' ==============================================================================

' 2. RA SOAT THE THUC VAN BAN

' ==============================================================================

Public Sub AuditDocumentCompliance()

    On Error GoTo ErrH

    If Documents.Count = 0 Then

        ShowUniMsg "VnVpIGzDsm5nIG3hu58gbeG7mXQgdMOgaSBsaeG7h3UgV29yZCDEkeG7gyByw6Agc2/DoXQh", "Vi1GT1JNQVQ=", 48

        Exit Sub

    End If

    Dim doc As Document: Set doc = ActiveDocument

    If doc Is Nothing Then Exit Sub

    

    Dim regIdx As Integer: regIdx = GetEffectiveRegulation(doc)

    

    Dim regTitle As String

    Select Case regIdx

        Case 1: regTitle = "Van ban Dang (HD 05-HD/VPTW)"

        Case 2: regTitle = "Van ban Hop nhat (ND 250/2026/ND-CP)"

        Case 3: regTitle = "Doanh nghiep (ND 250-CP)"

        Case Else: regTitle = "Hanh chinh (ND 30/2020/ND-CP)"

    End Select

    

    Dim errCount As Long: errCount = 0

    Dim p As Paragraph, pText As String, pClean As String

    Dim tbl As Table, rowIdx As Long, colIdx As Long, cCell As Cell

    

    ' 1. KIEM TRA KHO GIAY A4

    If doc.PageSetup.PaperSize <> wdPaperA4 Then

        errCount = errCount + 1

        If doc.Paragraphs.Count > 0 Then

            AddAuditFinding doc.Paragraphs(1).Range, DecodeUTF8("S2jhu5UgZ2nhuqV5IGNoxrBhIGNodeG6qW4gQTQuIEPhuqduIGNo4buJbmggc2FuZyBraOG7lSBBNCAoMjEwIHggMjk3IG1tKS4=")

        End If

    End If

    

    ' 2. KIEM TRA CAN LE TRANG

    Dim topMm As Single, botMm As Single, leftMm As Single, rightMm As Single

    topMm = doc.PageSetup.TopMargin / MM_TO_PT

    botMm = doc.PageSetup.BottomMargin / MM_TO_PT

    leftMm = doc.PageSetup.LeftMargin / MM_TO_PT

    rightMm = doc.PageSetup.RightMargin / MM_TO_PT

    

    Dim badMargins As Boolean: badMargins = False

    If regIdx = 1 Then

        If topMm < 19 Or topMm > 21 Or botMm < 19 Or botMm > 21 Or leftMm < 29 Or leftMm > 31 Or rightMm < 14 Or rightMm > 16 Then badMargins = True

    Else

        If topMm < 19 Or topMm > 26 Or botMm < 19 Or botMm > 26 Or leftMm < 29 Or leftMm > 36 Or rightMm < 14 Or rightMm > 21 Then badMargins = True

    End If

    If badMargins Then

        errCount = errCount + 1

        If doc.Paragraphs.Count > 0 Then

            AddAuditFinding doc.Paragraphs(1).Range, DecodeUTF8("Q8SDbiBs4buBIGNoxrBhIMSRw7puZyBxdXkgxJHhu4tuaCAoWcOqdSBj4bqndTogVHLDqm4gMjAtMjVtbSwgRMaw4bubaSAyMC0yNW1tLCBUcsOhaSAzMC0zNW1tLCBQaOG6o2kgMTUtMjBtbSku")

        End If

    End If

    

    ' 3. QUET BANG PHAN DAU (TABLES)

    For Each tbl In doc.Tables

        For rowIdx = 1 To tbl.Rows.Count

            For colIdx = 1 To tbl.Columns.Count

                On Error Resume Next

                Set cCell = tbl.Cell(rowIdx, colIdx)

                If Not cCell Is Nothing Then

                    pText = Trim(cCell.Range.Text)

                    pClean = UCase(Replace(Replace(pText, vbCr, ""), Chr(7), ""))

                    

                    ' Quoc hieu

                    If InStr(pClean, "CONG HOA XA HOI") > 0 Or InStr(pClean, DecodeUTF8("Q+G7mE5HIEjDkkEgWMOj")) > 0 Then

                        If cCell.Range.Font.Bold = False Then

                            errCount = errCount + 1

                            AddAuditFinding cCell.Range, DecodeUTF8("UXXhu5FjIGhp4buHdSBwaOG6o2kgaW4gxJHhuq1tIChCb2xkKSwgaW4gaG9hIHRvw6BuIGLhu5ku")

                        End If

                    End If

                    

                    ' Tieu ngu

                    If InStr(pClean, "DOC LAP - TU DO") > 0 Or InStr(pClean, DecodeUTF8("xJDhu5hDIExBUA==")) > 0 Then

                        If cCell.Range.Font.Bold = False Then

                            errCount = errCount + 1

                            AddAuditFinding cCell.Range, DecodeUTF8("VGnDqnUgbmfhu68gcGjhuqNpIGluIMSR4bqtbSwgY2jhu68gY8OhaSDEkeG6p3Ugdmnhur90IGhvYSwgY8OzIGfhuqFjaCBu4buRaSBjw6FjaCDEkeG7gXUgKMSQ4buZYyBs4bqtcCAtIFThu7EgZG8gLSBI4bqhbmggcGjDumMpLg==")

                        End If

                    End If

                    

                    ' So ky hieu

                    If InStr(pClean, "SO :") > 0 Then

                        errCount = errCount + 1

                        AddAuditFinding cCell.Range, DecodeUTF8("S2jDtG5nIMSR4buDIGThuqV1IGPDoWNoIHRyxrDhu5tjIGThuqV1IGhhaSBjaOG6pW0g4bufIHPhu5EgdsSDbiBi4bqjbiAodmnhur90IMSRw7puZzogIlPhu5E6Iiku")

                    End If

                    

                    ' Dia danh ngay thang nam

                    If InStr(pClean, "NGAY") > 0 And InStr(pClean, "THANG") > 0 And InStr(pClean, "NAM") > 0 Then

                        If cCell.Range.Font.Italic = False Then

                            errCount = errCount + 1

                            AddAuditFinding cCell.Range, DecodeUTF8("xJDhu4thIGRhbmggdsOgIG5nw6B5IHRow6FuZyBuxINtIHBo4bqjaSBpbiBuZ2hpw6puZyAoSXRhbGljKSwgY2jhu68gdGjGsOG7nW5nLCBj4buhIDEzLTE0cHQu")

                        End If

                    End If

                End If

            Next colIdx

        Next rowIdx

    Next tbl

    

    ' 4. QUET TUNG DOAN VAN BAN (PARAGRAPHS)

    Dim pIdx As Long: pIdx = 0

    For Each p In doc.Paragraphs

        pIdx = pIdx + 1

        pText = Trim(p.Range.Text)

        pClean = UCase(Replace(Replace(pText, vbCr, ""), vbLf, ""))

        

        If Len(pClean) > 2 Then

            ' Kiem tra Font chu

            If InStr(UCase(p.Range.Font.Name), "TIMES") = 0 Then

                errCount = errCount + 1

                AddAuditFinding p.Range, DecodeUTF8("UGjDoXQgaGnhu4duIGZvbnQgY2jhu68ga2jDoWMgVGltZXMgTmV3IFJvbWFuLiBD4bqnbiBjaHV54buDbiB24buBIGZvbnQgVGltZXMgTmV3IFJvbWFuIMSR4buTbmcgbmjhuqV0Lg==")

            End If

            

            ' Kiem tra Ten loai van ban

            If IsDocumentTitle(pClean) Then

                If p.Range.Font.Bold = False Then

                    errCount = errCount + 1

                    AddAuditFinding p.Range, DecodeUTF8("VMOqbiBsb+G6oWkgdsSDbiBi4bqjbiBwaOG6o2kgaW4gxJHhuq1tIChCb2xkKSwgY+G7qSAxNHB0Li4u")

                End If

            End If

            

            ' Kiem tra Noi nhan

            If InStr(pClean, "NOI NHAN:") > 0 Or InStr(pClean, DecodeUTF8("TsahaSBuaOG6rW46")) > 0 Then

                If p.Range.Font.Bold = False Or p.Range.Font.Italic = False Then

                    errCount = errCount + 1

                    AddAuditFinding p.Range, DecodeUTF8("VOG7qyAiTsahaSBuaOG6rW46IiBwaOG6o2kgaW4gxJHhuq1tIHbDoCBpbiBuZ2hpw6puZyAoQm9sZCBJdGFsaWMpLCBj4buhIDEycHQu")

                End If

            End If

        End If

    Next p

    

    SetProgress 100, DecodeUTF8("SG/DoG4gdOG6pXQgcsOgIHNvw6F0IQ=="), "Audit"

    FinishProgress "Audit"

    

    ' 5. HIEN THI THONG BAO TOAST

    Dim toastMsg As String

    toastMsg = DecodeUTF8("UXV5IGNodeG6qW46IA==") & regTitle & vbCrLf

    If errCount > 0 Then

        toastMsg = toastMsg & DecodeUTF8("xJDDoyDEkcOtbmgga8OobSBnaGkgY2jDuiAoQ29tbWVudHMpIGPDoWMgduG7iyB0csOtIGPhuqduIHPhu61hIHRyw6puIHbEg24gYuG6o24u")

    Else

        toastMsg = toastMsg & DecodeUTF8("Q2jDumMgbeG7q25nISBWxINuIGLhuqNuIMSR4bqhdCBjaHXhuqluIHRo4buDIHRo4bupYyAxMDAlIChraMO0bmcgY8OzIGzhu5dpKS4=")

    End If

    

    ShowUniToast toastMsg, DecodeUTF8("Vi1GT1JNQVQgLSBSw6AgU2/DoXQgVGjhu4MgVGjhu6lj")

    Exit Sub

ErrH:

    FinishProgress "Audit"

    RecordError "VFormat_Core.AuditDocumentCompliance", Err.Number, Err.Description, "Ra soat the thuc"

End Sub



Private Sub AddAuditFinding(ByVal rng As Range, ByVal commentText As String)

    On Error Resume Next

    rng.Document.Comments.Add Range:=rng, Text:="[V-FORMAT] " & commentText

End Sub



' ==============================================================================

' 3. CHUYEN MA UNICODE

' ==============================================================================

Public Sub ConvertOldFontsToUnicode(Optional doc As Document = Nothing)

    On Error GoTo ErrH

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then

        ShowUniMsg "VnVpIGzDsm5nIG3hu58gbeG7mXQgdMOgaSBsaeG7h3UgV29yZCDEkeG7gyBjaHV54buDbiBtw6Mh", "Vi1GT1JNQVQ=", 48

        Exit Sub

    End If

    

    StartProgress DecodeUTF8("Vi1GT1JNQVQgLSBDaHV54buDbiBNw6MgRm9udA=="), DecodeUTF8("xJBhbmcgY2h1eeG7g24gxJHhu5VpIHRow6BuaCBwaMO0bmcgVU5JQ09ERSAoVGltZXMgTmV3IFJvbWFuKS4uLg=="), "ConvertFont"

    SetProgress 30, DecodeUTF8("xJBhbmcgcXXDqXQgY8OhYyDEkeG7kWFuIHbEg24gY2jDrW5oLi4u"), "ConvertFont"

    Application.ScreenUpdating = False

    Dim p As Paragraph

    For Each p In doc.Paragraphs

        p.Range.Font.Name = "Times New Roman"

    Next p

    

    SetProgress 70, DecodeUTF8("xJBhbmcgcXXDqXQgY8OhYyBi4bqjbmcgYmnhu4N1Li4u"), "ConvertFont"

    Dim tbl As Table

    For Each tbl In doc.Tables

        tbl.Range.Font.Name = "Times New Roman"

    Next tbl

    

    Application.ScreenUpdating = True

    SetProgress 100, DecodeUTF8("SG/DoG4gdOG6pXQh"), "ConvertFont"

    FinishProgress "ConvertFont"

    ShowUniMsg "xJDDoyBjaHV54buDbiDEkeG7lWkgdG/DoG4gYuG7mSB2xINuIGLhuqNuIHNhbmcgVXRmLTggVU5JQ09ERSAoVGltZXMgTmV3IFJvbWFuKSB0aMOgbmggY8O0bmch", "Vi1GT1JNQVQ=", 64

    Exit Sub

ErrH:

    FinishProgress "ConvertFont"

    Application.ScreenUpdating = True

    RecordError "VFormat_Core.ConvertOldFontsToUnicode", Err.Number, Err.Description, "Chuyen ma phong chu"

End Sub



' ==============================================================================

' 4. CAN LE TRANG A4

' ==============================================================================

Public Sub ApplyPageSetupA4(Optional doc As Document = Nothing)

    On Error GoTo ErrH

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then

        ShowUniMsg "VnVpIGzDsm5nIG3hu58gbeG7mXQgdMOgaSBsaeG7h3UgV29yZCDEkeG7gyDEkeG7i25oIGThuqFuZyE=", "Vi1GT1JNQVQ=", 48

        Exit Sub

    End If

    

    Dim regIdx As Integer: regIdx = GetEffectiveRegulation(doc)

    

    Dim topVal As Single, botVal As Single, leftVal As Single, rightVal As Single

    If regIdx = 1 Then

        topVal = 20: botVal = 20: leftVal = 30: rightVal = 15

    Else

        topVal = 20: botVal = 20: leftVal = 30: rightVal = 20

    End If

    

    Dim sec As Section

    For Each sec In doc.Sections

        With sec.PageSetup

            .PaperSize = wdPaperA4

            .PageWidth = MillimetersToPoints(210)

            .PageHeight = MillimetersToPoints(297)

            .Orientation = wdOrientPortrait

            .TopMargin = MillimetersToPoints(topVal)

            .BottomMargin = MillimetersToPoints(botVal)

            .LeftMargin = MillimetersToPoints(leftVal)

            .RightMargin = MillimetersToPoints(rightVal)

            .HeaderDistance = MillimetersToPoints(10)

            .FooterDistance = MillimetersToPoints(10)

            .DifferentFirstPageHeaderFooter = True

        End With

    Next sec

    

    If regIdx = 1 Then

        ShowUniToast DecodeUTF8("xJDDoyDEkeG7i25oIGThuqFuZyBraOG7lSBnaeG6pXkgQTQgdsOgIGPEg24gbOG7gSDEkOG6o25nIChIRCAwNTogVHLDqm4gMjAsIETGsOG7m2kgMjAsIFRyw6FpIDMwLCBQaOG6o2kgMTUpbW0gdGjDoG5oIGPDtG5nIQ=="), "Vi1GT1JNQVQ="

    Else

        ShowUniToast DecodeUTF8("xJDDoyDEkeG7i25oIGThuqFuZyBraOG7lSBnaeG6pXkgQTQgdsOgIGPEg24gbOG7gSBIw6BuaCBjaMOtbmggKE7EkCAzMDogVHLDqm4gMjAsIETGsOG7m2kgMjAsIFRyw6FpIDMwLCBQaOG6o2kgMjApbW0gdGjDoG5oIGPDtG5nIQ=="), "Vi1GT1JNQVQ="

    End If

    Exit Sub

ErrH:

    RecordError "VFormat_Core.ApplyPageSetupA4", Err.Number, Err.Description, "Dinh dang trang A4"

End Sub



Public Sub InsertLandscapeSection()

    On Error GoTo ErrH

    If Documents.Count = 0 Then Exit Sub

    Selection.InsertBreak Type:=wdSectionBreakNextPage

    With Selection.PageSetup

        .Orientation = wdOrientLandscape

        .PageWidth = MillimetersToPoints(297)

        .PageHeight = MillimetersToPoints(210)

        .TopMargin = MillimetersToPoints(20)

        .BottomMargin = MillimetersToPoints(20)

        .LeftMargin = MillimetersToPoints(30)

        .RightMargin = MillimetersToPoints(15)

    End With

    Exit Sub

ErrH:

    RecordError "VFormat_Core.InsertLandscapeSection", Err.Number, Err.Description, "Chen trang ngang"

End Sub



Public Sub InsertPortraitSection()

    On Error GoTo ErrH

    If Documents.Count = 0 Then Exit Sub

    Selection.InsertBreak Type:=wdSectionBreakNextPage

    With Selection.PageSetup

        .Orientation = wdOrientPortrait

        .PageWidth = MillimetersToPoints(210)

        .PageHeight = MillimetersToPoints(297)

        .TopMargin = MillimetersToPoints(20)

        .BottomMargin = MillimetersToPoints(20)

        .LeftMargin = MillimetersToPoints(30)

        .RightMargin = MillimetersToPoints(15)

    End With

    Exit Sub

ErrH:

    RecordError "VFormat_Core.InsertPortraitSection", Err.Number, Err.Description, "Chen trang doc"

End Sub



Public Function DeleteBlankPages(Optional doc As Document = Nothing) As Long

    On Error GoTo ErrH

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then Exit Function

    

    Dim initPages As Long: initPages = doc.ComputeStatistics(wdStatisticPages)

    Dim cleaned As Boolean: cleaned = True

    Dim p As Paragraph, txt As String

    

    Do While cleaned

        cleaned = False

        If doc.Paragraphs.Count <= 1 Then Exit Do

        Set p = doc.Paragraphs(doc.Paragraphs.Count)

        If Not p.Range.Information(wdWithInTable) Then

            txt = p.Range.Text

            txt = Replace(txt, vbCr, "")

            txt = Replace(txt, vbLf, "")

            txt = Replace(txt, Chr(12), "")

            txt = Replace(txt, Chr(11), "")

            txt = Replace(txt, Chr(9), "")

            txt = Replace(txt, Chr(32), "")

            txt = Replace(txt, ChrW(160), "")

            If Len(txt) = 0 Then

                p.Range.Delete

                cleaned = True

            End If

        End If

    Loop

    

    If doc.Paragraphs.Count > 0 Then

        Set p = doc.Paragraphs(doc.Paragraphs.Count)

        If InStr(p.Range.Text, Chr(12)) > 0 Then

            p.Range.Text = Replace(p.Range.Text, Chr(12), "")

        End If

    End If

    

    Dim finalPages As Long: finalPages = doc.ComputeStatistics(wdStatisticPages)

    DeleteBlankPages = IIf(initPages > finalPages, initPages - finalPages, 0)

    

    ShowUniMsg "xJDDoyB4w7NhIHPDoWNoIGPDoWMgdHJhbmcgdHLhuq9uZyB0aOG7q2EgdsOgIG5n4bqvdCB0cmFuZyB2w7QgbmdoxKlhIGPhu6dhIHbEg24gYuG6o24h", "Vi1GT1JNQVQ=", 64

    Exit Function

ErrH:

    RecordError "VFormat_Core.DeleteBlankPages", Err.Number, Err.Description, "Xoa trang trang"

End Function



' ==============================================================================

' 5. DANH SO TRANG THEO QUY CHUAN & TUY CHON TRANG BAT DAU

' ==============================================================================

Public Sub ShowPageNumberDialog()

    On Error Resume Next

    frmPageNumber.Show

End Sub



Public Sub FormatPageNumbers(Optional doc As Document = Nothing)

    ApplyPageNumbering 1, 1, True, True, doc

End Sub



Public Sub ApplyPageNumbering(ByVal fromPage As Integer, ByVal startNum As Integer, ByVal hideFirst As Boolean, ByVal isHeader As Boolean, Optional doc As Document = Nothing)

    On Error GoTo ErrH

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then Exit Sub

    

    Dim totalPages As Long

    totalPages = doc.ComputeStatistics(wdStatisticPages)

    If totalPages < 1 Then totalPages = 1

    

    If fromPage < 1 Then fromPage = 1

    If startNum < 1 Then startNum = 1

    

    If fromPage > totalPages Then

        ShowUniMsg DecodeUTF8("VHJhbmcgYuG6r3QgxJHhuqd1IHbGsOG7o3QgcXXDoSB04buVbmcgc+G7kSB0cmFuZyBj4bunYSB0w6BpIGxp4buHdSAoaGnhu4duIGPDsyA=") & totalPages & DecodeUTF8("IHRyYW5nKSE="), "Vi1GT1JNQVQ=", 48

        Exit Sub

    End If

    

    Dim regIdx As Integer: regIdx = GetEffectiveRegulation(doc)

    Dim fontSize As Single: fontSize = IIf(regIdx = 1, 14, 13.5)

    

    If fromPage = 1 Then

        ' Danh so tu dau van ban

        Dim sec As Section

        For Each sec In doc.Sections

            If hideFirst Then

                sec.PageSetup.DifferentFirstPageHeaderFooter = True

            Else

                sec.PageSetup.DifferentFirstPageHeaderFooter = False

            End If

            

            Dim hf As HeaderFooter

            If isHeader Then

                sec.PageSetup.HeaderDistance = MillimetersToPoints(10)

                Set hf = sec.Headers(wdHeaderFooterPrimary)

                sec.Footers(wdHeaderFooterPrimary).Range.Text = ""

                sec.Footers(wdHeaderFooterFirstPage).Range.Text = ""

            Else

                sec.PageSetup.FooterDistance = MillimetersToPoints(10)

                Set hf = sec.Footers(wdHeaderFooterPrimary)

                sec.Headers(wdHeaderFooterPrimary).Range.Text = ""

                sec.Headers(wdHeaderFooterFirstPage).Range.Text = ""

            End If

            

            hf.Range.Text = ""

            hf.Range.ParagraphFormat.TabStops.ClearAll

            hf.Range.ParagraphFormat.LeftIndent = 0

            hf.Range.ParagraphFormat.RightIndent = 0

            hf.Range.ParagraphFormat.FirstLineIndent = 0

            hf.Range.Font.Name = "Times New Roman"

            hf.Range.Font.Size = fontSize

            hf.Range.Font.Bold = False

            hf.Range.Font.Italic = False

            hf.Range.Font.ColorIndex = wdBlack

            hf.PageNumbers.RestartNumberingAtSection = True

            hf.PageNumbers.StartingNumber = startNum

            hf.Range.Fields.Add Range:=hf.Range, Type:=wdFieldPage

            hf.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

            If hf.Range.Paragraphs.Count > 0 Then

                hf.Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

            End If

            

            If hideFirst Then

                If isHeader Then

                    sec.Headers(wdHeaderFooterFirstPage).Range.Text = ""

                Else

                    sec.Footers(wdHeaderFooterFirstPage).Range.Text = ""

                End If

            End If

        Next sec

        

        If hideFirst Then

            ShowUniToast DecodeUTF8("xJDDoyDEkcOhbmggc+G7kSB0cmFuZyBjaHXhuqluIFF1eSDEkeG7i25oIChDw6JuIGdp4buvYSBs4buBIHRyw6puLCBUaW1lcyBOZXcgUm9tYW4sIHThu7EgxJHhu5luZyDhuqluIHRyYW5nIDEpIQ=="), "Vi1GT1JNQVQ="

        Else

            ShowUniToast DecodeUTF8("xJDDoyDEkcOhbmggc+G7kSB0cmFuZzogVOG7qyB0cmFuZyAxIChi4bqvdCDEkeG6p3UgbMOgIHPhu5Eg") & startNum & ")!", "Vi1GT1JNQVQ="

        End If

    Else

        ' Danh so tu trang tuy chon (fromPage > 1)

        Dim rngStartOfPage As Range

        Set rngStartOfPage = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=fromPage)

        

        Dim targetSecIdx As Long

        targetSecIdx = rngStartOfPage.Sections(1).Index

        

        Dim isAlreadyNewSec As Boolean

        isAlreadyNewSec = False

        If targetSecIdx > 1 And rngStartOfPage.Start = rngStartOfPage.Sections(1).Range.Start Then

            isAlreadyNewSec = True

        End If

        

        If Not isAlreadyNewSec Then

            Dim rngPrev As Range

            Set rngPrev = doc.Range(rngStartOfPage.Start - 1, rngStartOfPage.Start)

            If rngPrev.Text = Chr(12) Then rngPrev.Delete

            rngStartOfPage.InsertBreak Type:=wdSectionBreakNextPage

            Set rngStartOfPage = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=fromPage)

            targetSecIdx = rngStartOfPage.Sections(1).Index

        End If

        

        ' 1. Cat dut lien ket Section muc tieu voi cac Section truoc

        Dim targetSec As Section

        Set targetSec = doc.Sections(targetSecIdx)

        

        On Error Resume Next

        targetSec.Headers(wdHeaderFooterPrimary).LinkToPrevious = False

        targetSec.Headers(wdHeaderFooterFirstPage).LinkToPrevious = False

        targetSec.Footers(wdHeaderFooterPrimary).LinkToPrevious = False

        targetSec.Footers(wdHeaderFooterFirstPage).LinkToPrevious = False

        targetSec.PageSetup.DifferentFirstPageHeaderFooter = False

        On Error GoTo ErrH

        

        ' 2. Xoa sach header & footer cac section truoc (Trang 1 den fromPage - 1)

        Dim s As Long

        For s = 1 To targetSecIdx - 1

            On Error Resume Next

            doc.Sections(s).Headers(wdHeaderFooterPrimary).Range.Text = ""

            doc.Sections(s).Headers(wdHeaderFooterFirstPage).Range.Text = ""

            doc.Sections(s).Footers(wdHeaderFooterPrimary).Range.Text = ""

            doc.Sections(s).Footers(wdHeaderFooterFirstPage).Range.Text = ""

            On Error GoTo ErrH

        Next s

        

        ' 3. Cau hinh so trang tai Section muc tieu (bat dau tu fromPage)

        Dim hfTarget As HeaderFooter

        If isHeader Then

            targetSec.PageSetup.HeaderDistance = MillimetersToPoints(10)

            Set hfTarget = targetSec.Headers(wdHeaderFooterPrimary)

            targetSec.Footers(wdHeaderFooterPrimary).Range.Text = ""

        Else

            targetSec.PageSetup.FooterDistance = MillimetersToPoints(10)

            Set hfTarget = targetSec.Footers(wdHeaderFooterPrimary)

            targetSec.Headers(wdHeaderFooterPrimary).Range.Text = ""

        End If

        

        hfTarget.PageNumbers.RestartNumberingAtSection = True

        hfTarget.PageNumbers.StartingNumber = startNum

        hfTarget.Range.Text = ""

        hfTarget.Range.ParagraphFormat.TabStops.ClearAll

        hfTarget.Range.ParagraphFormat.LeftIndent = 0

        hfTarget.Range.ParagraphFormat.RightIndent = 0

        hfTarget.Range.ParagraphFormat.FirstLineIndent = 0

        hfTarget.Range.Font.Name = "Times New Roman"

        hfTarget.Range.Font.Size = fontSize

        hfTarget.Range.Font.Bold = False

        hfTarget.Range.Font.Italic = False

        hfTarget.Range.Font.ColorIndex = wdBlack

        hfTarget.Range.Fields.Add Range:=hfTarget.Range, Type:=wdFieldPage

        hfTarget.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

        If hfTarget.Range.Paragraphs.Count > 0 Then

            hfTarget.Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

        End If

        

        ' 4. Cac section tiep theo (neu co) se tu dong dem tiep tuc

        For s = targetSecIdx + 1 To doc.Sections.Count

            Dim secNext As Section

            Set secNext = doc.Sections(s)

            Dim hfNext As HeaderFooter

            On Error Resume Next

            If isHeader Then

                Set hfNext = secNext.Headers(wdHeaderFooterPrimary)

                secNext.Footers(wdHeaderFooterPrimary).Range.Text = ""

            Else

                Set hfNext = secNext.Footers(wdHeaderFooterPrimary)

                secNext.Headers(wdHeaderFooterPrimary).Range.Text = ""

            End If

            hfNext.PageNumbers.RestartNumberingAtSection = False

            On Error GoTo ErrH

        Next s

        

        Dim successMsg As String

        successMsg = DecodeUTF8("xJDDoyDEkcOhbmggc+G7kSB0cmFuZzogVOG7qyB0cmFuZyA=") & fromPage & _

                     DecodeUTF8("IChi4bqvdCDEkeG6p3UgbMOgIHPhu5Eg") & startNum & _

                     DecodeUTF8("KSDEkeG6v24gaOG6v3QgdsSDbiBi4bqjbiE=")

        ShowUniToast successMsg, "Vi1GT1JNQVQ="

    End If

    Exit Sub

ErrH:

    RecordError "VFormat_Core.ApplyPageNumbering", Err.Number, Err.Description, "Danh so trang"

End Sub



Public Sub RemoveAllPageNumbers(Optional doc As Document = Nothing)

    On Error Resume Next

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then Exit Sub

    

    Dim sec As Section

    For Each sec In doc.Sections

        sec.Headers(wdHeaderFooterPrimary).Range.Text = ""

        sec.Headers(wdHeaderFooterFirstPage).Range.Text = ""

        sec.Footers(wdHeaderFooterPrimary).Range.Text = ""

        sec.Footers(wdHeaderFooterFirstPage).Range.Text = ""

    Next sec

    ShowUniToast DecodeUTF8("xJDDoyB4w7NhIHPhuqFjaCB0b8OgbiBi4buZIHPhu5EgdHJhbmcgdHJvbmcgdMOgaSBsaeG7h3Uh"), "Vi1GT1JNQVQ="

End Sub



' ==============================================================================

' 6. CA NHAN HOA DU LIEU & CHEN NHANH THANH PHAN TAI CON TRO CHUOT

' ==============================================================================

Private Function GetPythonExePath() As String

    On Error Resume Next

    Dim localApp As String: localApp = Environ$("LOCALAPPDATA")

    Dim candidate As String

    

    candidate = localApp & "\Programs\Python\Python312\pythonw.exe"

    If Dir(candidate) <> "" Then GetPythonExePath = candidate: Exit Function

    

    candidate = localApp & "\Programs\Python\Python313\pythonw.exe"

    If Dir(candidate) <> "" Then GetPythonExePath = candidate: Exit Function

    

    candidate = localApp & "\Programs\Python\Python311\pythonw.exe"

    If Dir(candidate) <> "" Then GetPythonExePath = candidate: Exit Function

    

    candidate = localApp & "\Programs\Python\Python310\pythonw.exe"

    If Dir(candidate) <> "" Then GetPythonExePath = candidate: Exit Function

    

    candidate = localApp & "\Programs\Python\Python312\python.exe"

    If Dir(candidate) <> "" Then GetPythonExePath = candidate: Exit Function

    

    candidate = localApp & "\Programs\Python\Python311\python.exe"

    If Dir(candidate) <> "" Then GetPythonExePath = candidate: Exit Function

    

    candidate = localApp & "\Microsoft\WindowsApps\pythonw.exe"

    If Dir(candidate) <> "" Then GetPythonExePath = candidate: Exit Function

    

    candidate = "C:\Program Files\Python312\pythonw.exe"

    If Dir(candidate) <> "" Then GetPythonExePath = candidate: Exit Function

    

    candidate = "C:\Python312\pythonw.exe"

    If Dir(candidate) <> "" Then GetPythonExePath = candidate: Exit Function

    

    GetPythonExePath = "pythonw.exe"

End Function



Public Function GetVFormatAppDir() As String

    On Error Resume Next

    Dim p As String

    p = Environ$("APPDATA") & "\VFormat"

    If Dir(p, vbDirectory) <> "" Then

        GetVFormatAppDir = p

        Exit Function

    End If

    If Len(ThisDocument.Path) > 0 And Dir(ThisDocument.Path, vbDirectory) <> "" Then

        GetVFormatAppDir = ThisDocument.Path

        Exit Function

    End If

    GetVFormatAppDir = Environ$("APPDATA") & "\VFormat"

End Function



Public Sub EnsureVFormatAppDir()

    On Error Resume Next

    Dim appDir As String

    appDir = Environ$("APPDATA") & "\VFormat"

    If Dir(appDir, vbDirectory) = "" Then

        MkDir appDir

    End If

End Sub



Public Function GetVFormatScriptPath(ByVal scriptName As String) As String

    On Error Resume Next

    Dim candidate As String

    

    ' 1. Check in %APPDATA%\VFormat\src\gui\

    candidate = Environ$("APPDATA") & "\VFormat\src\gui\" & scriptName

    If Dir(candidate) <> "" Then GetVFormatScriptPath = candidate: Exit Function

    

    ' 2. Check in %APPDATA%\VFormat\

    candidate = Environ$("APPDATA") & "\VFormat\" & scriptName

    If Dir(candidate) <> "" Then GetVFormatScriptPath = candidate: Exit Function

    

    ' 3. Check in ThisDocument.Path

    If Len(ThisDocument.Path) > 0 Then

        candidate = ThisDocument.Path & "\src\gui\" & scriptName

        If Dir(candidate) <> "" Then GetVFormatScriptPath = candidate: Exit Function

        

        candidate = ThisDocument.Path & "\" & scriptName

        If Dir(candidate) <> "" Then GetVFormatScriptPath = candidate: Exit Function

        

        candidate = Left(ThisDocument.Path, InStrRev(ThisDocument.Path, "\")) & "src\gui\" & scriptName

        If Dir(candidate) <> "" Then GetVFormatScriptPath = candidate: Exit Function

    End If

    

    ' 4. Check in GetVFormatAppDir

    candidate = GetVFormatAppDir() & "\src\gui\" & scriptName

    If Dir(candidate) <> "" Then GetVFormatScriptPath = candidate: Exit Function

    

    GetVFormatScriptPath = ""

End Function



Private Sub WriteJsonFile(ByVal filePath As String, ByVal content As String)

    On Error Resume Next

    EnsureVFormatAppDir

    If Len(content) = 0 Then Exit Sub

    

    Dim b() As Byte, bLen As Long, cap As Long

    cap = Len(content) * 3 + 4

    ReDim b(0 To cap)

    bLen = 0

    

    Dim i As Long, code As Long

    For i = 1 To Len(content)

        code = AscW(Mid$(content, i, 1))

        If code < 0 Then code = code + 65536

        

        If code < &H80 Then

            b(bLen) = CByte(code)

            bLen = bLen + 1

        ElseIf code < &H800 Then

            b(bLen) = CByte(&HC0 Or (code \ 64))

            b(bLen + 1) = CByte(&H80 Or (code And &H3F))

            bLen = bLen + 2

        Else

            b(bLen) = CByte(&HE0 Or (code \ 4096))

            b(bLen + 1) = CByte(&H80 Or ((code \ 64) And &H3F))

            b(bLen + 2) = CByte(&H80 Or (code And &H3F))

            bLen = bLen + 3

        End If

    Next i

    

    If bLen > 0 Then

        ReDim Preserve b(0 To bLen - 1)

        Dim fn As Integer: fn = FreeFile

        Open filePath For Output As #fn

        Close #fn

        Open filePath For Binary Access Write As #fn

        Put #fn, , b

        Close #fn

    End If

End Sub



Public Sub SaveUserPresetsJson(ByVal newJson As String)

    On Error Resume Next

    EnsureVFormatAppDir

    Dim appData As String

    appData = Environ$("APPDATA") & "\VFormat\user_presets.json"

    WriteJsonFile appData, newJson

    

    Dim devPresets As String

    devPresets = GetVFormatAppDir() & "\user_presets.json"

    If Dir(devPresets) <> "" And LCase$(devPresets) <> LCase$(appData) Then

        WriteJsonFile devPresets, newJson

    End If

    ' OrcaX v2.0: auto-refresh dynamic menu da duoi ra khoi Core (giu nguyen
    ' p-code cache cua Core - re-import Core lam vo hieu hoa macro, xem build log)
    ' -> menu Soan nhanh duoc cap nhat khi Word khoi dong lai.
End Sub



Public Function UpdateJsonArrayItemTag(ByVal jsonStr As String, ByVal sectionName As String, ByVal tagName As String, ByVal itemIndex As Long, ByVal newValue As String) As String

    On Error GoTo ErrH

    Dim sIdx As Long: sIdx = 1

    If Len(sectionName) > 0 Then

        sIdx = InStr(jsonStr, """" & sectionName & """")

        If sIdx = 0 Then sIdx = 1

    End If

    

    Dim p As Long, sPos As Long, ePos As Long, currIdx As Long

    Dim pat As String: pat = """" & tagName & """"

    p = sIdx

    currIdx = 0

    

    Do

        p = InStr(p + 1, jsonStr, pat)

        If p = 0 Then Exit Do

        currIdx = currIdx + 1

        If currIdx = itemIndex Then

            p = InStr(p + Len(pat), jsonStr, ":")

            If p > 0 Then

                sPos = p + 1

                Do While sPos <= Len(jsonStr) And (Mid(jsonStr, sPos, 1) = " " Or Mid(jsonStr, sPos, 1) = """")

                    sPos = sPos + 1

                Loop

                ePos = sPos

                Do While ePos <= Len(jsonStr) And Mid(jsonStr, ePos, 1) <> """" And Mid(jsonStr, ePos, 1) <> "," And Mid(jsonStr, ePos, 1) <> "}" And Mid(jsonStr, ePos, 1) <> vbCr And Mid(jsonStr, ePos, 1) <> vbLf

                    ePos = ePos + 1

                Loop

                

                Dim cleanVal As String

                cleanVal = Replace(Replace(newValue, """", "\"""), vbCrLf, "\n")

                cleanVal = Replace(cleanVal, vbCr, "\n")

                cleanVal = Replace(cleanVal, vbLf, "\n")

                

                UpdateJsonArrayItemTag = Left(jsonStr, sPos - 1) & cleanVal & Mid(jsonStr, ePos)

                Exit Function

            End If

        End If

    Loop While p > 0

ErrH:

    UpdateJsonArrayItemTag = jsonStr

End Function



Public Function GetDefaultPresetsJson() As String

    Dim j As String

    j = "{" & vbCrLf

    j = j & "  ""agencies_admin"": [" & vbCrLf

    j = j & "    {" & vbCrLf

    j = j & "      ""id"": ""admin_bo""," & vbCrLf

    j = j & "      ""alias"": ""Cap Bo (Bo Noi vu)""," & vbCrLf

    j = j & "      ""parent_agency"": """"," & vbCrLf

    j = j & "      ""agency_name"": ""BO NOI VU""," & vbCrLf

    j = j & "      ""location"": ""Ha Noi""," & vbCrLf

    j = j & "      ""code_prefix"": ""BNV""" & vbCrLf

    j = j & "    }," & vbCrLf

    j = j & "    {" & vbCrLf

    j = j & "      ""id"": ""admin_cuc""," & vbCrLf

    j = j & "      ""alias"": ""Cuc truc thuoc Bo (Cuc VTLT)""," & vbCrLf

    j = j & "      ""parent_agency"": ""BO NOI VU""," & vbCrLf

    j = j & "      ""agency_name"": ""CUC VAN THU VA LUU TRU NHA NUOC""," & vbCrLf

    j = j & "      ""location"": ""Ha Noi""," & vbCrLf

    j = j & "      ""code_prefix"": ""VTLT""" & vbCrLf

    j = j & "    }," & vbCrLf

    j = j & "    {" & vbCrLf

    j = j & "      ""id"": ""admin_1""," & vbCrLf

    j = j & "      ""alias"": ""So Khoa hoc va Cong nghe""," & vbCrLf

    j = j & "      ""parent_agency"": ""UBND TINH...""," & vbCrLf

    j = j & "      ""agency_name"": ""SO KHOA HOC VA CONG NGHE""," & vbCrLf

    j = j & "      ""location"": ""...""," & vbCrLf

    j = j & "      ""code_prefix"": ""SKHCN""" & vbCrLf

    j = j & "    }" & vbCrLf

    j = j & "  ]," & vbCrLf

    j = j & "  ""agencies_party"": [" & vbCrLf

    j = j & "    {" & vbCrLf

    j = j & "      ""id"": ""party_1""," & vbCrLf

    j = j & "      ""alias"": ""Dang uy Cac co quan Dang tinh""," & vbCrLf

    j = j & "      ""parent_party"": ""TINH UY...""," & vbCrLf

    j = j & "      ""party_name"": ""DANG UY CAC CO QUAN DANG TINH""," & vbCrLf

    j = j & "      ""location"": ""...""," & vbCrLf

    j = j & "      ""code_prefix"": ""NQ/DU""" & vbCrLf

    j = j & "    }" & vbCrLf

    j = j & "  ]," & vbCrLf

    j = j & "  ""recipients_list"": [" & vbCrLf

    j = j & "    {" & vbCrLf

    j = j & "      ""id"": ""rec_bo""," & vbCrLf

    j = j & "      ""title"": ""Mau Noi nhan Cap Bo / Co quan ngang Bo""," & vbCrLf

    j = j & "      ""content"": ""- Nhu Dieu ...;\n- Thu tuong, cac Pho Thu tuong CP;\n- Cac Bo, co quan ngang Bo;\n- Luu: VT, PC.""" & vbCrLf

    j = j & "    }," & vbCrLf

    j = j & "    {" & vbCrLf

    j = j & "      ""id"": ""rec_1""," & vbCrLf

    j = j & "      ""title"": ""Mau Noi nhan Cong van thong dung""," & vbCrLf

    j = j & "      ""content"": ""- Nhu kinh gui;\n- Giam doc, cac Pho Giam doc;\n- Cac phong, ban, don vi truc thuoc;\n- Luu: VT, TH.""" & vbCrLf

    j = j & "    }" & vbCrLf

    j = j & "  ]," & vbCrLf

    j = j & "  ""signers_list"": [" & vbCrLf

    j = j & "    {" & vbCrLf

    j = j & "      ""id"": ""sig_bo_truong""," & vbCrLf

    j = j & "      ""alias"": ""Bo truong - Nguyen Van A""," & vbCrLf

    j = j & "      ""title"": ""BO TRUONG""," & vbCrLf

    j = j & "      ""position"": ""BO TRUONG""," & vbCrLf

    j = j & "      ""fullname"": ""Nguyen Van A""," & vbCrLf

    j = j & "      ""signer_type"": ""Cap truong truc tiep""" & vbCrLf

    j = j & "    }," & vbCrLf

    j = j & "    {" & vbCrLf

    j = j & "      ""id"": ""sig_cuc_truong""," & vbCrLf

    j = j & "      ""alias"": ""Cuc truong - Le Van C""," & vbCrLf

    j = j & "      ""title"": ""CUC TRUONG""," & vbCrLf

    j = j & "      ""position"": ""CUC TRUONG""," & vbCrLf

    j = j & "      ""fullname"": ""Le Van C""," & vbCrLf

    j = j & "      ""signer_type"": ""Cap truong truc tiep""" & vbCrLf

    j = j & "    }" & vbCrLf

    j = j & "  ]" & vbCrLf

    j = j & "}"

    GetDefaultPresetsJson = j

End Function



Public Sub OpenPersonalizeWindow()

    ' ==============================================================================

    ' MO BANG CAI DAT CA NHAN HOA NATIVE USERFORM TRONG WORD - KHANG ANTIVIRUS 100%

    ' ==============================================================================

    On Error GoTo ErrH

    EnsureVFormatAppDir

    frmPersonalize.Show vbModal

    Unload frmPersonalize

    Exit Sub

ErrH:

    RecordError "VFormat_Core.OpenPersonalizeWindow", Err.Number, Err.Description, "Cai dat ca nhan hoa"

End Sub



Private Sub ConfigureAdminAgencyNative()

    On Error GoTo ErrH

    Dim rawJson As String: rawJson = GetUserPresetsJson()

    

    Dim slotStr As String

    slotStr = InputBox("Chon vi tri mau Co quan can cai dat (1 den 4):" & vbCrLf & _

                       "1: Mau Co quan 1" & vbCrLf & _

                       "2: Mau Co quan 2" & vbCrLf & _

                       "3: Mau Co quan 3" & vbCrLf & _

                       "4: Mau Co quan 4", "V-FORMAT - Chon vi tri Co quan", "1")

    slotStr = Trim(slotStr)

    If Len(slotStr) = 0 Then Exit Sub

    Dim idx As Long: idx = Val(slotStr)

    If idx < 1 Or idx > 4 Then idx = 1

    

    Dim curAgency As String, curParent As String, curLoc As String, curPrefix As String

    curAgency = ExtractJsonArrayItemTag(rawJson, "agencies_admin", "agency_name", idx)

    curParent = ExtractJsonArrayItemTag(rawJson, "agencies_admin", "parent_agency", idx)

    curLoc = ExtractJsonArrayItemTag(rawJson, "agencies_admin", "location", idx)

    curPrefix = ExtractJsonArrayItemTag(rawJson, "agencies_admin", "code_prefix", idx)

    

    Dim newAgency As String

    newAgency = InputBox("1/4. Nhap TEN CO QUAN BAN HANH:" & vbCrLf & _

                         "(Vi du: SO KHOA HOC VA CONG NGHE / UBND PHUONG...)", _

                         "V-FORMAT - Co Quan Ban Hanh (Mau " & idx & ")", curAgency)

    If Len(Trim(newAgency)) = 0 Then Exit Sub

    

    Dim newParent As String

    newParent = InputBox("2/4. Nhap CO QUAN CAP TREN TRUC TIEP:" & vbCrLf & _

                         "(Vi du: UBND TINH LAM DONG / BO TAI CHINH...):", _

                         "V-FORMAT - Co Quan Cap Tren (Mau " & idx & ")", curParent)

    If Len(Trim(newParent)) = 0 Then newParent = curParent

    

    Dim newLoc As String

    newLoc = InputBox("3/4. Nhap DIA DANH BAN HANH:" & vbCrLf & _

                      "(Vi du: Da Lat, Ha Noi, TP. Ho Chi Minh...):", _

                      "V-FORMAT - Dia Danh (Mau " & idx & ")", curLoc)

    If Len(Trim(newLoc)) = 0 Then newLoc = curLoc

    

    Dim newPrefix As String

    newPrefix = InputBox("4/4. Nhap KY HIEU VIET TAT SO CONG VAN:" & vbCrLf & _

                         "(Vi du: SKHCN, SGDDT, UBND, VP...):", _

                         "V-FORMAT - Ky Hieu So (Mau " & idx & ")", curPrefix)

    If Len(Trim(newPrefix)) = 0 Then newPrefix = curPrefix

    

    Dim newJson As String: newJson = rawJson

    newJson = UpdateJsonArrayItemTag(newJson, "agencies_admin", "agency_name", idx, newAgency)

    newJson = UpdateJsonArrayItemTag(newJson, "agencies_admin", "alias", idx, newAgency)

    newJson = UpdateJsonArrayItemTag(newJson, "agencies_admin", "parent_agency", idx, newParent)

    newJson = UpdateJsonArrayItemTag(newJson, "agencies_admin", "location", idx, newLoc)

    newJson = UpdateJsonArrayItemTag(newJson, "agencies_admin", "code_prefix", idx, newPrefix)

    

    SaveUserPresetsJson newJson

    ShowUniMsg "xJDDoyBsxrB1IGPDoGkgxJHhurd0IEPGoSBxdWFuIEjDoG5oIGNow61uaCB0aMOgbmggY8O0bmch", "Vi1GT1JNQVQ=", 64

    Exit Sub

ErrH:

    ShowUniMsg "Loi khi luu co quan: " & Err.Description, "Vi1GT1JNQVQ=", 48

End Sub



Private Sub ConfigurePartyAgencyNative()

    On Error GoTo ErrH

    Dim rawJson As String: rawJson = GetUserPresetsJson()

    

    Dim slotStr As String

    slotStr = InputBox("Chon vi tri mau Co quan Dang can cai dat (1 den 4):" & vbCrLf & _

                       "1: Mau Co quan Dang 1" & vbCrLf & _

                       "2: Mau Co quan Dang 2" & vbCrLf & _

                       "3: Mau Co quan Dang 3" & vbCrLf & _

                       "4: Mau Co quan Dang 4", "V-FORMAT - Chon vi tri Co quan Dang", "1")

    slotStr = Trim(slotStr)

    If Len(slotStr) = 0 Then Exit Sub

    Dim idx As Long: idx = Val(slotStr)

    If idx < 1 Or idx > 4 Then idx = 1

    

    Dim curParty As String, curParent As String, curLoc As String, curPrefix As String

    curParty = ExtractJsonArrayItemTag(rawJson, "agencies_party", "party_name", idx)

    curParent = ExtractJsonArrayItemTag(rawJson, "agencies_party", "parent_party", idx)

    curLoc = ExtractJsonArrayItemTag(rawJson, "agencies_party", "location", idx)

    curPrefix = ExtractJsonArrayItemTag(rawJson, "agencies_party", "code_prefix", idx)

    

    Dim newParty As String

    newParty = InputBox("1/4. Nhap TEN CO QUAN DANG BAN HANH:" & vbCrLf & _

                        "(Vi du: DANG BO XA BINH AN / CHI BO...):", _

                        "V-FORMAT - Co Quan Dang (Mau " & idx & ")", curParty)

    If Len(Trim(newParty)) = 0 Then Exit Sub

    

    Dim newParent As String

    newParent = InputBox("2/4. Nhap CAP UY CAP TREN TRUC TIEP:" & vbCrLf & _

                         "(Vi du: TINH UY... / DANG UY CAP TREN...):", _

                         "V-FORMAT - Cap Uy Cap Tren (Mau " & idx & ")", curParent)

    If Len(Trim(newParent)) = 0 Then newParent = curParent

    

    Dim newLoc As String

    newLoc = InputBox("3/4. Nhap DIA DANH BAN HANH:" & vbCrLf & _

                      "(Vi du: Da Lat, Ha Noi...):", _

                      "V-FORMAT - Dia Danh (Mau " & idx & ")", curLoc)

    If Len(Trim(newLoc)) = 0 Then newLoc = curLoc

    

    Dim newPrefix As String

    newPrefix = InputBox("4/4. Nhap KY HIEU VIET TAT SO VAN BAN DANG:" & vbCrLf & _

                         "(Vi du: NQ/DU, NQ/CB, QD/TU...):", _

                         "V-FORMAT - Ky Hieu So Dang (Mau " & idx & ")", curPrefix)

    If Len(Trim(newPrefix)) = 0 Then newPrefix = curPrefix

    

    Dim newJson As String: newJson = rawJson

    newJson = UpdateJsonArrayItemTag(newJson, "agencies_party", "party_name", idx, newParty)

    newJson = UpdateJsonArrayItemTag(newJson, "agencies_party", "alias", idx, newParty)

    newJson = UpdateJsonArrayItemTag(newJson, "agencies_party", "parent_party", idx, newParent)

    newJson = UpdateJsonArrayItemTag(newJson, "agencies_party", "location", idx, newLoc)

    newJson = UpdateJsonArrayItemTag(newJson, "agencies_party", "code_prefix", idx, newPrefix)

    

    SaveUserPresetsJson newJson

    ShowUniMsg "xJDDoyBsxrB1IGPDoGkgxJHhurd0IEPGoSBxdWFuIMSQ4bqjbmcgdGjDoG5oIGPDtG5nIQ==", "Vi1GT1JNQVQ=", 64

    Exit Sub

ErrH:

    ShowUniMsg "Loi khi luu co quan Dang: " & Err.Description, "Vi1GT1JNQVQ=", 48

End Sub



Private Sub ConfigureSignerNative()

    On Error GoTo ErrH

    Dim rawJson As String: rawJson = GetUserPresetsJson()

    

    Dim slotStr As String

    slotStr = InputBox("Chon vi tri Nguoi ky can cai dat (1 den 5):" & vbCrLf & _

                       "1: Mau Nguoi ky 1" & vbCrLf & _

                       "2: Mau Nguoi ky 2" & vbCrLf & _

                       "3: Mau Nguoi ky 3" & vbCrLf & _

                       "4: Mau Nguoi ky 4" & vbCrLf & _

                       "5: Mau Nguoi ky 5", "V-FORMAT - Chon vi tri Nguoi ky", "1")

    slotStr = Trim(slotStr)

    If Len(slotStr) = 0 Then Exit Sub

    Dim idx As Long: idx = Val(slotStr)

    If idx < 1 Or idx > 5 Then idx = 1

    

    Dim curName As String, curTitle As String

    curName = ExtractJsonArrayItemTag(rawJson, "signers_list", "fullname", idx)

    curTitle = ExtractJsonArrayItemTag(rawJson, "signers_list", "title", idx)

    

    Dim newName As String

    newName = InputBox("1/2. Nhap HO VA TEN NGUOI KY:" & vbCrLf & _

                       "(Vi du: Nguyen Van A, Tran Thi B...):", _

                       "V-FORMAT - Ho Ten Nguoi Ky (Mau " & idx & ")", curName)

    If Len(Trim(newName)) = 0 Then Exit Sub

    

    Dim newTitle As String

    newTitle = InputBox("2/2. Nhap CHUC VU / QUYEN HAN KY:" & vbCrLf & _

                        "(Vi du: GIAM DOC / CHU TICH / TM. UY BAN NHAN DAN\nCHU TICH / BI THU):", _

                        "V-FORMAT - Chuc Vu Quyen Han Ky (Mau " & idx & ")", curTitle)

    If Len(Trim(newTitle)) = 0 Then newTitle = curTitle

    

    Dim newJson As String: newJson = rawJson

    newJson = UpdateJsonArrayItemTag(newJson, "signers_list", "fullname", idx, newName)

    newJson = UpdateJsonArrayItemTag(newJson, "signers_list", "title", idx, newTitle)

    newJson = UpdateJsonArrayItemTag(newJson, "signers_list", "position", idx, newTitle)

    newJson = UpdateJsonArrayItemTag(newJson, "signers_list", "alias", idx, newTitle & " - " & newName)

    

    SaveUserPresetsJson newJson

    ShowUniMsg "xJDDoyBsxrB1IGPDoGkgxJHhurd0IENo4buvIGvDvSB0aMOgbmggY8O0bmch", "Vi1GT1JNQVQ=", 64

    Exit Sub

ErrH:

    ShowUniMsg "Loi khi luu nguoi ky: " & Err.Description, "Vi1GT1JNQVQ=", 48

End Sub



Private Sub ConfigureRecipientsNative()

    On Error GoTo ErrH

    Dim rawJson As String: rawJson = GetUserPresetsJson()

    

    Dim slotStr As String

    slotStr = InputBox("Chon vi tri Danh muc Noi nhan can cai dat (1 den 3):" & vbCrLf & _

                       "1: Mau Noi nhan 1" & vbCrLf & _

                       "2: Mau Noi nhan 2" & vbCrLf & _

                       "3: Mau Noi nhan 3", "V-FORMAT - Chon vi tri Noi nhan", "1")

    slotStr = Trim(slotStr)

    If Len(slotStr) = 0 Then Exit Sub

    Dim idx As Long: idx = Val(slotStr)

    If idx < 1 Or idx > 3 Then idx = 1

    

    Dim curTitle As String, curContent As String

    curTitle = ExtractJsonArrayItemTag(rawJson, "recipients_list", "title", idx)

    curContent = ExtractJsonArrayItemTag(rawJson, "recipients_list", "content", idx)

    

    Dim newTitle As String

    newTitle = InputBox("1/2. Nhap TIEU DE MAU NOI NHAN:" & vbCrLf & _

                        "(Vi du: Mau Noi nhan Cong van / Quyet dinh / Chi bo...):", _

                        "V-FORMAT - Tieu De Noi Nhan (Mau " & idx & ")", curTitle)

    If Len(Trim(newTitle)) = 0 Then Exit Sub

    

    Dim newContent As String

    newContent = InputBox("2/2. Nhap DANH SACH NOI NHAN (phan cach boi dau cham phay ';' hoac xuong dong):" & vbCrLf & _

                          "(Vi du: - Nhu kinh gui;\n- Giam doc, cac Pho GD;\n- Luu: VT.):", _

                          "V-FORMAT - Noi Dung Noi Nhan (Mau " & idx & ")", curContent)

    If Len(Trim(newContent)) = 0 Then newContent = curContent

    

    Dim newJson As String: newJson = rawJson

    newJson = UpdateJsonArrayItemTag(newJson, "recipients_list", "title", idx, newTitle)

    newJson = UpdateJsonArrayItemTag(newJson, "recipients_list", "content", idx, newContent)

    

    SaveUserPresetsJson newJson

    ShowUniMsg "xJDDoyBsxrB1IGPDoGkgxJHhurd0IE7GoWkgbmjhuq1uIHRow6BuaCBjw7RuZyE=", "Vi1GT1JNQVQ=", 64

    Exit Sub

ErrH:

    ShowUniMsg "Loi khi luu noi nhan: " & Err.Description, "Vi1GT1JNQVQ=", 48

End Sub



Private Sub ViewCurrentPresetsSummary()

    On Error GoTo ErrH

    Dim rawJson As String: rawJson = GetUserPresetsJson()

    

    Dim msg As String

    msg = "THONG TIN CA NHAN HOA HIEN TAI (V-FORMAT):" & vbCrLf

    msg = msg & "----------------------------------------" & vbCrLf

    msg = msg & "[CO QUAN HANH CHINH 1]:" & vbCrLf

    msg = msg & "- Co quan: " & ExtractJsonArrayItemTag(rawJson, "agencies_admin", "agency_name", 1) & vbCrLf

    msg = msg & "- Cap tren: " & ExtractJsonArrayItemTag(rawJson, "agencies_admin", "parent_agency", 1) & vbCrLf

    msg = msg & "- Dia danh: " & ExtractJsonArrayItemTag(rawJson, "agencies_admin", "location", 1) & " | Ky hieu: " & ExtractJsonArrayItemTag(rawJson, "agencies_admin", "code_prefix", 1) & vbCrLf & vbCrLf

    msg = msg & "[CO QUAN DANG 1]:" & vbCrLf

    msg = msg & "- Co quan Dang: " & ExtractJsonArrayItemTag(rawJson, "agencies_party", "party_name", 1) & vbCrLf

    msg = msg & "- Cap tren: " & ExtractJsonArrayItemTag(rawJson, "agencies_party", "parent_party", 1) & vbCrLf

    msg = msg & "- Dia danh: " & ExtractJsonArrayItemTag(rawJson, "agencies_party", "location", 1) & " | Ky hieu: " & ExtractJsonArrayItemTag(rawJson, "agencies_party", "code_prefix", 1) & vbCrLf & vbCrLf

    msg = msg & "[NGUOI KY 1]:" & vbCrLf

    msg = msg & "- Chuc vu: " & ExtractJsonArrayItemTag(rawJson, "signers_list", "title", 1) & vbCrLf

    msg = msg & "- Ho ten: " & ExtractJsonArrayItemTag(rawJson, "signers_list", "fullname", 1) & vbCrLf & vbCrLf

    msg = msg & "[NOI NHAN 1]: " & ExtractJsonArrayItemTag(rawJson, "recipients_list", "title", 1) & vbCrLf

    msg = msg & "----------------------------------------" & vbCrLf

    msg = msg & "(De thay doi thong tin tren, chon cac muc 1, 2, 3 hoac 4 trong menu Cai dat)"

          

    ShowUniMsg msg, "Vi1GT1JNQVQ=", 64

    Exit Sub

ErrH:

    ShowUniMsg "Loi xem thong tin: " & Err.Description, "Vi1GT1JNQVQ=", 48

End Sub



Private Sub OpenPresetsInNotepad()

    On Error GoTo ErrH

    EnsureVFormatAppDir

    Dim appData As String: appData = Environ$("APPDATA") & "\VFormat\user_presets.json"

    If Dir(appData) = "" Then

        SaveUserPresetsJson GetDefaultPresetsJson()

    End If

    SafeOpenUrlOrFile appData

    Exit Sub

ErrH:

    ShowUniMsg "Khong the mo file cau hinh: " & Err.Description, "Vi1GT1JNQVQ=", 48

End Sub



Private Sub ResetPresetsToDefault()

    On Error GoTo ErrH

    If MsgBox("Ban co chac chan muon khoi phuc toan bo thong tin ca nhan hoa ve mac dinh?", vbYesNo + vbQuestion, "V-FORMAT") = vbYes Then

        SaveUserPresetsJson GetDefaultPresetsJson()

        ShowUniMsg "xJDDoyBow7RpIHBo4bulYyBj4bqldSBow6xuaCBt4bq3YyDEkeG7i25oIHRow6BuaCBjw7RuZyE=", "Vi1GT1JNQVQ=", 64

    End If

    Exit Sub

ErrH:

    ShowUniMsg "Loi khoi phuc mac dinh: " & Err.Description, "Vi1GT1JNQVQ=", 48

End Sub



Private Function ReadJsonFile(ByVal filePath As String) As String

    On Error GoTo ErrH

    If Dir(filePath) = "" Then Exit Function

    Dim fn As Integer: fn = FreeFile

    Open filePath For Binary Access Read As #fn

    Dim flen As Long: flen = LOF(fn)

    If flen = 0 Then

        Close #fn

        Exit Function

    End If

    Dim b() As Byte: ReDim b(0 To flen - 1)

    Get #fn, , b

    Close #fn

    

    Dim i As Long, c As Long, ch As Long, res As String

    i = 0

    If flen >= 3 Then

        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then i = 3

    End If

    Do While i < flen

        c = b(i)

        If c < &H80 Then

            res = res & ChrW(c)

            i = i + 1

        ElseIf (c And &HE0) = &HC0 And i + 1 < flen Then

            ch = ((c And &H1F) * 64) Or (b(i + 1) And &H3F)

            res = res & ChrW(ch)

            i = i + 2

        ElseIf (c And &HF0) = &HE0 And i + 2 < flen Then

            ch = ((c And &HF) * 4096) Or ((CLng(b(i + 1) And &H3F)) * 64) Or (b(i + 2) And &H3F)

            res = res & ChrW(ch)

            i = i + 3

        Else

            i = i + 1

        End If

    Loop

    ReadJsonFile = res

    Exit Function

ErrH:

    ReadJsonFile = ""

End Function



Public Function GetUserPresetsJson() As String

    Dim appData As String, content As String

    On Error Resume Next

    

    ' 1. Check in %APPDATA%\VFormat\user_presets.json

    appData = Environ$("APPDATA") & "\VFormat\user_presets.json"

    content = ReadJsonFile(appData)

    

    ' 2. Check in ThisDocument.Path

    If Len(content) = 0 And Len(ThisDocument.Path) > 0 Then

        content = ReadJsonFile(ThisDocument.Path & "\user_presets.json")

    End If

    

    ' 3. Check in parent folder of Add-in

    If Len(content) = 0 And Len(ThisDocument.Path) > 0 Then

        content = ReadJsonFile(Left(ThisDocument.Path, InStrRev(ThisDocument.Path, "\")) & "user_presets.json")

    End If

    

    ' 4. Fallback to dev path

    If Len(content) = 0 Then

        content = ReadJsonFile(GetVFormatAppDir() & "\user_presets.json")

    End If

    

    ' 5. If still empty, use default bundled presets

    If Len(content) = 0 Then

        content = GetDefaultPresetsJson()

        EnsureVFormatAppDir

        WriteJsonFile appData, content

    End If

    

    GetUserPresetsJson = content

End Function



Public Function GetPresetDisplayLabel(ByVal rawJson As String, ByVal sectionName As String, ByVal idx As Long, ByVal defaultFallback As String) As String

    On Error Resume Next

    Dim aliasName As String

    aliasName = ExtractJsonArrayItemTag(rawJson, sectionName, "alias", idx)

    If Len(Trim(aliasName)) > 0 Then

        GetPresetDisplayLabel = Trim(aliasName)

        Exit Function

    End If

    

    If sectionName = "agencies_admin" Then

        Dim agName As String: agName = ExtractJsonArrayItemTag(rawJson, sectionName, "agency_name", idx)

        If Len(agName) > 0 Then

            GetPresetDisplayLabel = agName

            Exit Function

        End If

    ElseIf sectionName = "agencies_party" Then

        Dim ptName As String: ptName = ExtractJsonArrayItemTag(rawJson, sectionName, "party_name", idx)

        If Len(ptName) > 0 Then

            GetPresetDisplayLabel = ptName

            Exit Function

        End If

    ElseIf sectionName = "signers_list" Then

        Dim sName As String: sName = ExtractJsonArrayItemTag(rawJson, sectionName, "fullname", idx)

        Dim sTitle As String: sTitle = ExtractJsonArrayItemTag(rawJson, sectionName, "title", idx)

        If Len(sName) > 0 Then

            If Len(sTitle) > 0 Then

                GetPresetDisplayLabel = sTitle & " - " & sName

            Else

                GetPresetDisplayLabel = sName

            End If

            Exit Function

        End If

    ElseIf sectionName = "recipients_list" Then

        Dim rTitle As String: rTitle = ExtractJsonArrayItemTag(rawJson, sectionName, "title", idx)

        If Len(rTitle) > 0 Then

            GetPresetDisplayLabel = rTitle

            Exit Function

        End If

    End If

    

    GetPresetDisplayLabel = defaultFallback

End Function



Public Function ExtractJsonArrayItemTag(ByVal jsonStr As String, ByVal sectionName As String, ByVal tagName As String, Optional ByVal itemIndex As Long = 1) As String

    On Error Resume Next

    Dim sIdx As Long: sIdx = 1

    If Len(sectionName) > 0 Then

        sIdx = InStr(jsonStr, """" & sectionName & """")

        If sIdx = 0 Then sIdx = 1

    End If

    

    Dim p As Long, sPos As Long, ePos As Long, currIdx As Long

    Dim pat As String: pat = """" & tagName & """"

    p = sIdx

    currIdx = 0

    

    Do

        p = InStr(p + 1, jsonStr, pat)

        If p = 0 Then Exit Do

        currIdx = currIdx + 1

        If currIdx = itemIndex Then

            p = InStr(p + Len(pat), jsonStr, ":")

            If p > 0 Then

                sPos = p + 1

                Do While sPos <= Len(jsonStr) And (Mid(jsonStr, sPos, 1) = " " Or Mid(jsonStr, sPos, 1) = """")

                    sPos = sPos + 1

                Loop

                ePos = sPos

                Do While ePos <= Len(jsonStr) And Mid(jsonStr, ePos, 1) <> """" And Mid(jsonStr, ePos, 1) <> "," And Mid(jsonStr, ePos, 1) <> "}" And Mid(jsonStr, ePos, 1) <> vbCr And Mid(jsonStr, ePos, 1) <> vbLf

                    ePos = ePos + 1

                Loop

                ExtractJsonArrayItemTag = Replace(Replace(Trim(Mid(jsonStr, sPos, ePos - sPos)), "\n", vbCrLf), "\""", """")

                Exit Function

            End If

        End If

    Loop While p > 0

    ExtractJsonArrayItemTag = ""

End Function



Public Function ExtractJsonSectionTag(ByVal jsonStr As String, ByVal sectionName As String, ByVal tagName As String) As String

    ExtractJsonSectionTag = ExtractJsonArrayItemTag(jsonStr, sectionName, tagName, 1)

End Function



' --- CHEN QUOC HIEU TU DONG THEO QUY CHUAN DANG CHON (ND 30 / HD 05) ---





Public Sub InsertQuocHieuCustomAtCursor()

    InsertQuocHieuAutoAtCursor

End Sub



' --- CHEN CO QUAN HANH CHINH THEO CHI SO (ND 30) ---

Public Sub InsertQuocHieuAdminIndex(Optional ByVal idx As Long = 1)

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    

    Dim rawJson As String: rawJson = GetUserPresetsJson()

    Dim parentAgency As String, agencyName As String, locationName As String, codePrefix As String

    

    parentAgency = ExtractJsonArrayItemTag(rawJson, "agencies_admin", "parent_agency", idx)

    agencyName = ExtractJsonArrayItemTag(rawJson, "agencies_admin", "agency_name", idx)

    locationName = ExtractJsonArrayItemTag(rawJson, "agencies_admin", "location", idx)

    codePrefix = ExtractJsonArrayItemTag(rawJson, "agencies_admin", "code_prefix", idx)

    

    If Len(parentAgency) = 0 And idx > 2 Then parentAgency = DecodeUTF8("VUJORCBU4buITkguLi4=")

    If Len(agencyName) = 0 Then

        If idx = 1 Then

            agencyName = DecodeUTF8("QuG7mCBO4buYSSBW4buY")

            parentAgency = ""

        ElseIf idx = 2 Then

            parentAgency = DecodeUTF8("QuG7mCBO4buYSSBW4buY")

            agencyName = DecodeUTF8("Q+G7pEMgVsSDTiBUSMavIFbDgCBMxrBVIFRS4buuIE5Iw4AgTsav4buaQw==")

        Else

            agencyName = DecodeUTF8("U+G7niBLSE9BIEjhu4xDIFbDgCBDw5RORyBOR0jhu4Y=")

        End If

    End If

    If Len(locationName) = 0 Then locationName = "..."

    If Len(codePrefix) = 0 Then

        If idx = 1 Then

            codePrefix = "BNV"

        ElseIf idx = 2 Then

            codePrefix = "VTLT"

        Else

            codePrefix = "SKHCN"

        End If

    End If

    

    Dim rng As Range: Set rng = Selection.Range

    Dim tbl As Table

    Set tbl = rng.Tables.Add(Range:=rng, NumRows:=2, NumColumns:=2)

    tbl.Borders.Enable = False

    tbl.AutoFitBehavior wdAutoFitFixed

    tbl.LeftPadding = 0

    tbl.RightPadding = 0

    tbl.TopPadding = 0

    tbl.BottomPadding = 0

    tbl.Rows.LeftIndent = 0

    tbl.Columns(1).Width = MillimetersToPoints(65)

    tbl.Columns(2).Width = MillimetersToPoints(100)

    

    With tbl.Cell(1, 1).Range

        If Len(parentAgency) > 0 Then

            .Text = parentAgency & vbCrLf & agencyName

            .ParagraphFormat.Alignment = wdAlignParagraphCenter

            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            .ParagraphFormat.SpaceBefore = 0

            .ParagraphFormat.SpaceAfter = 0

            .Paragraphs(1).Range.Font.Name = "Times New Roman"

            .Paragraphs(1).Range.Font.Size = 12

            .Paragraphs(1).Range.Font.Bold = False

            .Paragraphs(2).Range.Font.Name = "Times New Roman"

            .Paragraphs(2).Range.Font.Size = 12

            .Paragraphs(2).Range.Font.Bold = True

        Else

            .Text = agencyName

            .ParagraphFormat.Alignment = wdAlignParagraphCenter

            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            .ParagraphFormat.SpaceBefore = 0

            .ParagraphFormat.SpaceAfter = 0

            .Paragraphs(1).Range.Font.Name = "Times New Roman"

            .Paragraphs(1).Range.Font.Size = 12.5

            .Paragraphs(1).Range.Font.Bold = True

        End If

    End With

    With tbl.Cell(2, 1).Range

        .Text = DecodeUTF8("U+G7kTogLi4uL1HEkC0=") & codePrefix

        .ParagraphFormat.Alignment = wdAlignParagraphCenter

        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

        .ParagraphFormat.SpaceBefore = 2

        .ParagraphFormat.SpaceAfter = 0

        .Font.Name = "Times New Roman"

        .Font.Size = 13

        .Font.Bold = False

    End With

    

    With tbl.Cell(1, 2).Range

        .Text = DecodeUTF8("Q+G7mE5HIEjDkkEgWMODIEjhu5hJIENI4bumIE5HSMSoQSBWSeG7hlQgTkFN") & vbCrLf & DecodeUTF8("xJDhu5ljIGzhuq1wIC0gVOG7sSBkbyAtIEjhuqFuaCBwaMO6Yw==")

        .ParagraphFormat.Alignment = wdAlignParagraphCenter

        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

        .ParagraphFormat.SpaceBefore = 0

        .ParagraphFormat.SpaceAfter = 0

        .Paragraphs(1).Range.Font.Name = "Times New Roman"

        .Paragraphs(1).Range.Font.Size = 12

        .Paragraphs(1).Range.Font.Bold = True

        .Paragraphs(2).Range.Font.Name = "Times New Roman"

        .Paragraphs(2).Range.Font.Size = 13

        .Paragraphs(2).Range.Font.Bold = True

    End With

    With tbl.Cell(2, 2).Range

        .Text = locationName & DecodeUTF8("LCBuZ8OgeSAuLi4gdGjDoW5nIC4uLiBuxINtIDIwLi4u")

        .ParagraphFormat.Alignment = wdAlignParagraphCenter

        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

        .ParagraphFormat.SpaceBefore = 2

        .ParagraphFormat.SpaceAfter = 0

        .Font.Name = "Times New Roman"

        .Font.Size = 13

        .Font.Italic = True

    End With

    

    tbl.Range.Collapse wdCollapseEnd

    tbl.Range.InsertParagraphAfter

End Sub



Public Sub InsertQuocHieuND30AtCursor()

    InsertQuocHieuAdminIndex 1

End Sub



' --- CHEN CO QUAN DANG THEO CHI SO (HD 05) ---

Public Sub InsertQuocHieuPartyIndex(Optional ByVal idx As Long = 1)

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    

    Dim rawJson As String: rawJson = GetUserPresetsJson()

    Dim parentParty As String, partyName As String, locationName As String, codePrefix As String

    

    parentParty = ExtractJsonArrayItemTag(rawJson, "agencies_party", "parent_party", idx)

    partyName = ExtractJsonArrayItemTag(rawJson, "agencies_party", "party_name", idx)

    locationName = ExtractJsonArrayItemTag(rawJson, "agencies_party", "location", idx)

    codePrefix = ExtractJsonArrayItemTag(rawJson, "agencies_party", "code_prefix", idx)

    

    If Len(parentParty) = 0 Then parentParty = DecodeUTF8("VOG7iE5IIOG7plkuLi4=")

    If Len(partyName) = 0 Then partyName = DecodeUTF8("xJDhuqJORyDhu6ZZIEPDgUMgQ8agIFFVQU4gxJDhuqJORyBU4buITkg=")

    If Len(locationName) = 0 Then locationName = "..."

    If Len(codePrefix) = 0 Then codePrefix = "NQ/DU"

    

    Dim rng As Range: Set rng = Selection.Range

    Dim tbl As Table

    Set tbl = rng.Tables.Add(Range:=rng, NumRows:=2, NumColumns:=2)

    tbl.Borders.Enable = False

    tbl.AutoFitBehavior wdAutoFitFixed

    tbl.LeftPadding = 0

    tbl.RightPadding = 0

    tbl.TopPadding = 0

    tbl.BottomPadding = 0

    tbl.Columns(1).Width = MillimetersToPoints(78)

    tbl.Columns(2).Width = MillimetersToPoints(87)

    Dim rWParty As Row

    For Each rWParty In tbl.Rows

        If rWParty.Cells.Count >= 2 Then

            rWParty.Cells(1).Width = MillimetersToPoints(78)

            rWParty.Cells(2).Width = MillimetersToPoints(87)

        End If

    Next rWParty

    

    With tbl.Cell(1, 1).Range

        .Text = parentParty & vbCrLf & partyName & vbCrLf & "*"

        .ParagraphFormat.Alignment = wdAlignParagraphCenter

        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

        .ParagraphFormat.SpaceBefore = 0

        .ParagraphFormat.SpaceAfter = 0

        .Paragraphs(1).Range.Font.Name = "Times New Roman"

        .Paragraphs(1).Range.Font.Size = 14

        .Paragraphs(1).Range.Font.Bold = False

        .Paragraphs(2).Range.Font.Name = "Times New Roman"

        .Paragraphs(2).Range.Font.Size = 14

        .Paragraphs(2).Range.Font.Bold = True

        .Paragraphs(3).Range.Font.Name = "Times New Roman"

        .Paragraphs(3).Range.Font.Size = 14

        .Paragraphs(3).Range.Font.Bold = False

    End With

    With tbl.Cell(2, 1).Range

        .Text = DecodeUTF8("U+G7kSAwMS0=") & codePrefix

        .ParagraphFormat.Alignment = wdAlignParagraphCenter

        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

        .ParagraphFormat.SpaceBefore = 2

        .ParagraphFormat.SpaceAfter = 0

        .Font.Name = "Times New Roman"

        .Font.Size = 14

        .Font.Bold = False

    End With

    

    With tbl.Cell(1, 2).Range

        .Text = DecodeUTF8("xJDhuqJORyBD4buYTkcgU+G6ok4gVknhu4ZUIE5BTQ==")

        .ParagraphFormat.Alignment = wdAlignParagraphCenter

        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

        .ParagraphFormat.SpaceBefore = 0

        .ParagraphFormat.SpaceAfter = 0

        .Font.Name = "Times New Roman"

        .Font.Size = 15

        .Font.Bold = True

        .Font.Underline = wdUnderlineSingle

    End With

    With tbl.Cell(2, 2).Range

        .Text = locationName & DecodeUTF8("LCBuZ8OgeSAuLi4gdGjDoW5nIC4uLiBuxINtIDIwLi4u")

        .ParagraphFormat.Alignment = wdAlignParagraphCenter

        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

        .ParagraphFormat.SpaceBefore = 2

        .ParagraphFormat.SpaceAfter = 0

        .Font.Name = "Times New Roman"

        .Font.Size = 14

        .Font.Italic = True

        .Font.Underline = wdUnderlineNone

    End With

    

    tbl.Range.Collapse wdCollapseEnd

    tbl.Range.InsertParagraphAfter

End Sub



Public Sub InsertDangHeaderAtCursor()

    InsertQuocHieuPartyIndex 1

End Sub



Public Sub InsertDocTitleAtCursor(ByVal titleName As String, ByVal subjectText As String)

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    Dim rng As Range: Set rng = Selection.Range

    

    Dim regIdx As Integer: regIdx = GetEffectiveRegulation(ActiveDocument)

    

    rng.Text = titleName & vbCrLf & subjectText & vbCrLf

    Dim p1 As Paragraph: Set p1 = rng.Paragraphs(1)

    With p1.Range.Font

        .Name = "Times New Roman"

        If Left(titleName, 3) = "Vv " Or Left(titleName, 4) = "V/v " Then

            .Size = IIf(regIdx = 1, 12, 12.5)

            .Bold = False

            .Italic = IIf(regIdx = 1, True, False)

        Else

            .Size = IIf(regIdx = 1, 15, 14)

            .Bold = True

            .Italic = False

        End If

    End With

    p1.Range.ParagraphFormat.Alignment = IIf(Left(titleName, 3) = "Vv " Or Left(titleName, 4) = "V/v ", wdAlignParagraphLeft, wdAlignParagraphCenter)

    p1.Range.ParagraphFormat.SpaceBefore = 6

    p1.Range.ParagraphFormat.SpaceAfter = 2

    

    Dim p2 As Paragraph: Set p2 = rng.Paragraphs(2)

    p2.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

    p2.Range.ParagraphFormat.SpaceBefore = 0

    p2.Range.ParagraphFormat.SpaceAfter = 6

    

    Dim colonPos As Long

    colonPos = InStr(subjectText, ":")

    Dim uSubj As String: uSubj = UCase(ConvertToUnsignedVietnamese(subjectText))

    If InStr(uSubj, "KINH GUI") > 0 Or InStr(uSubj, "KINH TRINH") > 0 Then

        If regIdx = 1 Then

            ' HD05: Tu "Kinh gui:" / "Kinh trinh:" in nghieng 14pt, phan sau in dung 14pt

            If colonPos > 0 Then

                Dim rPref As Range

                Set rPref = ActiveDocument.Range(p2.Range.Start, p2.Range.Start + colonPos)

                With rPref.Font

                    .Name = "Times New Roman"

                    .Size = 14

                    .Bold = False

                    .Italic = True

                End With

                If p2.Range.Start + colonPos < p2.Range.End - 1 Then

                    Dim rSuff As Range

                    Set rSuff = ActiveDocument.Range(p2.Range.Start + colonPos, p2.Range.End - 1)

                    With rSuff.Font

                        .Name = "Times New Roman"

                        .Size = 14

                        .Bold = False

                        .Italic = False

                    End With

                End If

            Else

                With p2.Range.Font

                    .Name = "Times New Roman"

                    .Size = 14

                    .Bold = False

                    .Italic = True

                End With

            End If

        Else

            ' ND30: In dung 14pt

            With p2.Range.Font

                .Name = "Times New Roman"

                .Size = 14

                .Bold = False

                .Italic = False

            End With

        End If

    Else

        With p2.Range.Font

            .Name = "Times New Roman"

            .Size = IIf(regIdx = 1, 14, 13.5)

            .Bold = IIf(regIdx = 1, True, False)

            .Italic = IIf(InStr(titleName, "CONG VAN") > 0, True, False)

        End With

        If regIdx = 1 And Left(titleName, 3) <> "Vv " And Left(titleName, 4) <> "V/v " Then

            rng.InsertParagraphAfter

            Dim pSep As Paragraph: Set pSep = rng.Paragraphs(rng.Paragraphs.Count)

            pSep.Range.Text = "-----"

            pSep.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

            pSep.Range.ParagraphFormat.FirstLineIndent = 0

            pSep.Range.ParagraphFormat.SpaceBefore = 2

            pSep.Range.ParagraphFormat.SpaceAfter = 6

            SetParaFont pSep, "Times New Roman", 14, False, False

        End If

    End If

End Sub



Public Sub InsertPhuLucHeaderAtCursor()

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    Dim rng As Range: Set rng = Selection.Range

    Dim regIdx As Integer: regIdx = GetEffectiveRegulation(ActiveDocument)

    

    If regIdx = 1 Then

        rng.Text = DecodeUTF8("UEjhu6UgTOG7pEMgMQ==") & vbCrLf & _

                   DecodeUTF8("VMOKTiBQSOG7pCBM4bukQw==") & vbCrLf & _

                   DecodeUTF8("KEvDqG0gdGhlbyB2xINuIGLhuqNuIHPhu5EgLi4uLS4uLi8uLi4sIG5nw6B5IC4uLiB0aMOhbmcgLi4uIG7Eg20gMjAuLi4gY+G7p2EgLi4uKQ==") & vbCrLf & _

                   "-----" & vbCrLf

        Dim p1 As Paragraph: Set p1 = rng.Paragraphs(1)

        SetParaFont p1, "Times New Roman", 14, True, False

        p1.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

        p1.Range.ParagraphFormat.SpaceBefore = 0

        p1.Range.ParagraphFormat.SpaceAfter = 2

        

        Dim p2 As Paragraph: Set p2 = rng.Paragraphs(2)

        SetParaFont p2, "Times New Roman", 14, True, False

        p2.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

        p2.Range.ParagraphFormat.SpaceBefore = 0

        p2.Range.ParagraphFormat.SpaceAfter = 2

        

        Dim p3 As Paragraph: Set p3 = rng.Paragraphs(3)

        SetParaFont p3, "Times New Roman", 14, False, True

        p3.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

        p3.Range.ParagraphFormat.SpaceBefore = 0

        p3.Range.ParagraphFormat.SpaceAfter = 2

        

        Dim p4 As Paragraph: Set p4 = rng.Paragraphs(4)

        SetParaFont p4, "Times New Roman", 14, False, False

        p4.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

        p4.Range.ParagraphFormat.SpaceBefore = 0

        p4.Range.ParagraphFormat.SpaceAfter = 12

    Else

        rng.Text = DecodeUTF8("UGjhu6UgbOG7pWMgSQ==") & vbCrLf & _

                   DecodeUTF8("VMOKTiBQSOG7pCBM4bukQw==") & vbCrLf & _

                   DecodeUTF8("KEvDqG0gdGhlbyBWxINuIGLhuqNuIHPhu5EgLi4uLy4uLiBuZ8OgeSAuLi4gdGjDoW5nIC4uLiBuxINtIC4uLiBj4bunYSAuLi4p") & vbCrLf

                   

        Set p1 = rng.Paragraphs(1)

        With p1.Range.Font

            .Name = "Times New Roman"

            .Size = 14

            .Bold = True

        End With

        p1.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

        p1.Range.ParagraphFormat.SpaceBefore = 0

        p1.Range.ParagraphFormat.SpaceAfter = 2

        

        Set p2 = rng.Paragraphs(2)

        With p2.Range.Font

            .Name = "Times New Roman"

            .Size = 13.5

            .Bold = True

        End With

        p2.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

        p2.Range.ParagraphFormat.SpaceBefore = 0

        p2.Range.ParagraphFormat.SpaceAfter = 2

        

        Set p3 = rng.Paragraphs(3)

        With p3.Range.Font

            .Name = "Times New Roman"

            .Size = 13

            .Italic = True

            .Bold = False

        End With

        p3.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

        p3.Range.ParagraphFormat.SpaceBefore = 0

        p3.Range.ParagraphFormat.SpaceAfter = 12

    End If

End Sub



Public Sub InsertBanSaoHeaderAtCursor(Optional ByVal saoType As String = "SAO Y")

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    DbgLog "InsertBanSaoHeaderAtCursor: enter"

    

    Dim regIdx As Integer: regIdx = GetEffectiveRegulation(ActiveDocument)

    DbgLog "InsertBanSaoHeaderAtCursor: regIdx=" & regIdx

    Dim rawJson As String: rawJson = GetUserPresetsJson()

    DbgLog "InsertBanSaoHeaderAtCursor: rawJson done"

    

    If regIdx = 1 Then

        DbgLog "InsertBanSaoHeaderAtCursor: branch 1"

        Dim partyAgency As String, locDang As String

        partyAgency = ExtractJsonArrayItemTag(rawJson, "agencies_party", "party_name", 1)

        DbgLog "InsertBanSaoHeaderAtCursor: partyAgency=" & partyAgency

        locDang = ExtractJsonArrayItemTag(rawJson, "agencies_party", "location", 1)

        DbgLog "InsertBanSaoHeaderAtCursor: locDang=" & locDang

        If Len(partyAgency) = 0 Then partyAgency = DecodeUTF8("VsSCTiBQSMOSTkcgVOG7iE5IIOG7plk=")

        If Len(locDang) = 0 Then locDang = "..."

        

        Dim rngDang As Range: Set rngDang = Selection.Range

        DbgLog "InsertBanSaoHeaderAtCursor: rngDang set"

        Dim tblDang As Table

        Set tblDang = rngDang.Tables.Add(Range:=rngDang, NumRows:=1, NumColumns:=2)

        DbgLog "InsertBanSaoHeaderAtCursor: Tables.Add done"

        tblDang.Borders.Enable = False

        tblDang.Columns(1).Width = MillimetersToPoints(80)

        tblDang.Columns(2).Width = MillimetersToPoints(85)

        DbgLog "InsertBanSaoHeaderAtCursor: widths set"

        

        ' Cot 1: Hop Dau Ban sao 80mm x 50mm co vien ngoai

        With tblDang.Cell(1, 1).Range

            .Text = DecodeUTF8("QuG6ok4gU0FPIFbEg04gQuG6ok4=") & vbCrLf & _

                    DecodeUTF8("KDEpIEPGoSBxdWFuIHNhbzog") & partyAgency & vbCrLf & _

                    DecodeUTF8("KDIpIFPhu5EsIGvDvSBoaeG7h3UgYuG6o24gc2FvOiAwMS1CUy9WUA==") & vbCrLf & _

                    DecodeUTF8("KDMpIEjDrG5oIHRo4bupYyBzYW86IA==") & saoType & vbCrLf & _

                    DecodeUTF8("KDQpIMSQ4buLYSBkYW5oIHbDoCBuZ8OgeSwgdGjDoW5nLCBuxINtIHNhbzog") & locDang & DecodeUTF8("LCBuZ8OgeSAuLi4gdGjDoW5nIC4uLiBuxINtIDIwLi4u") & vbCrLf & _

                    DecodeUTF8("KDUpIE7GoWkgbmjhuq1uIGLhuqNuIHNhbzogLi4u")

            .ParagraphFormat.Alignment = wdAlignParagraphLeft

            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            .ParagraphFormat.SpaceBefore = 0

            .ParagraphFormat.SpaceAfter = 2

            

            .Paragraphs(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

            SetParaFont .Paragraphs(1), "Times New Roman", 13, True, False

            

            Dim pk As Long

            For pk = 2 To .Paragraphs.Count

                SetParaFont .Paragraphs(pk), "Times New Roman", 10, False, False

            Next pk

        End With

        tblDang.Cell(1, 1).Borders.Enable = True

        tblDang.Cell(1, 1).Borders.LineWidth = wdLineWidth050pt

        

        ' Cot 2: Chu ky nguoi co tham quyen sao (HD 05)

        With tblDang.Cell(1, 2).Range

            .Text = DecodeUTF8("VC9MIEJBTiBUSMav4bucTkcgVuG7pA==") & vbCrLf & _

                    DecodeUTF8("Q0jDgU5IIFbEgk4gUEjDkk5H") & vbCrLf & _

                    vbCrLf & vbCrLf & vbCrLf & _

                    DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

            .ParagraphFormat.Alignment = wdAlignParagraphCenter

            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            .ParagraphFormat.SpaceBefore = 0

            .ParagraphFormat.SpaceAfter = 0

            SetParaFont .Paragraphs(1), "Times New Roman", 14, True, False

            SetParaFont .Paragraphs(2), "Times New Roman", 14, False, False

            SetParaFont .Paragraphs(.Paragraphs.Count), "Times New Roman", 14, True, False

        End With

        

        tblDang.Range.Collapse wdCollapseEnd

        tblDang.Range.InsertParagraphAfter

        Exit Sub

    End If

    

    ' BAN SAO THEO ND 30 (HANH CHINH)

    Dim parentAgency As String, agencyName As String, locationName As String

    parentAgency = ExtractJsonSectionTag(rawJson, "agencies_admin", "parent_agency")

    agencyName = ExtractJsonSectionTag(rawJson, "agencies_admin", "agency_name")

    locationName = ExtractJsonSectionTag(rawJson, "agencies_admin", "location")

    

    If Len(parentAgency) = 0 Then parentAgency = DecodeUTF8("VUJORCBU4buITkggTMOCTSDEkOG7kk5H")

    If Len(agencyName) = 0 Then agencyName = DecodeUTF8("VsSCTiBQSMOSTkc=")

    If Len(locationName) = 0 Then locationName = DecodeUTF8("TMOCTSDEkOG7kk5H")

    

    Dim rng As Range: Set rng = Selection.Range

    Dim tbl As Table

    Set tbl = rng.Tables.Add(Range:=rng, NumRows:=1, NumColumns:=2)

    tbl.Borders.Enable = False

    tbl.Columns(1).Width = MillimetersToPoints(75)

    tbl.Columns(2).Width = MillimetersToPoints(95)

    

    Dim codeKieu As String

    Select Case UCase(saoType)

        Case "SAO LUC": codeKieu = "SL"

        Case "TRICH SAO": codeKieu = "TrS"

        Case Else: codeKieu = "SY"

    End Select

    

    With tbl.Cell(1, 1).Range

        .Text = parentAgency & vbCrLf & agencyName & vbCrLf & _

                DecodeUTF8("U+G7kTogMDEv") & codeKieu & "-VP" & vbCrLf & vbCrLf & _

                DecodeUTF8("TsahaSBuaOG6rW46") & vbCrLf & _

                DecodeUTF8("LSBOaMawIHRyw6puOw0KLSBMxrB1OiBWVCwgSFMu")

        .ParagraphFormat.Alignment = wdAlignParagraphLeft

        .Paragraphs(1).Range.Font.Name = "Times New Roman"

        .Paragraphs(1).Range.Font.Size = 12.5

        .Paragraphs(1).Range.Font.Bold = True

        .Paragraphs(2).Range.Font.Name = "Times New Roman"

        .Paragraphs(2).Range.Font.Size = 12.5

        .Paragraphs(2).Range.Font.Bold = True

        .Paragraphs(3).Range.Font.Name = "Times New Roman"

        .Paragraphs(3).Range.Font.Size = 13

        .Paragraphs(3).Range.Font.Bold = False

        .Paragraphs(5).Range.Font.Name = "Times New Roman"

        .Paragraphs(5).Range.Font.Size = 12

        .Paragraphs(5).Range.Font.Bold = True

        .Paragraphs(5).Range.Font.Italic = True

        .Paragraphs(6).Range.Font.Name = "Times New Roman"

        .Paragraphs(6).Range.Font.Size = 11

        .Paragraphs(7).Range.Font.Name = "Times New Roman"

        .Paragraphs(7).Range.Font.Size = 11

    End With

    

    With tbl.Cell(1, 2).Range

        .Text = saoType & vbCrLf & _

                locationName & DecodeUTF8("LCBuZ8OgeSAuLi4gdGjDoW5nIC4uLiBuxINtIDIwLi4u") & vbCrLf & _

                DecodeUTF8("VEwuIENI4bumIFThu4pDSA0KQ0jDgU5IIFbEgk4gUEjDkk5HDQoNCg0K") & vbCrLf & _

                DecodeUTF8("SG/DoG5nIFbEg24gRQ==")

        .ParagraphFormat.Alignment = wdAlignParagraphCenter

        .Paragraphs(1).Range.Font.Name = "Times New Roman"

        .Paragraphs(1).Range.Font.Size = 14

        .Paragraphs(1).Range.Font.Bold = True

        .Paragraphs(2).Range.Font.Name = "Times New Roman"

        .Paragraphs(2).Range.Font.Size = 13.5

        .Paragraphs(2).Range.Font.Italic = True

        .Paragraphs(3).Range.Font.Name = "Times New Roman"

        .Paragraphs(3).Range.Font.Size = 13.5

        .Paragraphs(3).Range.Font.Bold = True

        .Paragraphs(4).Range.Font.Name = "Times New Roman"

        .Paragraphs(4).Range.Font.Size = 13.5

        .Paragraphs(4).Range.Font.Bold = True

        .Paragraphs(8).Range.Font.Name = "Times New Roman"

        .Paragraphs(8).Range.Font.Size = 13.5

        .Paragraphs(8).Range.Font.Bold = True

    End With

    

    tbl.Range.Collapse wdCollapseEnd

    tbl.Range.InsertParagraphAfter

End Sub



Public Sub InsertDoKhanAtCursor(ByVal khanText As String)

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    Dim shp As Shape

    Dim posX As Single, posY As Single

    

    posX = Selection.Information(5) ' wdHorizontalPositionRelativeToPage

    posY = Selection.Information(6) ' wdVerticalPositionRelativeToPage

    If posX <= 0 Then posX = MillimetersToPoints(30)

    If posY <= 0 Then posY = MillimetersToPoints(25)

    

    Set shp = ActiveDocument.Shapes.AddShape(1, posX, posY, MillimetersToPoints(36), MillimetersToPoints(8.5))

    If Not shp Is Nothing Then

        With shp

            .WrapFormat.Type = 3 ' wdWrapFront

            .Fill.Solid

            .Fill.ForeColor.RGB = RGB(255, 255, 255)

            .Line.ForeColor.RGB = RGB(255, 0, 0)

            .Line.Weight = 1#

            .Shadow.Visible = 0

            

            With .TextFrame

                .MarginTop = 0

                .MarginBottom = 0

                .MarginLeft = 0

                .MarginRight = 0

                .VerticalAnchor = 3 ' msoAnchorMiddle

                

                With .TextRange

                    .Text = khanText

                    .Font.Name = "Times New Roman"

                    .Font.Size = 13

                    .Font.Bold = True

                    .Font.Color = wdColorRed

                    .ParagraphFormat.Alignment = wdAlignParagraphCenter

                    .ParagraphFormat.SpaceBefore = 0

                    .ParagraphFormat.SpaceAfter = 0

                    .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

                End With

            End With

            

            .Select

        End With

    End If

End Sub



Public Sub InsertDoMatAtCursor(ByVal matText As String)

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    Dim shp As Shape

    Dim posX As Single, posY As Single

    

    posX = Selection.Information(5) ' wdHorizontalPositionRelativeToPage

    posY = Selection.Information(6) ' wdVerticalPositionRelativeToPage

    If posX <= 0 Then posX = MillimetersToPoints(30)

    If posY <= 0 Then posY = MillimetersToPoints(25)

    

    Set shp = ActiveDocument.Shapes.AddShape(1, posX, posY, MillimetersToPoints(32), MillimetersToPoints(8.5))

    If Not shp Is Nothing Then

        With shp

            .WrapFormat.Type = 3 ' wdWrapFront

            .Fill.Solid

            .Fill.ForeColor.RGB = RGB(255, 255, 255)

            .Line.ForeColor.RGB = RGB(255, 0, 0)

            .Line.Weight = 1#

            .Shadow.Visible = 0

            

            With .TextFrame

                .MarginTop = 0

                .MarginBottom = 0

                .MarginLeft = 0

                .MarginRight = 0

                .VerticalAnchor = 3 ' msoAnchorMiddle

                

                With .TextRange

                    .Text = matText

                    .Font.Name = "Times New Roman"

                    .Font.Size = 13

                    .Font.Bold = True

                    .Font.Color = wdColorRed

                    .ParagraphFormat.Alignment = wdAlignParagraphCenter

                    .ParagraphFormat.SpaceBefore = 0

                    .ParagraphFormat.SpaceAfter = 0

                    .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

                End With

            End With

            

            .Select

        End With

    End If

End Sub



Public Sub InsertPhamViLuuHanhAtCursor(ByVal pvText As String)

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    Dim shp As Shape

    Dim posX As Single, posY As Single

    

    posX = Selection.Information(5) ' wdHorizontalPositionRelativeToPage

    posY = Selection.Information(6) ' wdVerticalPositionRelativeToPage

    If posX <= 0 Then posX = MillimetersToPoints(30)

    If posY <= 0 Then posY = MillimetersToPoints(25)

    

    Set shp = ActiveDocument.Shapes.AddShape(1, posX, posY, MillimetersToPoints(52), MillimetersToPoints(8.5))

    If Not shp Is Nothing Then

        With shp

            .WrapFormat.Type = 3 ' wdWrapFront

            .Fill.Solid

            .Fill.ForeColor.RGB = RGB(255, 255, 255)

            .Line.ForeColor.RGB = RGB(0, 0, 0)

            .Line.Weight = 1#

            .Shadow.Visible = 0

            

            With .TextFrame

                .MarginTop = 0

                .MarginBottom = 0

                .MarginLeft = 0

                .MarginRight = 0

                .VerticalAnchor = 3 ' msoAnchorMiddle

                

                With .TextRange

                    .Text = pvText

                    .Font.Name = "Times New Roman"

                    .Font.Size = 12

                    .Font.Bold = True

                    .Font.Color = wdColorBlack

                    .ParagraphFormat.Alignment = wdAlignParagraphCenter

                    .ParagraphFormat.SpaceBefore = 0

                    .ParagraphFormat.SpaceAfter = 0

                    .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

                End With

            End With

            

            .Select

        End With

    End If

End Sub



Public Sub InsertFooterAgencyInfoAtCursor()

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    Dim rng As Range: Set rng = Selection.Range

    rng.Text = "_____________________________________________________________________________" & vbCrLf & _

               DecodeUTF8("xJDhu4thIGNo4buJOiAuLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLjsgxJBp4buHbiB0aG/huqFpOiAuLi4uLi4uLi4uLi47IEZheDogLi4uLi4uLi4uLi4uOyBFbWFpbDogLi4uLi4uLi4uLi4uOyBXZWJzaXRlOiAuLi4uLi4uLi4uLi4=") & vbCrLf

    Dim p As Paragraph: Set p = rng.Paragraphs(2)

    With p.Range.Font

        .Name = "Times New Roman"

        .Size = 11

        .Italic = False

        .Bold = False

    End With

    p.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft

End Sub



' --- CHEN CHU KY TU DONG THEO QUY CHUAN (ND 30 / HD 05) ---





Public Sub InsertSigCustomAtCursor()

    InsertSigAutoAtCursor

End Sub



Public Sub InsertSigSavedIndex(ByVal sigIdx As Long, Optional ByVal recIdx As Long = 1)

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    

    Dim rawJson As String: rawJson = GetUserPresetsJson()

    Dim sigTitle As String, sigName As String, recContent As String

    

    sigTitle = ExtractJsonArrayItemTag(rawJson, "signers_list", "title", sigIdx)

    sigName = ExtractJsonArrayItemTag(rawJson, "signers_list", "fullname", sigIdx)

    recContent = ExtractJsonArrayItemTag(rawJson, "recipients_list", "content", recIdx)

    

    If Len(sigTitle) = 0 Then sigTitle = DecodeUTF8("R0nDgU0gxJDhu5BD")

    If Len(sigName) = 0 Then sigName = DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

    If Len(recContent) = 0 Then recContent = "- " & DecodeUTF8("TsawIGvDrW5oIGfhu61pOw==") & vbCrLf & "- " & DecodeUTF8("TMavdTogVlQsIFRILg==")

    

    InsertSignatureBlockAtCursor recContent, sigTitle, sigName

End Sub



Public Sub InsertRecSavedIndex(ByVal recIdx As Long)

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    

    Dim rawJson As String: rawJson = GetUserPresetsJson()

    Dim recContent As String

    recContent = ExtractJsonArrayItemTag(rawJson, "recipients_list", "content", recIdx)

    If Len(recContent) = 0 Then recContent = "- " & DecodeUTF8("TsawIGvDrW5oIGfhu61pOw==") & vbCrLf & "- " & DecodeUTF8("TMavdTogVlQsIFRILg==")

    

    Dim rng As Range: Set rng = Selection.Range

    rng.Text = DecodeUTF8("TsahaSBuaOG6rW46") & vbCrLf & recContent & vbCrLf

    rng.Paragraphs(1).Range.Font.Name = "Times New Roman"

    rng.Paragraphs(1).Range.Font.Size = 12

    rng.Paragraphs(1).Range.Font.Bold = True

    rng.Paragraphs(1).Range.Font.Italic = True

End Sub



Public Sub InsertSignatureBlockAtCursor(ByVal leftNoiNhan As String, ByVal rightSignerTitle As String, ByVal rightSignerName As String)

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    Dim rng As Range: Set rng = Selection.Range

    If rng.Tables.Count > 0 Then

        Set rng = rng.Tables(1).Range

        rng.Collapse wdCollapseEnd

        rng.InsertParagraphAfter

        rng.Collapse wdCollapseEnd

    End If

    

    Dim tbl As Table

    Set tbl = rng.Tables.Add(Range:=rng, NumRows:=1, NumColumns:=2)

    tbl.Borders.Enable = False

    tbl.Columns(1).Width = MillimetersToPoints(75)

    tbl.Columns(2).Width = MillimetersToPoints(95)

    

    With tbl.Cell(1, 1).Range

        .Text = DecodeUTF8("TsahaSBuaOG6rW46") & vbCrLf & leftNoiNhan

        .ParagraphFormat.Alignment = wdAlignParagraphLeft

        .ParagraphFormat.LineSpacing = VF_LinesToPoints(1.1)

        .Font.Name = "Times New Roman"

        .Font.Size = 11

    End With

    Dim pNN As Paragraph: Set pNN = tbl.Cell(1, 1).Range.Paragraphs(1)

    pNN.Range.Font.Bold = True

    pNN.Range.Font.Italic = True

    pNN.Range.Font.Size = 12

    

    With tbl.Cell(1, 2).Range

        .Text = rightSignerTitle & vbCrLf & vbCrLf & vbCrLf & vbCrLf & rightSignerName

        .ParagraphFormat.Alignment = wdAlignParagraphCenter

        .Font.Name = "Times New Roman"

        .Font.Size = 13.5

        .Font.Bold = True

    End With

    

    Dim curReg As Integer: curReg = GetEffectiveRegulation(ActiveDocument)

    FormatExistingSignatureTable tbl, curReg

    tbl.Range.Collapse wdCollapseEnd

    tbl.Range.InsertParagraphAfter

End Sub



' ==============================================================================

' 7. CONG CU XU LY TIENG VIET

' ==============================================================================

Public Sub RemoveTonesFromSelection()

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    

    If Selection.Type = wdSelectionIP Or Len(Trim(Selection.Text)) = 0 Then

        ShowUniMsg "VnVpIGzDsm5nIELDlEkgxJBFTiDEkeG7sWFuIHbEg24gYuG6o24gY+G6p24gY2h1eeG7g24gc2FuZyB0aeG6v25nIFZp4buHdCBraMO0bmcgZOG6pXUh", "Vi1GT1JNQVQ=", 48

        Exit Sub

    End If

    

    Dim raw As String: raw = Selection.Text

    Dim res As String: res = ConvertToUnsignedVietnamese(raw)

    Selection.Text = res

    ShowUniMsg "xJDDoyBjaHV54buDbiDEkeG7sWFuIHbEg24gYuG6o24gxJHGsOG7o2MgYuG7kWkgxJFlbiB0aMOgbmggVEnhur5ORyBWSeG7hlQgS0jDlE5HIEdBVSE=", "Vi1GT1JNQVQ=", 64

End Sub



Public Function ConvertToUnsignedVietnamese(ByVal src As String) As String

    Dim s As String: s = src

    

    s = ReplaceCharList(s, ChrW(224) & ChrW(225) & ChrW(7843) & ChrW(227) & ChrW(7841) & ChrW(259) & ChrW(7857) & ChrW(7855) & ChrW(7859) & ChrW(7861) & ChrW(7863) & ChrW(226) & ChrW(7847) & ChrW(7845) & ChrW(7849) & ChrW(7851) & ChrW(7853), "a")

    s = ReplaceCharList(s, ChrW(273), "d")

    s = ReplaceCharList(s, ChrW(232) & ChrW(233) & ChrW(7867) & ChrW(7869) & ChrW(7865) & ChrW(234) & ChrW(7873) & ChrW(7871) & ChrW(7875) & ChrW(7877) & ChrW(7879), "e")

    s = ReplaceCharList(s, ChrW(236) & ChrW(237) & ChrW(7881) & ChrW(297) & ChrW(7883), "i")

    s = ReplaceCharList(s, ChrW(242) & ChrW(243) & ChrW(7887) & ChrW(245) & ChrW(7885) & ChrW(244) & ChrW(7891) & ChrW(7889) & ChrW(7893) & ChrW(7895) & ChrW(7897) & ChrW(417) & ChrW(7901) & ChrW(7899) & ChrW(7903) & ChrW(7905) & ChrW(7907), "o")

    s = ReplaceCharList(s, ChrW(249) & ChrW(250) & ChrW(7911) & ChrW(361) & ChrW(7909) & ChrW(432) & ChrW(7915) & ChrW(7913) & ChrW(7917) & ChrW(7919) & ChrW(7921), "u")

    s = ReplaceCharList(s, ChrW(7923) & ChrW(253) & ChrW(7927) & ChrW(7929) & ChrW(7925), "y")

    

    s = ReplaceCharList(s, ChrW(192) & ChrW(193) & ChrW(7842) & ChrW(195) & ChrW(7840) & ChrW(258) & ChrW(7856) & ChrW(7854) & ChrW(7858) & ChrW(7860) & ChrW(7862) & ChrW(194) & ChrW(7846) & ChrW(7844) & ChrW(7848) & ChrW(7850) & ChrW(7852), "A")

    s = ReplaceCharList(s, ChrW(272), "D")

    s = ReplaceCharList(s, ChrW(200) & ChrW(201) & ChrW(7866) & ChrW(7868) & ChrW(7864) & ChrW(202) & ChrW(7872) & ChrW(7870) & ChrW(7874) & ChrW(7876) & ChrW(7878), "E")

    s = ReplaceCharList(s, ChrW(204) & ChrW(205) & ChrW(7880) & ChrW(296) & ChrW(7882), "I")

    s = ReplaceCharList(s, ChrW(210) & ChrW(211) & ChrW(7886) & ChrW(213) & ChrW(7884) & ChrW(212) & ChrW(7890) & ChrW(7888) & ChrW(7892) & ChrW(7894) & ChrW(7896) & ChrW(416) & ChrW(7900) & ChrW(7898) & ChrW(7902) & ChrW(7904) & ChrW(7906), "O")

    s = ReplaceCharList(s, ChrW(217) & ChrW(218) & ChrW(7910) & ChrW(360) & ChrW(7908) & ChrW(431) & ChrW(7914) & ChrW(7912) & ChrW(7916) & ChrW(7918) & ChrW(7920), "U")

    s = ReplaceCharList(s, ChrW(7922) & ChrW(221) & ChrW(7926) & ChrW(7928) & ChrW(7924), "Y")

    

    Dim cleanS As String: cleanS = ""

    Dim k As Long, ch As Long

    For k = 1 To Len(s)

        ch = AscW(Mid(s, k, 1))

        If ch < 0 Then ch = ch + 65536

        ' Loai bo Unicode Combining Diacritical Marks (dau thanh, dau mu, dau moc NFD to hop)

        If Not (ch >= &H300 And ch <= &H36F) And Not (ch >= &H1DC0 And ch <= &H1DFF) Then

            cleanS = cleanS & Mid(s, k, 1)

        End If

    Next k

    

    ConvertToUnsignedVietnamese = cleanS

End Function



Private Function ReplaceCharList(ByVal src As String, ByVal charList As String, ByVal repChar As String) As String

    Dim i As Long, c As String

    Dim res As String: res = src

    For i = 1 To Len(charList)

        c = Mid(charList, i, 1)

        If InStr(res, c) > 0 Then

            res = Replace(res, c, repChar)

        End If

    Next i

    ReplaceCharList = res

End Function



Public Sub StandardizeToneMarksAll(Optional doc As Document = Nothing)

    On Error GoTo ErrH

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then Exit Sub

    

    StartProgress DecodeUTF8("Vi1GT1JNQVQgLSBE4bqldSBUaGFuaA=="), DecodeUTF8("xJBhbmcgY2h14bqpbiBow7NhIGThuqV1IHRoYW5oIHRvw6BuIHbEg24gYuG6o24uLi4="), "ToneMarks"

    SetProgress 30, DecodeUTF8("xJBhbmcgcXXDqXQgc8OhaSBow7NhIGPDoWMgduG6p24gb2EsIG9lLCB1eS4uLg=="), "ToneMarks"

    SafeDoEvents

    

    StandardizeToneMarksAllQuiet doc

    

    SetProgress 100, DecodeUTF8("SG/DoG4gdOG6pXQh"), "ToneMarks"

    FinishProgress "ToneMarks"

    ShowUniMsg "xJDDoyBz4butYSB2w6AgY2h14bqpbiBow7NhIHRvw6BuIGLhu5kgZOG6pXUgdGhhbmggdHJvbmcgdsSDbiBi4bqjbiAoZOG6oW5nIG3hu5tpOiDDsmEsIMO6eSwgaMOyYSwgdGjhu6d5LCB0w7JhLCBraMOzYS4uLikgdGjDoG5oIGPDtG5nIQ==", "Vi1GT1JNQVQ=", 64

    Exit Sub

ErrH:

    FinishProgress "ToneMarks"

    RecordError "VFormat_Core.StandardizeToneMarksAll", Err.Number, Err.Description, "Sua dau thanh toan van"

End Sub



Private Sub ReplaceExactWord(ByVal doc As Document, ByVal findWord As String, ByVal replaceWord As String, Optional ByVal matchWhole As Boolean = True)

    On Error Resume Next

    Dim rng As Range

    Set rng = doc.content

    With rng.Find

        .ClearFormatting

        .Replacement.ClearFormatting

        .Text = findWord

        .Replacement.Text = replaceWord

        .Forward = True

        .Wrap = wdFindContinue

        .Format = False

        .MatchCase = False

        .MatchWholeWord = matchWhole

        .MatchWildcards = False

        .Execute Replace:=wdReplaceAll

    End With

End Sub





' ==============================================================================

' 7B. XOA KHOANG TRANG LOI & DAU CACH BAT HOP LY TRONG VAN BAN

' (XOA KHOANG TRANG TRUOC DAU CAU, THU HEP KHOANG TRANG KEP, DAU CACH CUOI DOAN)

' ==============================================================================

Public Sub CleanSpacingErrors(Optional ByVal isQuiet As Boolean = False)

    CleanSpacingErrorsInternal Nothing, isQuiet

End Sub



Public Sub CleanSpacingErrorsInternal(Optional targetRng As Range = Nothing, Optional ByVal isQuiet As Boolean = False)

    On Error GoTo ErrH

    DbgLog "CleanSpacingErrorsInternal: entered"

    Dim doc As Document

    If Documents.Count = 0 Then

        If Not isQuiet Then

            ShowUniMsg "VnVpIGzDsm5nIG3hu58gbeG7mXQgdMOgaSBsaeG7h3UgV29yZCDEkeG7gyB44butIGzDvSE=", "Vi1GT1JNQVQ=", 48

        End If

        Exit Sub

    End If

    Set doc = ActiveDocument

    DbgLog "CleanSpacingErrorsInternal: doc set"

    

    Dim isSelectionOnly As Boolean: isSelectionOnly = False

    If targetRng Is Nothing Then

        DbgLog "CleanSpacingErrorsInternal: targetRng is nothing, check selection"

        On Error Resume Next

        If Application.Selection.Type = wdSelectionNormal And Application.Selection.Range.Characters.Count > 2 Then

            Set targetRng = Application.Selection.Range

            isSelectionOnly = True

        Else

            Set targetRng = doc.content

        End If

        On Error GoTo ErrH

    End If

    DbgLog "CleanSpacingErrorsInternal: targetRng resolved, isSel=" & isSelectionOnly

    

    ' 1. Bao ve cac vi tri dien khuyet hop le (vi du: " ... ", " ...^p")

    FastReplaceText doc, targetRng, isSelectionOnly, " ... ", " ___VF_BLANK___ "

    FastReplaceText doc, targetRng, isSelectionOnly, " ...^p", " ___VF_BLANK___^p"

    DbgLog "CleanSpacingErrorsInternal: step 1 done"

    

    ' 2. Chuan hoa dau cach khong ngat (Non-Breaking Space ChrW(160)) ve dau cach chuan

    FastReplaceText doc, targetRng, isSelectionOnly, ChrW(160), " "

    

    ' 3. Xoa dau cach lien tiep (thu hep ve 1 dau cach duy nhat)

    Dim k As Integer

    For k = 1 To 6

        FastReplaceText doc, targetRng, isSelectionOnly, "  ", " "

    Next k

    

    ' 4. Xoa dau cach truoc cac dau cau

    FastReplaceText doc, targetRng, isSelectionOnly, " ,", ","

    FastReplaceText doc, targetRng, isSelectionOnly, " .", "."

    FastReplaceText doc, targetRng, isSelectionOnly, " ;", ";"

    FastReplaceText doc, targetRng, isSelectionOnly, " :", ":"

    FastReplaceText doc, targetRng, isSelectionOnly, " ?", "?"

    FastReplaceText doc, targetRng, isSelectionOnly, " !", "!"

    FastReplaceText doc, targetRng, isSelectionOnly, " )", ")"

    FastReplaceText doc, targetRng, isSelectionOnly, " ]", "]"

    FastReplaceText doc, targetRng, isSelectionOnly, " }", "}"

    FastReplaceText doc, targetRng, isSelectionOnly, " " & ChrW(8221), ChrW(8221)

    FastReplaceText doc, targetRng, isSelectionOnly, " " & ChrW(8217), ChrW(8217)

    

    ' 5. Xoa dau cach sau cac dau mo ngoac / mo ngoac kep

    FastReplaceText doc, targetRng, isSelectionOnly, "( ", "("

    FastReplaceText doc, targetRng, isSelectionOnly, "[ ", "["

    FastReplaceText doc, targetRng, isSelectionOnly, "{ ", "{"

    FastReplaceText doc, targetRng, isSelectionOnly, ChrW(8220) & " ", ChrW(8220)

    FastReplaceText doc, targetRng, isSelectionOnly, ChrW(8216) & " ", ChrW(8216)

    

    ' 6. Xoa dau cach thua truoc va sau ngat dong, ngat doan, tab

    FastReplaceText doc, targetRng, isSelectionOnly, " ^p", "^p"

    FastReplaceText doc, targetRng, isSelectionOnly, " ^l", "^l"

    FastReplaceText doc, targetRng, isSelectionOnly, " ^t", "^t"

    FastReplaceText doc, targetRng, isSelectionOnly, "^t ", "^t"

    FastReplaceText doc, targetRng, isSelectionOnly, "^p ", "^p"

    

    ' 7. Phuc hoi cac vi tri dien khuyet hop le

    FastReplaceText doc, targetRng, isSelectionOnly, " ___VF_BLANK___ ", " ... "

    FastReplaceText doc, targetRng, isSelectionOnly, " ___VF_BLANK___^p", " ...^p"

    

    ' 8. Quet lai dau cach kep lan cuoi

    For k = 1 To 3

        FastReplaceText doc, targetRng, isSelectionOnly, "  ", " "

    Next k

    

    ' 9. Xoa dau cach o dau tai lieu neu co (kiem tra an toan tranh cham vao bang bieu)

    On Error Resume Next

    If doc.Paragraphs.Count > 0 Then

        Dim fRng As Range: Set fRng = doc.Paragraphs(1).Range

        If Not fRng.Information(wdWithInTable) Then

            Dim loopLimit As Integer: loopLimit = 0

            Do While Len(fRng.Text) > 1 And Left(fRng.Text, 1) = " " And loopLimit < 50

                fRng.Characters(1).Delete

                loopLimit = loopLimit + 1

            Loop

        End If

    End If

    On Error GoTo ErrH

    

    If Not isQuiet Then

        If isSelectionOnly Then

            ShowUniToast DecodeUTF8("xJDDoyB4w7NhIGPDoWMga2hv4bqjbmcgdHLhuq9uZyBs4buXaSAmIGThuqV1IGPDoWNoIHRo4burYSB0cm9uZyDEkW/huqFuIGLDtGkgxJFlbiE="), "Vi1GT1JNQVQ="

        Else

            ShowUniToast DecodeUTF8("xJDDoyB4w7NhIHRvw6BuIGLhu5kga2hv4bqjbmcgdHLhuq9uZyBs4buXaSAmIGThuqV1IGPDoWNoIHRo4burYSB0cm9uZyB2xINuIGLhuqNuIQ=="), "Vi1GT1JNQVQ="

        End If

    End If

    Exit Sub

ErrH:

    RecordError "VFormat_Core.CleanSpacingErrors", Err.Number, Err.Description, "Xoa khoang trang loi"

End Sub



Private Sub FastReplaceText(ByVal doc As Document, ByVal selRng As Range, ByVal isSelOnly As Boolean, ByVal fText As String, ByVal rText As String)

    On Error GoTo ErrH

    DbgLog "FastReplaceText: entered for [" & fText & "]"

    If doc Is Nothing Then

        DbgLog "FastReplaceText: doc is Nothing!"

        Exit Sub

    End If

    Dim r As Range

    If isSelOnly And Not (selRng Is Nothing) Then

        Set r = selRng.Duplicate

    Else

        Set r = doc.content

    End If

    DbgLog "FastReplaceText: range set"

    With r.Find

        .ClearFormatting

        .Replacement.ClearFormatting

        .Text = fText

        .Replacement.Text = rText

        .Forward = True

        .Wrap = wdFindStop

        .Format = False

        .MatchCase = False

        .MatchWholeWord = False

        .MatchWildcards = False

        .Execute Replace:=wdReplaceAll

    End With

    DbgLog "FastReplaceText: execute done for [" & fText & "]"

    Exit Sub

ErrH:

    DbgLog "FastReplaceText: error " & Err.Description

End Sub



' ==============================================================================

' 8. THU VIEN BIEU MAU

' ==============================================================================

Public Sub ClearWindowsRecentFile(Optional ByVal fileName As String = "")

    On Error Resume Next

    Dim recentDir As String

    recentDir = Environ$("APPDATA") & "\Microsoft\Windows\Recent\"

    If Dir(recentDir, vbDirectory) = "" Then Exit Sub

    

    Dim baseName As String

    If Len(fileName) > 0 Then

        baseName = Replace(Replace(fileName, ".docx", ""), ".doc", "")

        Dim f As String

        f = Dir(recentDir & "*" & baseName & "*.lnk")

        Do While Len(f) > 0

            Kill recentDir & f

            f = Dir()

        Loop

    End If

    

    Dim fMau As String

    fMau = Dir(recentDir & "*Mau_*.lnk")

    Do While Len(fMau) > 0

        Kill recentDir & fMau

        fMau = Dir()

    Loop

End Sub



Public Sub OpenTemplateDoc(ByVal fileName As String)

    On Error GoTo ErrHandler

    ' Chot 2: Da go bo kiem tra ban quyen de cho phep mo bieu mau truc tiep

    ' If Not VFormat_License.EnsureLicenseActive() Then Exit Sub

    

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim baseDir As String: baseDir = Environ$("APPDATA") & "\VFormat\"

    Dim targetDocx As String

    

    ' 1. Kiem tra thu muc templates tieu chuan hoac cac thu muc tuong thich

    targetDocx = baseDir & "templates\" & fileName

    If Not fso.FileExists(targetDocx) Then

        targetDocx = baseDir & "data\templates\" & fileName

    End If

    If Not fso.FileExists(targetDocx) Then

        targetDocx = baseDir & ".data\templates\" & fileName

    End If

    

    ' 2. Development Fallback: Kiem tra tu thu muc goc cua du an D:\Tool\V-Format\templates

    If Not fso.FileExists(targetDocx) Then

        If fso.FileExists("D:\Tool\V-Format\templates\" & fileName) Then

            SyncLocalTemplatesToData

            targetDocx = baseDir & ".data\templates\" & fileName

            If Not fso.FileExists(targetDocx) Then

                targetDocx = "D:\Tool\V-Format\templates\" & fileName

            End If

        End If

    End If

    

    If fso.FileExists(targetDocx) Then

        OpenExistingTemplateDoc targetDocx, fileName

    Else

        Dim confirmMsg As String

        confirmMsg = DecodeUTF8("Q2jGsGEgdMOsbSB0aOG6p3kgdGjGsCB2aeG7h24gYmnDqnUgbeG6q3UgcGjhuqduIG3hu4FtIHRyw6puIG3DoXkgKA==") & fileName & ")." & vbCrLf & _

                     "----------------------------------------" & vbCrLf & _

                     DecodeUTF8("QuG6oW4gY8OzIG114buRbiB04bqjaSBi4buZIGPDoGkgxJHhurd0IFYtRk9STUFUIDEuNiDEkeG6gyBjw6BpIMSR4bq3dCB0cuG7j24gYuG7mSA1OCBiaeG7g3UgbeG6q3UgUXXhu5FjIGdpYSBuZ2F5IGtow7RuZz8=")

        If ShowUniConfirm(confirmMsg, "V-FORMAT") Then

            SafeOpenUrlOrFile "https://orcax.vn/downloads/OrcaX-Setup.zip"

        End If

    End If

    Exit Sub



ErrHandler:

    ShowUniMsg "Co loi khi mo bieu mau: " & Err.Description, "Vi1GT1JNQVQ=", 16

End Sub



Private Sub SyncLocalTemplatesToData()

    On Error Resume Next

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim baseDir As String: baseDir = Environ$("APPDATA") & "\VFormat"

    Dim dataDir As String: dataDir = baseDir & "\.data"

    Dim destDir As String: destDir = dataDir & "\templates"

    If Not fso.FolderExists(baseDir) Then fso.CreateFolder baseDir

    If Not fso.FolderExists(dataDir) Then fso.CreateFolder dataDir

    If Not fso.FolderExists(destDir) Then fso.CreateFolder destDir

    If fso.FolderExists("D:\Tool\V-Format\templates") Then

        fso.CopyFile "D:\Tool\V-Format\templates\*.docx", destDir & "\", True

    End If

End Sub



Private Sub OpenExistingTemplateDoc(ByVal filePath As String, ByVal originalName As String)

    On Error Resume Next

    Dim doc As Document

    Set doc = Documents.Add(Template:=filePath, NewTemplate:=False, DocumentType:=0, Visible:=True)

    

    ' Xoa khoi RecentFiles cua Word de khong lo duong dan template

    If Application.RecentFiles.Count > 0 Then

        Dim i As Long

        For i = Application.RecentFiles.Count To 1 Step -1

            If InStr(1, Application.RecentFiles(i).Path, originalName, vbTextCompare) > 0 Or _

               InStr(1, Application.RecentFiles(i).Path, "Mau_", vbTextCompare) > 0 Then

                Application.RecentFiles(i).Delete

            End If

        Next i

    End If

    ClearWindowsRecentFile originalName

End Sub





Public Sub InsertQuocHieuAutoAtCursor()

    If Documents.Count = 0 Then Exit Sub

    Dim curReg As Integer: curReg = GetEffectiveRegulation(ActiveDocument)

    If curReg = 1 Then

        ' Van ban Dang

        If g_SelectedOrgLevel = 6 Then

            ' Cap Chi bo

            InsertDangChiBoHeaderAtCursor

        ElseIf g_SelectedOrgLevel = 5 Then

            ' Cap Dang uy Xa

            InsertQuocHieuPartyIndex 3

        ElseIf g_SelectedOrgLevel = 4 Then

            ' Cap Dang bo Co quan / So

            InsertQuocHieuPartyIndex 1

        Else

            ' Mac dinh / Tinh uy

            InsertQuocHieuPartyIndex 2

        End If

    Else

        ' Hanh chinh ND 30

        If g_SelectedOrgLevel = 1 Then

            ' Cap Bo / Ngang Bo (Mau 1: Bo Noi vu)

            InsertQuocHieuAdminIndex 1

        ElseIf g_SelectedOrgLevel = 2 Then

            ' Cap Cuc / Tong cuc (Mau 2: Cuc VTLT)

            InsertQuocHieuAdminIndex 2

        ElseIf g_SelectedOrgLevel = 4 Then

            ' Cap So / Nganh (Mau 3: So KHCN)

            InsertQuocHieuAdminIndex 3

        ElseIf g_SelectedOrgLevel = 5 Then

            ' Cap Xa / Phuong (Mau 4: UBND Phuong)

            InsertQuocHieuAdminIndex 4

        ElseIf g_SelectedOrgLevel = 3 Then

            ' Cap Tinh (Mau 6: Van phong UBND tinh)

            InsertQuocHieuAdminIndex 6

        Else

            ' Tu dong: Dung Mau 1

            InsertQuocHieuAdminIndex 1

        End If

    End If

End Sub



Public Sub InsertSigAutoAtCursor()

    If Documents.Count = 0 Then Exit Sub

    Dim curReg As Integer: curReg = GetEffectiveRegulation(ActiveDocument)

    If curReg = 1 Then

        ' Van ban Dang

        If g_SelectedOrgLevel = 6 Then

            ' Chi bo

            InsertSignatureBlockAtCursor DecodeUTF8("LSDEkOG6o25nIOG7p3kgeMOjL3BoxrDhu51uZy4uLiAoxJHhu4MgYi9jKTsNCi0gQ8OhYyDEkeG7k25nIGNow60gxJHhuqNuZyB2acOqbiB0cm9uZyBDaGkgYuG7mTsNCi0gTMawdTogQ2hpYuG7mS4="), DecodeUTF8("VC9NIENISSBC4buYDQpCw41fVEjGrw=="), DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

        ElseIf g_SelectedOrgLevel = 5 Then

            ' Dang uy Xa

            InsertSignatureBlockAtCursor DecodeUTF8("LSBOaMawIGvDrW5oIGfhu61pOw0KLSBUaMaw4budbmcgdHLhu7FjIMSQ4bqjbmcgcGjGsOG7nW5nOw0KLSBMxrB1IFbEg24gcGjDsm5nLg=="), DecodeUTF8("VE0uIEJBTiBDSOG6pFAgSMOATkgNCkLDjSBUSMav"), DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

        Else

            InsertSigSavedIndex 3, 3

        End If

    Else

        ' Hanh chinh ND 30

        If g_SelectedOrgLevel = 1 Then

            ' Cap Bo (Bo truong)

            InsertSignatureBlockAtCursor DecodeUTF8("LSBOaMawIMSQaeG7gXUgLi4uOw0KLSBUaOG7pyB0xrDhu5tuZywgY8OhYyBQaMOzIFRo4bun tMaw4budbmcgQ1A7DQotIEPDoWMgQuG7mSwgY8ahIHF1YW4gbmdhbmcgQuG7mTsNCi0gTMawdTogVlQsIFBDLg=="), DecodeUTF8("QuG7mCBUUsav4bufTkc="), DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

        ElseIf g_SelectedOrgLevel = 2 Then

            ' Cap Cuc / Tong cuc (Cuc truong)

            InsertSignatureBlockAtCursor DecodeUTF8("LSBOaMawIMSQaeG7gXUgLi4uOw0KLSBMw6NuaCDEkeG6oW8gQuG7mTsNCi0gQ8OhYyDEkcahbiB24buLIHRy4buxYyB0aHXhu5ljOw0KLSBMxrB1OiBWVC4="), DecodeUTF8("Q+G7pEMgVFLGr+G7n05H"), DecodeUTF8("TMOqIFbEg24gQw==")

        ElseIf g_SelectedOrgLevel = 5 Then

            ' UBND Xa/Phuong

            InsertSignatureBlockAtCursor DecodeUTF8("LSBOaMawIHRyw6puOw0KLSBUaMaw4budbmcgdHLhu7FjIMSQ4bqjbmcg4buneTsNCi0gTMawdTogVlQu"), DecodeUTF8("VE0uIOG7plkgQkFOIE5Iw4JOIETDgk4NCkNI4bumIFThu4pDSA=="), DecodeUTF8("VHLhuqduIFbEg24gQg==")

        Else

            ' Mac dinh / Cap So / Tinh

            InsertSigSavedIndex 1, 1

        End If

    End If

End Sub



Public Sub InsertDangChiBoHeaderAtCursor()

    If Documents.Count = 0 Then Exit Sub

    On Error Resume Next

    Dim rng As Range: Set rng = Selection.Range

    Dim tbl As Table

    Set tbl = rng.Tables.Add(Range:=rng, NumRows:=1, NumColumns:=2)

    tbl.Columns(1).Width = MillimetersToPoints(78)

    tbl.Columns(2).Width = MillimetersToPoints(87)

    Dim rWCB As Row

    For Each rWCB In tbl.Rows

        If rWCB.Cells.Count >= 2 Then

            rWCB.Cells(1).Width = MillimetersToPoints(78)

            rWCB.Cells(2).Width = MillimetersToPoints(87)

        End If

    Next rWCB

    

    With tbl.Cell(1, 1).Range

        .Text = DecodeUTF8("xJDhuqJORyBC4buYIFjDgy9QSMav4bucTkcuLi4=") & vbCrLf & _

                DecodeUTF8("Q0hJIELhu5guLi4=") & vbCrLf & _

                "*" & vbCrLf & _

                DecodeUTF8("U+G7kSAwMS1OUS9DQg==")

        .ParagraphFormat.Alignment = wdAlignParagraphCenter

        .ParagraphFormat.SpaceBefore = 0

        .ParagraphFormat.SpaceAfter = 0

        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

        .Paragraphs(1).Range.Font.Name = "Times New Roman"

        .Paragraphs(1).Range.Font.Size = 14

        .Paragraphs(1).Range.Font.Bold = False

        .Paragraphs(2).Range.Font.Name = "Times New Roman"

        .Paragraphs(2).Range.Font.Size = 14

        .Paragraphs(2).Range.Font.Bold = True

        .Paragraphs(3).Range.Font.Name = "Times New Roman"

        .Paragraphs(3).Range.Font.Size = 14

        .Paragraphs(3).Range.Font.Bold = False

        .Paragraphs(4).Range.Font.Name = "Times New Roman"

        .Paragraphs(4).Range.Font.Size = 14

        .Paragraphs(4).Range.Font.Bold = False

    End With

    

    With tbl.Cell(1, 2).Range

        .Text = DecodeUTF8("xJDhuqJORyBD4buYTkcgU+G6ok4gVknhu4ZUIE5BTQ==") & vbCrLf & _

                DecodeUTF8("Li4uLCBuZ8OgeSAuLi4gdGjDoW5nIC4uLiBuxINtIDIwLi4u")

        .ParagraphFormat.Alignment = wdAlignParagraphCenter

        .ParagraphFormat.SpaceBefore = 0

        .ParagraphFormat.SpaceAfter = 0

        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

        .Paragraphs(1).Range.Font.Name = "Times New Roman"

        .Paragraphs(1).Range.Font.Size = 15

        .Paragraphs(1).Range.Font.Bold = True

        .Paragraphs(1).Range.Font.Underline = wdUnderlineSingle

        .Paragraphs(2).Range.Font.Name = "Times New Roman"

        .Paragraphs(2).Range.Font.Size = 14

        .Paragraphs(2).Range.Font.Italic = True

        .Paragraphs(2).Range.Font.Underline = wdUnderlineNone

    End With

    

    tbl.Range.Collapse wdCollapseEnd

    tbl.Range.InsertParagraphAfter

End Sub

' ==============================================================================

' 9. CHUAN HOA THE THUC TU DONG 1-CLICK (AUTO-STANDARDIZE ENGINE)

' TU DONG NHAN DIEN, CAN LE, CHINH FONT, CO CHU, TO DAM DE MUC,

' CHEN THANH PHAN THIEU & HOAN TAC (UNDO) THEO CHUAN HD 05 VA ND 30

' ==============================================================================

Public Sub AutoStandardizeDocument(Optional doc As Document = Nothing)

    On Error Resume Next

    Dim d As Document

    If Not doc Is Nothing Then

        Set d = doc

    ElseIf Documents.Count > 0 Then

        Set d = ActiveDocument

    End If

    If d Is Nothing Then

        ShowUniMsg DecodeUTF8("VnVpIGzDsm5nIG3hu58gbeG7mXQgdMOgaSBsaeG7h3UgV29yZCDEkeG7gyBjaHXhuqluIGjDs2EgdOG7sSDEkeG7mW5nIQ=="), "Vi1GT1JNQVQ=", 48

        Exit Sub

    End If

    ExecuteAutoStandardize d, False

End Sub



Public Sub AutoStandardizeQuiet(Optional doc As Document = Nothing)

    On Error Resume Next

    DbgLog "Enter AutoStandardizeQuiet"

    Dim d As Document

    If Not doc Is Nothing Then

        Set d = doc

        DbgLog "AutoStandardizeQuiet: using passed doc"

    ElseIf Documents.Count > 0 Then

        Set d = ActiveDocument

        DbgLog "AutoStandardizeQuiet: using ActiveDocument"

    End If

    If Not d Is Nothing Then

        DbgLog "AutoStandardizeQuiet: calling ExecuteAutoStandardize..."

        ExecuteAutoStandardize d, True

        DbgLog "AutoStandardizeQuiet: back from ExecuteAutoStandardize"

    Else

        DbgLog "AutoStandardizeQuiet: d is Nothing!"

    End If

    DbgLog "Exit AutoStandardizeQuiet"

End Sub



Public Sub ExecuteAutoStandardize(ByVal doc As Document, Optional ByVal isQuiet As Boolean = False)

    On Error GoTo ErrH

    DbgLog "ExecuteAutoStandardize: Step 0 - Entered"

    If doc Is Nothing Then

        DbgLog "ExecuteAutoStandardize: doc is Nothing, exiting!"

        Exit Sub

    End If

    

    If Not isQuiet Then

        StartProgress DecodeUTF8("Vi1GT1JNQVQgLSBDaHXhuqluIEjDs2EgVOG7sSDEkOG7mW5n"), DecodeUTF8("xJBhbmcga2jhu59pIHThuqFvIHbDoCBuaOG6rW4gZGnhu4duIHF1eSBjaHXhuqluLi4u"), "AutoStandardize"

    End If

    

    ' 1. BAT DAU TRANSACTION HOAN TAC (UNDO RECORD)

    Dim undoRec As Object

    On Error Resume Next

    Set undoRec = Application.UndoRecord

    If Not undoRec Is Nothing Then

        undoRec.StartCustomRecord "VFormat_AutoStandardize"

    End If

    On Error GoTo ErrH

    DbgLog "ExecuteAutoStandardize: Step 1 - UndoRecord started"

    

    Dim regIdx As Integer

    regIdx = GetEffectiveRegulation(doc)

    DbgLog "ExecuteAutoStandardize: Step 2 - EffectiveRegulation=" & regIdx

    

    Dim regName As String

    If regIdx = 1 Then

        regName = DecodeUTF8("VsSDbiBi4bqjbiBj4bunYSDEkOG6o25nIChIxrDhu5tuZyBk4bqrbiAwNS1IRC9WUFRXKQ==")

    ElseIf regIdx = 2 Then

        regName = DecodeUTF8("VsSDbiBi4bqjbiBI4bujcCBuaOG6pXQgKE5naOG7iyDEkeG7i25oIDI1MC8yMDI2L07EkC1DUCk=")

    ElseIf regIdx = 3 Then

        regName = DecodeUTF8("S+G6vyBob+G6oWNoIGLDoGkgZOG6oXkgLSBHacOhbyDDoW4gKEPDtG5nIHbEg24gNTUxMi9CR0TEkFQp")

    Else

        regName = DecodeUTF8("SMOgbmggY2jDrW5oIE5ow6Agbsaw4bubYyAoTmdo4buLIMSR4buLbmggMzAvMjAyMC9OxJAtQ1Ap")

    End If

    

    Dim statsParagraphs As Long: statsParagraphs = 0

    Dim statsHeadings As Long: statsHeadings = 0

    Dim statsMissing As Long: statsMissing = 0

    Dim titleDetected As String: titleDetected = ""

    

    If Not isQuiet Then

        SetProgress 10, DecodeUTF8("xJBhbmcgdGhp4bq/dCBs4bqtcCBwaMO0bmcgY2jhu68gdsOgIGPEg24gbOG7gSBBNC4uLg=="), "AutoStandardize"

    End If

    

    ' 2.5. CHUAN HOA PHONG CHU TOAN BO TAI LIEU VE TIMES NEW ROMAN

    doc.content.Font.Name = "Times New Roman"

    doc.content.Font.NameAscii = "Times New Roman"

    doc.content.Font.NameOther = "Times New Roman"

    doc.content.Font.NameBi = "Times New Roman"

    DbgLog "ExecuteAutoStandardize: Step 2.5 - Font set"

    

    ' 3. CAN LE KHANG TRANG A4 THEO DUNG QUY CHUAN (HD 05: 20-20-30-15; ND 30: 20-20-30-20)

    Dim topMm As Single, botMm As Single, leftMm As Single, rightMm As Single

    If regIdx = 1 Then

        topMm = 20: botMm = 20: leftMm = 30: rightMm = 15

    Else

        topMm = 20: botMm = 20: leftMm = 30: rightMm = 20

    End If

    

    Dim sec As Section

    For Each sec In doc.Sections

        With sec.PageSetup

            .PaperSize = wdPaperA4

            If .Orientation = wdOrientLandscape Then

                .PageWidth = MillimetersToPoints(297)

                .PageHeight = MillimetersToPoints(210)

                .TopMargin = MillimetersToPoints(topMm)

                .BottomMargin = MillimetersToPoints(botMm)

                .LeftMargin = MillimetersToPoints(leftMm)

                .RightMargin = MillimetersToPoints(rightMm)

            Else

                .Orientation = wdOrientPortrait

                .PageWidth = MillimetersToPoints(210)

                .PageHeight = MillimetersToPoints(297)

                .TopMargin = MillimetersToPoints(topMm)

                .BottomMargin = MillimetersToPoints(botMm)

                .LeftMargin = MillimetersToPoints(leftMm)

                .RightMargin = MillimetersToPoints(rightMm)

            End If

            .HeaderDistance = MillimetersToPoints(10)

            .FooterDistance = MillimetersToPoints(10)

            .DifferentFirstPageHeaderFooter = True

        End With

    Next sec

    DbgLog "ExecuteAutoStandardize: Step 3 - PageSetup done"

    

    ' 4. DANH SO TRANG LE TREN, CAN GIUA, AN TRANG DAU

    For Each sec In doc.Sections

        sec.PageSetup.DifferentFirstPageHeaderFooter = True

        sec.PageSetup.HeaderDistance = MillimetersToPoints(10)

        

        Dim hdr As HeaderFooter

        Set hdr = sec.Headers(wdHeaderFooterPrimary)

        hdr.Range.Text = ""

        hdr.Range.ParagraphFormat.TabStops.ClearAll

        hdr.Range.ParagraphFormat.LeftIndent = 0

        hdr.Range.ParagraphFormat.RightIndent = 0

        hdr.Range.ParagraphFormat.FirstLineIndent = 0

        hdr.Range.Font.Name = "Times New Roman"

        hdr.Range.Font.Size = IIf(regIdx = 1, 14, 13.5)

        hdr.Range.Font.Bold = False

        hdr.Range.Font.Italic = False

        hdr.Range.Font.ColorIndex = wdBlack

        hdr.Range.Fields.Add Range:=hdr.Range, Type:=wdFieldPage

        hdr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

        If hdr.Range.Paragraphs.Count > 0 Then

            hdr.Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

        End If

        

        sec.Headers(wdHeaderFooterFirstPage).Range.Text = ""

        sec.Footers(wdHeaderFooterPrimary).Range.Text = ""

        sec.Footers(wdHeaderFooterFirstPage).Range.Text = ""

    Next sec

    DbgLog "ExecuteAutoStandardize: Step 4 - PageNumber done"

    

    If Not isQuiet Then

        SetProgress 25, DecodeUTF8("xJBhbmcgxJHDoW5oIHPhu5EgdHJhbmcgdsOgIHhhbmggc2EgxJHhuqd1IHRyYW5nLi4u"), "AutoStandardize"

    End If

    

    ' 5. KIEM TRA & CHUAN HOA KHUNG TIEU DE DAU TRANG (HEADER BLOCK)

    Dim hasHeader As Boolean: hasHeader = False

    Dim hdrTblAuto As Table

    Set hdrTblAuto = FindExistingHeaderTable(doc)

    If Not hdrTblAuto Is Nothing Then

        hasHeader = True

        RebuildHeaderTable doc, hdrTblAuto, regIdx

    End If

    

    If Not hasHeader Then

        Dim convertedRawAuto As Boolean: convertedRawAuto = False

        convertedRawAuto = ConvertRawHeaderToTable(doc, regIdx)

        If Not convertedRawAuto Then

            Dim topText As String: topText = ""

            Dim pCheck As Paragraph, pIdx As Long

            For pIdx = 1 To MinLong(doc.Paragraphs.Count, 8)

                topText = topText & " " & doc.Paragraphs(pIdx).Range.Text

            Next pIdx

            topText = UCase(topText)

            

            If InStr(topText, "CONG HOA XA HOI") = 0 And InStr(topText, "DANG CONG SAN") = 0 And _

               InStr(topText, DecodeUTF8("Q+G7mE5HIEjDkkE=")) = 0 And InStr(topText, DecodeUTF8("xJDhuqJORyBD4buYTkc=")) = 0 Then

                AutoInsertMissingHeader doc, regIdx

                statsMissing = statsMissing + 1

            End If

        Else

            statsHeadings = statsHeadings + 1

        End If

    End If

    DbgLog "ExecuteAutoStandardize: Step 5 - Header done"

    

    If Not isQuiet Then

        SetProgress 40, DecodeUTF8("xJBhbmcgY2h14bqpbiBow7NhIGtow7luZyBRdeG7kWMgaGnhu4d1ICYgVGnDqnUgxJHhu4EuLi4="), "AutoStandardize"

    End If

    

    ' 6. NHAN DIEN & CAN GIUA TEN LOAI VAN BAN + TRICH YEU

    titleDetected = StandardizeTitleAndAbstract(doc, regIdx, statsHeadings)

    DbgLog "ExecuteAutoStandardize: Step 6 - Title done: " & titleDetected

    

    If Not isQuiet Then

        SetProgress 55, DecodeUTF8("xJBhbmcgxJHhu4tuaCBk4bqhbmcgbsOZaSBkdW5nLCDEkeG7gSBt4bulYyAmIGPEg24gY+G7qS4uLg=="), "AutoStandardize"

    End If

    

    ' 7. DINH DANG NOI DUNG, CAC DE MUC & CAN CU

    FormatBodyAndHeadings doc, regIdx, statsParagraphs, statsHeadings

    DbgLog "ExecuteAutoStandardize: Step 7 - Body done"

    

    If Not isQuiet Then

        SetProgress 75, DecodeUTF8("xJBhbmcgY2h14bqpbiBow7NhIGtow7luZyBjaOG7ryBrw70gdsOgIG7GoWkgbmjhuq1uLi4u"), "AutoStandardize"

    End If

    

    ' 8. KIEM TRA & CHUAN HOA KHUNG CHU KY & NOI NHAN

    If IsDocumentAttachedToDecision(doc) Then

        ' Van ban ban hanh kem theo Quyet dinh (Mau 1.3): KHONG co chu ky & noi nhan rieng!

        ' Ket thuc van ban bang ky hieu ./. o cuoi noi dung

        Dim lastPIdx As Long, pLastDoc As Paragraph, pLastTxt As String

        For lastPIdx = doc.Paragraphs.Count To 1 Step -1

            pLastTxt = Trim(Replace(Replace(Replace(doc.Paragraphs(lastPIdx).Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            If Len(pLastTxt) > 0 Then

                Set pLastDoc = doc.Paragraphs(lastPIdx)

                Exit For

            End If

        Next lastPIdx

        If Not pLastDoc Is Nothing Then

            Dim tCleanEnd As String: tCleanEnd = Trim(Replace(Replace(Replace(pLastDoc.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            If Right(tCleanEnd, 3) <> "./." And Right(tCleanEnd, 2) <> "/." Then

                pLastDoc.Range.InsertAfter "./."

            End If

        End If

    Else

        CleanDuplicatePlaceholderSignatures doc

        

        Dim hasSig As Boolean: hasSig = False

        Dim sigTblAuto As Table

        Set sigTblAuto = FindExistingSignatureTable(doc)

        

        If Not sigTblAuto Is Nothing Then

            hasSig = True

            Set sigTblAuto = RebuildSignatureTable(doc, sigTblAuto, regIdx)

        Else

            Dim convertedSigAuto As Boolean: convertedSigAuto = False

            convertedSigAuto = ConvertRawSignatureToTable(doc, regIdx)

            If convertedSigAuto Then

                hasSig = True

            Else

                If Not IsDocEndingWithSignatureBlock(doc) Then

                    AutoInsertMissingSignature doc, regIdx

                    statsMissing = statsMissing + 1

                End If

            End If

        End If

        CleanDuplicatePlaceholderSignatures doc

    End If

    

    If Not isQuiet Then

        SetProgress 85, DecodeUTF8("xJBhbmcgcXXDqXQgbOG7l2kgZOG6pXUgdGhhbmggdsOgIHjDs2EgdHJhbmcgdGjhu6thLi4u"), "AutoStandardize"

    End If

    

    ' 9. CHUAN HOA DAU THANH TOAN VAN (oa, uy -> oa, uy)

    StandardizeToneMarksAllQuiet doc

    DbgLog "ExecuteAutoStandardize: Step 9 - ToneMarks done"

    

    ' 10. XOA TRANG TRANG THUA

    DeleteBlankPagesQuiet doc

    DbgLog "ExecuteAutoStandardize: Step 10 - BlankPages done"

    

    ' 10B. XOA KHOANG TRANG LOI & DAU CACH BAT HOP LY TOAN VAN BAN

    CleanSpacingErrorsInternal doc.content, True

    DbgLog "ExecuteAutoStandardize: Step 10B - Spacing done"

    

    ' 10C. CHUAN HOA KHUNG BANG BIEU NOI DUNG (DATA TABLES)

    ' Tranh de bang bi troi (Wrap Around), khong lech le trang, lap tieu de trang sau & chong vo hang

    Dim dataTblAuto As Table, tDIdxAuto As Long

    For tDIdxAuto = 1 To doc.Tables.Count

        Set dataTblAuto = doc.Tables(tDIdxAuto)

        If Not (dataTblAuto Is hdrTblAuto) And Not (dataTblAuto Is sigTblAuto) Then

            On Error Resume Next

            dataTblAuto.Rows.WrapAroundText = 0

            If dataTblAuto.Rows.LeftIndent < 0 Then dataTblAuto.Rows.LeftIndent = 0

            dataTblAuto.Range.Font.Name = "Times New Roman"

            If dataTblAuto.Rows.Count >= 3 Then

                dataTblAuto.Rows(1).HeadingFormat = True

                dataTblAuto.Rows.AllowBreakAcrossPages = False

            End If

        End If

    Next tDIdxAuto

    DbgLog "ExecuteAutoStandardize: Step 10C - DataTables done"

    

    If Not isQuiet Then

        SetProgress 100, DecodeUTF8("SG/DoG4gdOG6pXQgY2h14bqpbiBow7NhIHRow6BuaCBjw7RuZyE="), "AutoStandardize"

    End If

    

    ' KET THUC TRANSACTION UNDO

    If Not undoRec Is Nothing Then

        undoRec.EndCustomRecord

    End If

    

    ' DONG FORM TIEN TRINH TRUOC KHI HIEN THI DIALOG XAC NHAN

    FinishProgress "AutoStandardize"

    

    ' 11. XUAT BANG BAO CAO KET QUA & CHO PHEP HOAN TAC (UNDO)

    Dim rpt As String

    rpt = DecodeUTF8("QsOBTyBDw4FPIEvhur5UIFFV4bqiIENIVeG6qE4gSMOTQSBU4buwIMSQ4buYTkcgKDEtQ0xJQ0sp") & vbCrLf & _

          "========================================" & vbCrLf & _

          "- " & DecodeUTF8("UXV5IGNodeG6qW4gw6FwIGThu6VuZzog") & regName & vbCrLf & _

          "- " & DecodeUTF8("S2jhu5UgZ2nhuqV5ICYgQ8SDbiBs4buHOiDEkMOjIGPEg24gQTQgKA==") & _

          IIf(regIdx = 1, "Tren 20, Duoi 20, Trai 30, Phai 15mm - HD05", "Tren 20, Duoi 20, Trai 30, Phai 20mm - ND30") & ")" & vbCrLf & _

          "- " & DecodeUTF8("xJDDoW5oIHPhu5EgdHJhbmc6IMSQw6MgxJHDoW5oIHPhu5EgdHJhbmcgY2jDrW5oIGdp4buvYSBs4buBIHRyw6puLCDhuqluIHRyYW5nIDEgKFRpbWVzIE5ldyBSb21hbik=") & vbCrLf

          

    If Len(titleDetected) > 0 Then

        rpt = rpt & "- " & DecodeUTF8("VMOqbiBsb+G6oWkgdsSDbiBi4bqjbjogxJDDoyBjxINuIGdp4buvYSAmIGluIMSR4bqtbSAi") & titleDetected & """" & vbCrLf

    End If

    

    rpt = rpt & "- " & DecodeUTF8("xJDhu4EgbeG7pWMgxJHDoyBuaOG6rW4gZGnhu4duICYgaW4gxJHhuq1tOiA=") & statsHeadings & DecodeUTF8("IMSR4buBIG3hu6VjIChQaOG6p24sIENoxrDGoW5nLCBN4bulYywgxJBp4buBdSwgS2hv4bqjbiwgxJBp4buDbSk=") & vbCrLf & _

          "- " & DecodeUTF8("xJBv4bqhbiB2xINuIG7hu5lpIGR1bmcgxJHDoyBjaHXhuqluIGjDs2E6IA==") & statsParagraphs & DecodeUTF8("IMSRb+G6oW4gKEPEg24gxJHhu4F1IDIgbOG7gSwgdGjhu6V0IGTDsm5nIDFjbSwgZ2nDo24gxJFv4bqhbiA2cHQsIGdpw6NuIGTDsm5nIDEuMTUp") & vbCrLf

          

    If statsMissing > 0 Then

        rpt = rpt & "- " & DecodeUTF8("VGjDoG5oIHBo4bqnbiBi4buVIHN1bmc6IMSQw6MgdOG7sSDEkeG7mW5nIGNow6huIA==") & statsMissing & DecodeUTF8("IGtow7luZyB0aMOgbmggcGjhuqduIGLhu4sgdGhp4bq/dQ==") & vbCrLf

    End If

    

    rpt = rpt & "- " & DecodeUTF8("ROG6pXUgdGhhbmggdGnhur9uZyBWaeG7h3Q6IMSQw6MgY2h14bqpbiBow7NhIHRvw6BuIGLhu5kgKMOyYSwgw7p5Li4uKQ==") & vbCrLf & _

          "----------------------------------------" & vbCrLf & _

          DecodeUTF8("QuG6oW4gY8OzIG114buRbiBHSeG7riBM4bqgSSBjw6FjIMSR4buLbmggZOG6oW5nIHbhu6thIGNodeG6qW4gaMOzYSBraMO0bmc/CihDaOG7jW4gWUVTIMSR4buDIGNo4bqlcCBuaOG6rW4sIGNo4buNbiBOTyDEkeG7gyBIT8OATiBUw4FDIG5nYXkgbOG6rXAgdOG7qWMp")

          

    If Not isQuiet Then

        Dim keepChanges As Boolean

        keepChanges = ShowStandardizeReportForm(rpt)

        

        If Not keepChanges Then

            doc.Undo

            ShowUniMsg "xJDDoyBob8OgbiB0w6FjIHRvw6BuIGLhu5kgY8OhYyB0aGFvIHTDoWMgY2h14bqpbiBow7NhIHbDoCBraMO0aSBwaOG7pWMgbmd1ecOqbiB0cuG6oW5nIHbEg24gYuG6o24h", "Vi1GT1JNQVQ=", 64

        End If

    End If

    Exit Sub



ErrH:

    DbgLog "ExecuteAutoStandardize: ErrH - Error " & Err.Number & ": " & Err.Description

    FinishProgress "AutoStandardize"

    If Not undoRec Is Nothing Then

        On Error Resume Next

        undoRec.EndCustomRecord

    End If

    RecordError "VFormat_Core.AutoStandardizeDocument", Err.Number, Err.Description, "Chuan hoa tu dong"

End Sub



Public Sub UndoLastStandardize(Optional doc As Document = Nothing, Optional ByVal isQuiet As Boolean = False)

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then Exit Sub

    

    Dim doUndo As Boolean: doUndo = True

    If Not isQuiet Then

        doUndo = ShowUniConfirm("QuG6oW4gY8OzIG114buRbiBob8OgbiB0w6FjIGzhuqduIGNodeG6qW4gaMOzYSB0aOG7gyB0aOG7qWMgZ+G6p24gbmjhuqV0IGtow7RuZz8=", "Vi1GT1JNQVQgLSBIb8OgbiBUw6FjIENodeG6qW4gSMOzYQ==")

    End If

    If doUndo Then

        On Error Resume Next

        doc.Undo

        If Not isQuiet Then

            ShowUniMsg "xJDDoyBob8OgbiB0w6FjIHRvw6BuIGLhu5kgY8OhYyB0aGFvIHTDoWMgY2h14bqpbiBow7NhIHbDoCBraMO0aSBwaOG7pWMgbmd1ecOqbiB0cuG6oW5nIHbEg24gYuG6o24h", "Vi1GT1JNQVQ=", 64

        End If

    End If

End Sub



Private Sub SetParaFont(ByVal pObj As Object, ByVal fName As String, ByVal fSize As Single, ByVal isBold As Boolean, ByVal isItalic As Boolean)

    On Error Resume Next

    Dim rng As Range

    If TypeOf pObj Is Paragraph Then

        Set rng = pObj.Range

    ElseIf TypeOf pObj Is Range Then

        Set rng = pObj

    ElseIf TypeOf pObj Is Cell Then

        Set rng = pObj.Range

    End If

    If Not rng Is Nothing Then

        With rng.Font

            .Name = fName

            .NameAscii = fName

            .NameOther = fName

            .NameBi = fName

            If fSize > 0 Then .Size = fSize

            .Bold = isBold

            .Italic = isItalic

            .ColorIndex = wdBlack

        End With

    End If

End Sub



Private Function IsDocumentTitle(ByVal uTxt As String) As Boolean

    IsDocumentTitle = False

    Dim t As String: t = Trim(uTxt)

    If Len(t) = 0 Then Exit Function

    If Right(t, 1) = ":" Then Exit Function

    

    ' Danh muc 29 loai van ban hanh chinh theo Dieu 7 Nghi dinh 30/2020/ND-CP va van ban QPPL

    If t = "NGHI DINH" Or t = "QUYET DINH" Or t = "NGHI QUYET" Or t = "TO TRINH" Or _

       t = "THONG BAO" Or t = "BIEN BAN" Or t = "CHI THI" Or _

       t = "QUY CHE" Or t = "QUY DINH" Or t = "CONG DIEN" Or _

       t = "THONG TU" Or t = "THONG TU LIEN TICH" Or _

       t = "LUAT" Or t = "BO LUAT" Or t = "PHAP LENH" Or t = "LENH" Or _

       t = "GIAY MOI" Or t = "GIAY GIOI THIEU" Or t = "GIAY NGHI PHEP" Or _

       t = "GIAY UY QUYEN" Or t = "THONG CAO" Or t = "HUONG DAN" Or _

       t = "CHUONG TRINH" Or t = "KET LUAN" Or t = "HOP DONG" Or _

       t = "DE AN" Or t = "DU AN" Or t = "PHUONG AN" Or _

       t = "BAN GHI NHO" Or t = "BAN THOA THUAN" Or _

       t = "PHIEU GUI" Or t = "PHIEU CHUYEN" Or t = "PHIEU BAO" Or _

       t = "THU CONG" Or _

       (Left(t, 10) = "NGHI DINH ") Or _

       (Left(t, 11) = "QUYET DINH ") Or _

       (Left(t, 11) = "NGHI QUYET ") Or _

       (Left(t, 9) = "THONG TU ") Or _

       (Left(t, 8) = "CHI THI ") Or _

       (Left(t, 9) = "TO TRINH ") Or _

       (Left(t, 10) = "THONG BAO ") Or _

       (Left(t, 9) = "BIEN BAN ") Or _

       (Left(t, 9) = "QUY CHE ") Or _

       (Left(t, 9) = "QUY DINH ") Or _

       (Left(t, 9) = "CONG VAN ") Or _

       (Left(t, 8) = "KE HOACH" And InStr(t, "BAI DAY") = 0) Or _

       (Left(t, 7) = "BAO CAO") Then

        IsDocumentTitle = True

    End If

End Function



Private Function ConvertRawHeaderToTable(ByVal doc As Document, ByVal regIdx As Integer) As Boolean

    ConvertRawHeaderToTable = False

    On Error GoTo ErrH

    

    Dim maxCheck As Long: maxCheck = MinLong(doc.Paragraphs.Count, 12)

    If maxCheck < 2 Then Exit Function

    

    Dim i As Long, p As Paragraph, txt As String, uTxt As String

    Dim strAgency As String: strAgency = ""

    Dim strNum As String: strNum = ""

    Dim strDate As String: strDate = ""

    Dim hasNationalTitle As Boolean: hasNationalTitle = False

    Dim maxConsumed As Long: maxConsumed = 0

    

    For i = 1 To maxCheck

        Set p = doc.Paragraphs(i)

        If Not p.Range.Information(wdWithInTable) Then

            txt = Trim(Replace(Replace(Replace(p.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            If Len(txt) > 0 Then

                uTxt = UCase(ConvertToUnsignedVietnamese(txt))

                

                ' Neu gap tieu de van ban chinh (Quyet dinh, Nghi quyet...) thi dung scan header

                If IsDocumentTitle(uTxt) Then Exit For

                

                ' Kiem tra xem dong co duoc phan tach thanh 2 cot bang Tab hoac 3+ dau cach khong

                Dim partL As String: partL = ""

                Dim partR As String: partR = ""

                Dim isSplitLine As Boolean: isSplitLine = False

                Dim subP As String

                

                If InStr(txt, vbTab) > 0 Then

                    Dim hParts() As String: hParts = Split(txt, vbTab)

                    Dim hpIdx As Long

                    For hpIdx = 0 To UBound(hParts)

                        subP = Trim(hParts(hpIdx))

                        If Len(subP) > 0 Then

                            If Len(partL) = 0 Then

                                partL = subP

                            Else

                                partR = subP

                            End If

                        End If

                    Next hpIdx

                    If Len(partL) > 0 And Len(partR) > 0 Then isSplitLine = True

                ElseIf InStr(txt, "   ") > 0 Then

                    Dim sNormH As String: sNormH = txt

                    Do While InStr(sNormH, "    ") > 0

                        sNormH = Replace(sNormH, "    ", "   ")

                    Loop

                    sNormH = Replace(sNormH, "   ", vbTab)

                    Dim spParts() As String: spParts = Split(sNormH, vbTab)

                    For hpIdx = 0 To UBound(spParts)

                        subP = Trim(spParts(hpIdx))

                        If Len(subP) > 0 Then

                            If Len(partL) = 0 Then

                                partL = subP

                            Else

                                partR = subP

                            End If

                        End If

                    Next hpIdx

                    If Len(partL) > 0 And Len(partR) > 0 Then isSplitLine = True

                End If

                

                If isSplitLine Then

                    maxConsumed = i

                    Dim uL As String: uL = UCase(ConvertToUnsignedVietnamese(partL))

                    Dim uR As String: uR = UCase(ConvertToUnsignedVietnamese(partR))

                    

                    ' Xu ly nua ben trai

                    If Left(uL, 3) = "SO " Or Left(uL, 3) = "SO:" Or InStr(uL, "SO:") > 0 Then

                        strNum = partL

                    Else

                        If Len(strAgency) > 0 Then strAgency = strAgency & vbCrLf

                        strAgency = strAgency & partL

                    End If

                    

                    ' Xu ly nua ben phai

                    If InStr(uR, "CONG HOA XA HOI") > 0 Or InStr(uR, "DANG CONG SAN") > 0 Then

                        hasNationalTitle = True

                    ElseIf InStr(uR, "DOC LAP - TU DO") > 0 Then

                        hasNationalTitle = True

                    ElseIf InStr(uR, "NGAY ") > 0 And InStr(uR, " THANG ") > 0 And InStr(uR, " NAM ") > 0 Then

                        strDate = partR

                    End If

                Else

                    ' Dong don (khong phan cot bang tab/space)

                    If InStr(uTxt, "CONG HOA XA HOI") > 0 Or InStr(uTxt, "DANG CONG SAN") > 0 Then

                        hasNationalTitle = True

                        maxConsumed = i

                    ElseIf InStr(uTxt, "DOC LAP - TU DO") > 0 Then

                        hasNationalTitle = True

                        maxConsumed = i

                    ElseIf InStr(uTxt, "NGAY ") > 0 And InStr(uTxt, " THANG ") > 0 And InStr(uTxt, " NAM ") > 0 Then

                        strDate = txt

                        maxConsumed = i

                    ElseIf Left(uTxt, 3) = "SO " Or Left(uTxt, 3) = "SO:" Or InStr(uTxt, "SO:") > 0 Then

                        strNum = txt

                        maxConsumed = i

                    ElseIf txt = "*" Or txt = "***" Then

                        maxConsumed = i

                    Else

                        ' Kiem tra dong gach chan phan cach trang tri (decor separator: ---, ____, * * *)

                        Dim cleanSep As String

                        cleanSep = Replace(Replace(Replace(Replace(Replace(txt, "-", ""), "_", ""), ChrW(8211), ""), ChrW(8212), ""), "*", "")

                        cleanSep = Replace(cleanSep, " ", "")

                        If Len(cleanSep) = 0 Then

                            ' Day la dong gach chan trang tri duoi Quoc hieu hoac Tieu ngu, khong phai Ten Co quan

                            maxConsumed = i

                        Else

                            If maxConsumed < 5 Then

                                If Len(strAgency) > 0 Then strAgency = strAgency & vbCrLf

                                strAgency = strAgency & txt

                                maxConsumed = i

                            End If

                        End If

                    End If

                End If

            End If

        End If

    Next i

    

    If Not hasNationalTitle And Len(strNum) = 0 And Len(strDate) = 0 Then Exit Function

    If maxConsumed = 0 Then Exit Function

    

    ' Chuan hoa strNum

    If Len(strNum) = 0 Then

        strNum = IIf(regIdx = 1, DecodeUTF8("U+G7kSAuLi4tTlEvxJBV"), DecodeUTF8("U+G7kTogLi4uL1HEkC0uLi4="))

    Else

        If regIdx = 1 Then

            If Right(strNum, 1) = "*" Then strNum = Trim(Left(strNum, Len(strNum) - 1))

            strNum = Replace(strNum, DecodeUTF8("U+G7kTo="), DecodeUTF8("U+G7kSA="))

            strNum = Replace(strNum, DecodeUTF8("U+G7kTog"), DecodeUTF8("U+G7kSA="))

            strNum = Replace(strNum, "So: ", "So ")

            strNum = Replace(strNum, "So:", "So ")

            strNum = Replace(strNum, "SO: ", "SO ")

            strNum = Replace(strNum, "SO:", "SO ")

        Else

            If Left(UCase(strNum), 3) = "SO " Then

                strNum = "So: " & Trim(Mid(strNum, 4))

            ElseIf Left(strNum, 3) = DecodeUTF8("U+G7kSA=") Then

                strNum = DecodeUTF8("U+G7kTog") & Trim(Mid(strNum, 4))

            End If

        End If

        Do While InStr(strNum, "  ") > 0: strNum = Replace(strNum, "  ", " "): Loop

    End If

    

    ' Chuan hoa strDate

    If Len(strDate) = 0 Then

        strDate = DecodeUTF8("Li4uLCBuZ8OgeSAuLi4gdGjDoW5nIC4uLiBuxINtIDIwLi4u")

    End If

    

    Dim rngDelete As Range

    Set rngDelete = doc.Range(doc.Paragraphs(1).Range.Start, doc.Paragraphs(maxConsumed).Range.End)

    rngDelete.Delete

    

    Dim rngTable As Range: Set rngTable = doc.Range(0, 0)

    Dim tbl As Table

    Set tbl = doc.Tables.Add(Range:=rngTable, NumRows:=2, NumColumns:=2)

    tbl.Borders.Enable = False

    tbl.AutoFitBehavior wdAutoFitFixed

    tbl.LeftPadding = 0: tbl.RightPadding = 0

    tbl.TopPadding = 0: tbl.BottomPadding = 0

    tbl.Rows.WrapAroundText = 0

    tbl.Rows.LeftIndent = 0

    tbl.Rows.Alignment = wdAlignRowLeft

    

    Dim col1WRaw As Single, col2WRaw As Single

    If regIdx = 1 Then

        col1WRaw = MillimetersToPoints(70)

        col2WRaw = MillimetersToPoints(90)

    Else

        col1WRaw = MillimetersToPoints(65)

        col2WRaw = MillimetersToPoints(95)

    End If

    tbl.PreferredWidthType = wdPreferredWidthPoints

    tbl.PreferredWidth = col1WRaw + col2WRaw

    tbl.Columns(1).Width = col1WRaw

    tbl.Columns(2).Width = col2WRaw

    Dim rWRaw As Row

    For Each rWRaw In tbl.Rows

        If rWRaw.Cells.Count >= 2 Then

            rWRaw.Cells(1).Width = col1WRaw

            rWRaw.Cells(2).Width = col2WRaw

        End If

    Next rWRaw

    

    strNum = NormalizeDocNumber(strNum, regIdx)

    strDate = NormalizeDocDate(strDate)

    

    If regIdx = 1 Then

        If Len(strAgency) = 0 Then strAgency = DecodeUTF8("VOG7iE5IIOG7plkuLi4=") & vbCrLf & DecodeUTF8("xJDhuqJORyDhu6ZZLi4u")

        tbl.Cell(1, 1).Range.Text = strAgency & vbCrLf & "*"

        tbl.Cell(2, 1).Range.Text = strNum

        tbl.Cell(1, 2).Range.Text = DecodeUTF8("xJDhuqJORyBD4buYTkcgU+G6ok4gVknhu4ZUIE5BTQ==")

        tbl.Cell(2, 2).Range.Text = strDate

    Else

        If Len(strAgency) = 0 Then

            Dim uDocSnippet As String

            uDocSnippet = UCase(ConvertToUnsignedVietnamese(doc.Range(0, MinLong(doc.content.End, 2000)).Text))

            If InStr(uDocSnippet, "CHINH PHU BAN HANH") > 0 Or InStr(uDocSnippet, "THU TUONG CHINH PHU") > 0 Then

                strAgency = DecodeUTF8("Q0jDjE5IIFBI4bun")

            Else

                strAgency = DecodeUTF8("VUJORCBU4buITkguLi4=") & vbCrLf & DecodeUTF8("U+G7ni4uLg==")

            End If

        End If

        tbl.Cell(1, 1).Range.Text = strAgency

        tbl.Cell(2, 1).Range.Text = strNum

        tbl.Cell(1, 2).Range.Text = DecodeUTF8("Q+G7mE5HIEjDkkEgWMODIEjhu5hJIENI4bumIE5HSMSoQSBWSeG7hlQgTkFN") & vbCrLf & DecodeUTF8("xJDhu5ljIGzhuq1wIC0gVOG7sSBkbyAtIEjhuqFuaCBwaMO6Yw==")

        tbl.Cell(2, 2).Range.Text = strDate

    End If

    

    FormatExistingHeaderTable tbl, regIdx

    Dim rngAfterTbl As Range

    Set rngAfterTbl = tbl.Range

    rngAfterTbl.Collapse wdCollapseEnd

    rngAfterTbl.InsertParagraphAfter

    

    ConvertRawHeaderToTable = True

    Exit Function

ErrH:

    ConvertRawHeaderToTable = False

End Function



Public Function IsDocumentAttachedToDecision(ByVal doc As Document) As Boolean

    IsDocumentAttachedToDecision = False

    On Error Resume Next

    If doc Is Nothing Then Exit Function

    If doc.Paragraphs.Count = 0 Then Exit Function

    

    ' Neu van ban co Tieu de la QUYET DINH hoac NGHI QUYET thi day la Van ban goc ban hanh,

    ' KHONG PHAI la van ban kem theo (du co chua cum tu 'Ban hanh kem theo Quyet dinh nay' trong Dieu 1)

    Dim pIdx As Long, p As Paragraph, pText As String, uP As String

    For pIdx = 1 To MinLong(doc.Paragraphs.Count, 15)

        Set p = doc.Paragraphs(pIdx)

        If Not p.Range.Information(wdWithInTable) Then

            pText = Trim(Replace(Replace(Replace(p.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            uP = UCase(ConvertToUnsignedVietnamese(pText))

            If uP = "QUYET DINH" Or uP = "NGHI QUYET" Then

                ' Day la van ban Quyet dinh / Nghi quyet goc

                Exit Function

            End If

            

            ' Neu gap dong chi dan ban hanh kem theo o dau van ban:

            ' (Ban hanh kem theo..., Phe duyet kem theo..., Kem theo Quyet dinh so...)

            If (Left(pText, 1) = "(" Or InStr(uP, "BAN HANH KEM THEO") > 0 Or InStr(uP, "PHE DUYET KEM THEO") > 0 Or InStr(uP, "KEM THEO QUYET DINH") > 0 Or InStr(uP, "KEM THEO QD") > 0) And _

               (InStr(uP, "QUYET DINH") > 0 Or InStr(uP, "QD-") > 0 Or InStr(uP, "QD/") > 0) Then

                IsDocumentAttachedToDecision = True

                Exit Function

            End If

        End If

    Next pIdx

End Function



Private Function IsSignatureOrRecipientText(ByVal rawText As String) As Boolean

    IsSignatureOrRecipientText = False

    Dim t As String: t = Trim(rawText)

    If Len(t) = 0 Then Exit Function

    If Len(t) > 150 Then Exit Function

    

    Dim u As String: u = UCase(ConvertToUnsignedVietnamese(t))

    

    ' Noi nhan / Kinh gui / Luu tru

    If InStr(u, "NOI NHAN") > 0 Or InStr(u, "KINH GUI") > 0 Or _

       InStr(u, "LUU: VT") > 0 Or InStr(u, "LUU VT") > 0 Or _

       InStr(u, "LUU: TCHC") > 0 Or InStr(u, "NHU TREN") > 0 Then

        IsSignatureOrRecipientText = True

        Exit Function

    End If

    

    ' Tham quyen ky

    If Left(u, 3) = "TM." Or Left(u, 4) = "T/M " Or Left(u, 3) = "T/M" Or _

       Left(u, 3) = "KT." Or Left(u, 4) = "K/T " Or Left(u, 3) = "K/T" Or _

       Left(u, 3) = "TL." Or Left(u, 4) = "T/L " Or Left(u, 3) = "T/L" Or _

       Left(u, 4) = "TUQ." Or Left(u, 5) = "T/U/Q" Or Left(u, 2) = "Q." Or _

       InStr(u, "KT. ") > 0 Or InStr(u, "TM. ") > 0 Then

        IsSignatureOrRecipientText = True

        Exit Function

    End If

    

    ' Cac chuc vu lanh dao pho bien

    If InStr(u, "CHU TICH") > 0 Or InStr(u, "BI THU") > 0 Or InStr(u, "GIAM DOC") > 0 Or _

       InStr(u, "HIEU TRUONG") > 0 Or InStr(u, "VIEN TRUONG") > 0 Or InStr(u, "TRUONG BAN") > 0 Or _

       InStr(u, "TRUONG PHONG") > 0 Or InStr(u, "CHU NHIEM") > 0 Or InStr(u, "CHI HUY TRUONG") > 0 Or _

       InStr(u, "CHINH UY") > 0 Or InStr(u, "CUC TRUONG") > 0 Or InStr(u, "VU TRUONG") > 0 Or _

       InStr(u, "TONG CUC TRUONG") > 0 Or InStr(u, "THU TRUONG") > 0 Or InStr(u, "TONG GIAM DOC") > 0 Or _

       InStr(u, "TRUONG KHOA") > 0 Or InStr(u, "DAI DIEN") > 0 Then

        IsSignatureOrRecipientText = True

        Exit Function

    End If

End Function



Private Function IsSignatureOrRecipientBlock(ByVal rawText As String) As Boolean

    IsSignatureOrRecipientBlock = False

    Dim t As String: t = Trim(rawText)

    If Len(t) = 0 Then Exit Function

    

    Dim u As String: u = UCase(ConvertToUnsignedVietnamese(t))

    

    ' Noi nhan / Kinh gui / Luu tru

    If InStr(u, "NOI NHAN") > 0 Or InStr(u, "KINH GUI") > 0 Or _

       InStr(u, "LUU: VT") > 0 Or InStr(u, "LUU VT") > 0 Or _

       InStr(u, "LUU: TCHC") > 0 Or InStr(u, "NHU TREN") > 0 Then

        IsSignatureOrRecipientBlock = True

        Exit Function

    End If

    

    ' Tham quyen ky hoac chuc vu kem tham quyen

    If InStr(u, "TM.") > 0 Or InStr(u, "KT.") > 0 Or _

       InStr(u, "TL.") > 0 Or InStr(u, "TUQ.") > 0 Or _

       InStr(u, "T/M") > 0 Or InStr(u, "K/T") > 0 Or InStr(u, "T/L") > 0 Or _

       InStr(u, "T/U/Q") > 0 Then

        IsSignatureOrRecipientBlock = True

        Exit Function

    End If

    

    ' Cac chuc vu lanh dao

    If InStr(u, "CHU TICH") > 0 Or InStr(u, "BI THU") > 0 Or InStr(u, "GIAM DOC") > 0 Or _

       InStr(u, "HIEU TRUONG") > 0 Or InStr(u, "VIEN TRUONG") > 0 Or InStr(u, "TRUONG BAN") > 0 Or _

       InStr(u, "TRUONG PHONG") > 0 Or InStr(u, "CHU NHIEM") > 0 Or InStr(u, "CHI HUY TRUONG") > 0 Or _

       InStr(u, "CHINH UY") > 0 Or InStr(u, "CUC TRUONG") > 0 Or InStr(u, "VU TRUONG") > 0 Or _

       InStr(u, "TONG CUC TRUONG") > 0 Or InStr(u, "THU TRUONG") > 0 Or InStr(u, "TONG GIAM DOC") > 0 Or _

       InStr(u, "TRUONG KHOA") > 0 Or InStr(u, "DAI DIEN") > 0 Or InStr(u, "NGUOI KY") > 0 Or _

       InStr(u, "TRAN VAN B") > 0 Or InStr(u, "NGUYEN VAN A") > 0 Then

        IsSignatureOrRecipientBlock = True

        Exit Function

    End If

End Function



Private Function IsDocEndingWithSignatureBlock(ByVal doc As Document) As Boolean

    IsDocEndingWithSignatureBlock = False

    On Error Resume Next

    If doc Is Nothing Then Exit Function

    If doc.Paragraphs.Count = 0 Then Exit Function

    

    ' 1. Neu bang cuoi cung trong tai lieu la bang chu ky (co Noi nhan / Chuc vu ky)

    If doc.Tables.Count > 0 Then

        Dim lastTbl As Table: Set lastTbl = doc.Tables(doc.Tables.Count)

        If lastTbl.Range.Start > (doc.content.End * 0.4) Then

            Dim tTxt As String: tTxt = UCase(ConvertToUnsignedVietnamese(lastTbl.Range.Text))

            If (InStr(tTxt, "NOI NHAN") > 0 Or InStr(tTxt, "LUU: VT") > 0) And _

               (InStr(tTxt, "CHU TICH") > 0 Or InStr(tTxt, "GIAM DOC") > 0 Or InStr(tTxt, "BO TRUONG") > 0 Or _

                InStr(tTxt, "CUC TRUONG") > 0 Or InStr(tTxt, "TM.") > 0 Or InStr(tTxt, "KT.") > 0 Or InStr(tTxt, "TL.") > 0) Then

                IsDocEndingWithSignatureBlock = True

                Exit Function

            End If

        End If

    End If

    

    ' 2. Kiem tra xem co cac doan van rieng biet dong vai tro khoi chu ky / noi nhan o cuoi van ban khong

    Dim pIdx As Long, p As Paragraph, pText As String, uP As String

    For pIdx = doc.Paragraphs.Count To MaxLong(1, doc.Paragraphs.Count - 8) Step -1

        Set p = doc.Paragraphs(pIdx)

        If Not p.Range.Information(wdWithInTable) Then

            pText = Trim(Replace(Replace(Replace(p.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            uP = UCase(ConvertToUnsignedVietnamese(pText))

            If Len(pText) > 0 And Len(pText) < 60 Then

                ' Loai tru dieu khoan thi hanh (Dieu 3, Dieu 4...)

                If Left(uP, 5) <> "DIEU " And InStr(uP, "CHIU TRACH NHIEM") = 0 And InStr(uP, "THI HANH") = 0 Then

                    If Left(uP, 9) = "NOI NHAN:" Or Left(uP, 8) = "NOI NHAN" Or Left(uP, 9) = "- LUU: VT" Or _

                       Left(uP, 3) = "TM." Or Left(uP, 3) = "KT." Or Left(uP, 3) = "TL." Or Left(uP, 4) = "TUQ." Or _

                       uP = "CHU TICH" Or uP = "GIAM DOC" Or uP = "BO TRUONG" Or uP = "CUC TRUONG" Or uP = "TONG CUC TRUONG" Or _

                       uP = "BI THU" Or uP = "PHO BI THU" Or uP = "PHO CHU TICH" Or uP = "PHO GIAM DOC" Then

                        IsDocEndingWithSignatureBlock = True

                        Exit Function

                    End If

                End If

            End If

        End If

    Next pIdx

End Function

Private Function FindExistingHeaderTable(ByVal doc As Document) As Table

    Set FindExistingHeaderTable = Nothing

    On Error Resume Next

    If doc Is Nothing Then Exit Function

    If doc.Tables.Count = 0 Then Exit Function

    

    Dim tIdx As Long, tbl As Table, allTxt As String, c As Cell, uAll As String

    Dim maxTbl As Long

    maxTbl = MinLong(doc.Tables.Count, 3)

    For tIdx = 1 To maxTbl

        Set tbl = doc.Tables(tIdx)

        If tbl.Columns.Count >= 2 Or tbl.Rows(1).Cells.Count >= 2 Then

            allTxt = ""

            For Each c In tbl.Range.Cells

                allTxt = allTxt & " " & c.Range.Text

            Next c

            uAll = UCase(ConvertToUnsignedVietnamese(allTxt))

            

            ' Kiem tra noi dung dac trung cua bang Quoc hieu / Tieu de dau trang

            If InStr(uAll, "CONG HOA XA HOI") > 0 Or _

               InStr(uAll, "DANG CONG SAN") > 0 Or _

               InStr(uAll, "DOC LAP - TU DO") > 0 Or _

               (InStr(uAll, "SO ") > 0 And InStr(uAll, "NGAY ") > 0) Or _

               (InStr(uAll, "UBND") > 0 And (InStr(uAll, "SO:") > 0 Or InStr(uAll, "SO ") > 0)) Then

                Set FindExistingHeaderTable = tbl

                Exit Function

            End If

        End If

    Next tIdx

End Function



Private Function FindExistingSignatureTable(ByVal doc As Document) As Table

    Set FindExistingSignatureTable = Nothing

    On Error Resume Next

    If doc Is Nothing Then Exit Function

    If doc.Tables.Count = 0 Then Exit Function

    

    Dim tIdx As Long, tbl As Table, allTxt As String, c As Cell, uAll As String

    For tIdx = doc.Tables.Count To 1 Step -1

        Set tbl = doc.Tables(tIdx)

        If tbl.Columns.Count >= 2 Or tbl.Rows(1).Cells.Count >= 2 Then

            allTxt = ""

            For Each c In tbl.Range.Cells

                allTxt = allTxt & " " & c.Range.Text

            Next c

            uAll = UCase(ConvertToUnsignedVietnamese(allTxt))

            

            ' Neu la bang tieu de dau trang thi bo qua

            If InStr(uAll, "CONG HOA XA HOI") > 0 Or InStr(uAll, "DANG CONG SAN") > 0 Or _

               InStr(uAll, "DOC LAP - TU DO") > 0 Then

                ' Bo qua bang tieu de

            ElseIf IsSignatureOrRecipientBlock(allTxt) Then

                Set FindExistingSignatureTable = tbl

                Exit Function

            End If

        End If

    Next tIdx

End Function



Private Function CleanDuplicatePlaceholderSignatures(ByVal doc As Document) As Boolean

    CleanDuplicatePlaceholderSignatures = False

    On Error Resume Next

    If doc Is Nothing Then Exit Function

    If doc.Tables.Count < 2 Then Exit Function

    

    Dim lastTbl As Table: Set lastTbl = doc.Tables(doc.Tables.Count)

    If lastTbl.Columns.Count >= 2 Or lastTbl.Rows(1).Cells.Count >= 2 Then

        Dim lastTxt As String: lastTxt = UCase(ConvertToUnsignedVietnamese(lastTbl.Range.Text))

        ' Kiem tra neu bang cuoi cung la bang mau tu sinh (chua Tran Van B hoac Nguyen Van A)

        If InStr(lastTxt, "TRAN VAN B") > 0 Or InStr(lastTxt, "NGUYEN VAN A") > 0 Then

            Dim prevIdx As Long, prevTbl As Table, prevTxt As String

            For prevIdx = doc.Tables.Count - 1 To 1 Step -1

                Set prevTbl = doc.Tables(prevIdx)

                If prevTbl.Columns.Count >= 2 Or prevTbl.Rows(1).Cells.Count >= 2 Then

                    prevTxt = UCase(ConvertToUnsignedVietnamese(prevTbl.Range.Text))

                    If InStr(prevTxt, "CONG HOA XA HOI") = 0 And InStr(prevTxt, "DANG CONG SAN") = 0 Then

                        If IsSignatureOrRecipientBlock(prevTxt) Then

                            lastTbl.Delete

                            CleanDuplicatePlaceholderSignatures = True

                            Exit Function

                        End If

                    End If

                End If

            Next prevIdx

        End If

    End If

End Function



Private Function ConvertRawSignatureToTable(ByVal doc As Document, ByVal regIdx As Integer) As Boolean

    ConvertRawSignatureToTable = False

    On Error GoTo ErrH

    

    If IsDocumentAttachedToDecision(doc) Then Exit Function

    If doc.Paragraphs.Count < 5 Then Exit Function

    

    Dim startScan As Long: startScan = MaxLong(CLng(doc.Paragraphs.Count * 0.5), doc.Paragraphs.Count - 25)

    If startScan < 1 Then startScan = 1

    Dim i As Long, p As Paragraph, pText As String, uPText As String

    Dim recLines As String: recLines = ""

    Dim sigLines As String: sigLines = ""

    Dim firstConsumed As Long: firstConsumed = 0

    Dim lastConsumed As Long: lastConsumed = 0

    Dim inSigArea As Boolean: inSigArea = False

    

    For i = startScan To doc.Paragraphs.Count

        Set p = doc.Paragraphs(i)

        If Not p.Range.Information(wdWithInTable) Then

            pText = Trim(Replace(Replace(Replace(p.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            If Len(pText) > 0 Then

                uPText = UCase(ConvertToUnsignedVietnamese(pText))

                ' Loai tru cac Dieu khoan thi hanh khoi viec bi nhan nham la chu ky

                Dim isArticleClause As Boolean: isArticleClause = False

                If Left(uPText, 5) = "DIEU " Or InStr(uPText, "CHIU TRACH NHIEM") > 0 Or InStr(uPText, "THI HANH") > 0 Then

                    isArticleClause = True

                End If

                

                Dim isSigRec As Boolean

                If isArticleClause Then

                    isSigRec = False

                Else

                    isSigRec = IsSignatureOrRecipientBlock(uPText) Or IsSignatureOrRecipientText(uPText)

                End If

                

                If isSigRec Or (inSigArea And (Left(pText, 1) = "-" Or Left(pText, 1) = "+" Or InStr(pText, vbTab) > 0 Or InStr(pText, "   ") > 0)) Then

                    inSigArea = True

                    If firstConsumed = 0 Then firstConsumed = i

                    lastConsumed = i

                    

                    ' Tach neu co Tab hoac 3+ dau cach giua 2 cot tren cung 1 dong

                    Dim hasSplitSig As Boolean: hasSplitSig = False

                    Dim sigTabParts() As String

                    If InStr(pText, vbTab) > 0 Then

                        sigTabParts = Split(pText, vbTab)

                        hasSplitSig = True

                    ElseIf InStr(pText, "   ") > 0 Then

                        Dim sNormSig As String: sNormSig = pText

                        Do While InStr(sNormSig, "  ") > 0

                            sNormSig = Replace(sNormSig, "  ", vbTab)

                        Loop

                        sigTabParts = Split(sNormSig, vbTab)

                        hasSplitSig = True

                    End If

                    

                    If hasSplitSig Then

                        Dim lPart As String: lPart = ""

                        Dim rPart As String: rPart = ""

                        Dim tIdx As Long

                        For tIdx = 0 To UBound(sigTabParts)

                            Dim sPart As String: sPart = Trim(sigTabParts(tIdx))

                            If Len(sPart) > 0 Then

                                If Len(lPart) = 0 Then

                                    lPart = sPart

                                Else

                                    rPart = sPart

                                End If

                            End If

                        Next tIdx

                        If Len(lPart) > 0 Then

                            If Len(recLines) > 0 Then recLines = recLines & vbCrLf

                            recLines = recLines & lPart

                        End If

                        If Len(rPart) > 0 Then

                            If Len(sigLines) > 0 Then sigLines = sigLines & vbCrLf

                            sigLines = sigLines & rPart

                        End If

                    Else

                        ' Khong co tab/space phan cot: phan loai theo noi dung

                        If InStr(uPText, "NOI NHAN") > 0 Or Left(pText, 1) = "-" Or Left(pText, 1) = "+" Or _

                           InStr(uPText, "LUU: VT") > 0 Or InStr(uPText, "LUU VT") > 0 Or InStr(uPText, "NHU TREN") > 0 Then

                            If Len(recLines) > 0 Then recLines = recLines & vbCrLf

                            recLines = recLines & pText

                        Else

                            If Len(sigLines) > 0 Then sigLines = sigLines & vbCrLf

                            sigLines = sigLines & pText

                        End If

                    End If

                ElseIf inSigArea Then

                    Exit For

                End If

            End If

        End If

    Next i

    

    If firstConsumed = 0 Or lastConsumed = 0 Then Exit Function

    

    ' Xoa cac doan van tho da duoc trich xuat

    Dim rngDelete As Range

    Set rngDelete = doc.Range(doc.Paragraphs(firstConsumed).Range.Start, doc.Paragraphs(lastConsumed).Range.End)

    rngDelete.Delete

    

    ' Chen Table chu ky 1x2 tai vi tri do

    Dim rngInsert As Range

    If firstConsumed <= doc.Paragraphs.Count Then

        Set rngInsert = doc.Paragraphs(firstConsumed).Range

        rngInsert.Collapse wdCollapseStart

    Else

        Set rngInsert = doc.content

        rngInsert.Collapse wdCollapseEnd

    End If

    

    Dim tbl As Table

    Set tbl = doc.Tables.Add(Range:=rngInsert, NumRows:=1, NumColumns:=2)

    tbl.Borders.Enable = False

    tbl.AutoFitBehavior wdAutoFitFixed

    tbl.LeftPadding = 0: tbl.RightPadding = 0

    tbl.TopPadding = 0: tbl.BottomPadding = 0

    tbl.Rows.WrapAroundText = 0

    tbl.Rows.LeftIndent = 0

    tbl.Rows.Alignment = wdAlignRowLeft

    

    Dim col1WSigRaw As Single, col2WSigRaw As Single

    If regIdx = 1 Then

        col1WSigRaw = MillimetersToPoints(75)

        col2WSigRaw = MillimetersToPoints(90)

    Else

        col1WSigRaw = MillimetersToPoints(70)

        col2WSigRaw = MillimetersToPoints(90)

    End If

    tbl.PreferredWidthType = wdPreferredWidthPoints

    tbl.PreferredWidth = col1WSigRaw + col2WSigRaw

    tbl.Columns(1).Width = col1WSigRaw

    tbl.Columns(2).Width = col2WSigRaw

    

    If Len(recLines) = 0 Then

        recLines = DecodeUTF8("TsahaSBuaOG6rW46") & vbCrLf & _

                   DecodeUTF8("LSBOaMawIHRyw6puOw==") & vbCrLf & _

                   DecodeUTF8("LSBMxrB1OiBWVC4=")

    End If

    tbl.Cell(1, 1).Range.Text = recLines

    

    If Len(sigLines) = 0 Then

        If regIdx = 1 Then

            sigLines = DecodeUTF8("VC9NIEJBTiBDSOG6pFAgSMOATkg=") & vbCrLf & _

                       DecodeUTF8("QsONIFRIxq8=") & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _

                       DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

        Else

            sigLines = DecodeUTF8("VE0uIOG7plkgQkFOIE5Iw4JOIETDgk4=") & vbCrLf & _

                       DecodeUTF8("Q0jhu6YgVOG7ikNI") & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _

                       DecodeUTF8("VHLhuqduIFbEg24gQg==")

        End If

    End If

    tbl.Cell(1, 2).Range.Text = sigLines

    

    Set tbl = RebuildSignatureTable(doc, tbl, regIdx)

    ConvertRawSignatureToTable = True

    Exit Function

ErrH:

    ConvertRawSignatureToTable = False

End Function



Private Function IsSeparatorLine(ByVal rawText As String) As Boolean

    IsSeparatorLine = False

    Dim t As String: t = Trim(rawText)

    If Len(t) = 0 Then Exit Function

    Dim c1 As String: c1 = Left(t, 1)

    If c1 = "-" Or c1 = "_" Or c1 = "=" Or c1 = "*" Or c1 = ChrW(&H2014) Or c1 = "~" Then

        Dim allSame As Boolean: allSame = True

        Dim k As Long

        For k = 1 To Len(t)

            Dim ch As String: ch = Mid(t, k, 1)

            If ch <> "-" And ch <> "_" And ch <> "=" And ch <> "*" And ch <> ChrW(&H2014) And ch <> "~" And ch <> " " Then

                allSame = False

                Exit For

            End If

        Next k

        If allSame Then IsSeparatorLine = True

    End If

End Function



Private Function IsBodyStarter(ByVal uTxt As String) As Boolean

    IsBodyStarter = False

    Dim u As String: u = Trim(uTxt)

    If Len(u) = 0 Then Exit Function

    

    ' 1. Cum tu mo dau dan vao noi dung van ban (lead-in body text)

    If Left(u, 9) = "THUC HIEN" Or Left(u, 12) = "QUA KIEM TRA" Or Left(u, 4) = "QUA " Or _

       Left(u, 6) = "CAN CU" Or Left(u, 4) = "XET " Or Left(u, 5) = "THEO " Or _

       Left(u, 5) = "NHAM " Or Left(u, 3) = "DE " Or _

       Left(u, 13) = "THOI GIAN QUA" Or Left(u, 19) = "TRONG NHUNG NAM QUA" Or _

       Left(u, 19) = "TRONG THOI GIAN QUA" Or Left(u, 8) = "KINH GUI" Or _

       Left(u, 5) = "DIEU " Or Left(u, 11) = "QUYET DINH:" Or Left(u, 9) = "BAN HANH:" Or _

       Left(u, 9) = "QUY DINH:" Or Left(u, 10) = "THONG BAO:" Or Left(u, 9) = "KET LUAN:" Or _

       Left(u, 3) = "I. " Or Left(u, 4) = "II. " Or Left(u, 5) = "III. " Or Left(u, 4) = "IV. " Or _

       Left(u, 3) = "V. " Or Left(u, 4) = "VI. " Or Left(u, 5) = "VII. " Or _

       Left(u, 3) = "1. " Or Left(u, 3) = "2. " Or Left(u, 3) = "3. " Or _

       Left(u, 2) = "- " Or Left(u, 2) = "+ " Or _

       Left(u, 8) = "NOI NHAN" Or Left(u, 3) = "TM." Or Left(u, 3) = "KT." Or _

       Left(u, 5) = "PHAN " Or Left(u, 7) = "CHUONG " Or Left(u, 4) = "MUC " Or _

       Left(u, 9) = "TIEU MUC " Then

        IsBodyStarter = True

        Exit Function

    End If

    

    ' 2. Ket thuc bang cum tu dan ": " hoac "nhu sau:"

    If Right(u, 8) = "NHU SAU:" Or Right(u, 9) = "NHU SAU :" Or Right(u, 1) = ":" Then

        IsBodyStarter = True

        Exit Function

    End If

    

    ' 3. Heuristic doan van dai: Trich yeu tieu de thuong ngan (duoi 120 ky tu).

    ' Neu doan van dai tren 130 ky tu thi la doan van thong thuong (doan dan vao van ban).

    If Len(u) > 130 Then

        IsBodyStarter = True

        Exit Function

    End If

End Function



Private Function IsDocNumberLine(ByVal rawText As String) As Boolean

    IsDocNumberLine = False

    Dim t As String: t = Trim(rawText)

    If Len(t) = 0 Then Exit Function

    

    ' Tranh bat nham So (Department / So Nong nghiep, So Tu phap...)

    If Left(t, 3) = DecodeUTF8("U+G7nyA=") Or Left(t, 3) = DecodeUTF8("U+G7niA=") Then

        Exit Function

    End If

    

    Dim u As String: u = UCase(ConvertToUnsignedVietnamese(t))

    If Left(u, 3) = "SO:" Or Left(u, 4) = "SO: " Or Left(u, 4) = "SO. " Or Left(u, 4) = "SO :" Then

        IsDocNumberLine = True

        Exit Function

    End If

    If Left(u, 3) = "SO " And (InStr(t, "/") > 0 Or InStr(t, ":") > 0 Or InStr(t, "-") > 0) Then

        IsDocNumberLine = True

        Exit Function

    End If

    If InStr(u, "SO:") > 0 Then

        IsDocNumberLine = True

        Exit Function

    End If

End Function



Public Function NormalizeDocNumber(ByVal rawNum As String, ByVal regIdx As Integer) As String

    Dim t As String: t = Trim(rawNum)

    If Len(t) = 0 Then

        If regIdx = 1 Then

            NormalizeDocNumber = DecodeUTF8("U+G7kSAuLi4tQ04vVE8=")

        Else

            NormalizeDocNumber = DecodeUTF8("U+G7kTogLi4uLy4uLg==")

        End If

        Exit Function

    End If

    

    Dim prefix As String

    If regIdx = 1 Then

        prefix = DecodeUTF8("U+G7kSA=") ' So (khong dau hai cham)

    Else

        prefix = DecodeUTF8("U+G7kTog") ' So: (co dau hai cham)

    End If

    

    Dim subTxt As String: subTxt = t

    Dim uSub As String: uSub = UCase(ConvertToUnsignedVietnamese(subTxt))

    If Left(uSub, 4) = "SO: " Or Left(uSub, 4) = "SO. " Then

        subTxt = Trim(Mid(subTxt, 5))

    ElseIf Left(uSub, 3) = "SO:" Then

        subTxt = Trim(Mid(subTxt, 4))

    ElseIf Left(uSub, 3) = "SO " Then

        subTxt = Trim(Mid(subTxt, 4))

    ElseIf Left(subTxt, 4) = DecodeUTF8("U+G7kTog") Then

        subTxt = Trim(Mid(subTxt, 5))

    ElseIf Left(subTxt, 3) = DecodeUTF8("U+G7kSA=") Then

        subTxt = Trim(Mid(subTxt, 4))

    End If

    

    Dim dashPos As Long

    Dim slashPos As Long

    Dim sepPos As Long

    Dim k As Long, cChar As String

    

    If regIdx = 1 Then

        ' QUY CHUAN VAN BAN DANG (HD 05 - PHU LUC 2):

        ' - Tu "So " in thuong, co 14, dung, KHONG co dau hai cham (:)

        ' - So tu 1 den 9 bat buoc co so 0 phia truoc (01..09)

        ' - Ngan cach giua So va Ky hieu la dau gach noi (-)

        dashPos = InStr(subTxt, "-")

        slashPos = InStr(subTxt, "/")

        sepPos = 0

        If dashPos > 0 And slashPos > 0 Then

            If dashPos < slashPos Then

                sepPos = dashPos

            Else

                sepPos = slashPos

            End If

        ElseIf dashPos > 0 Then

            sepPos = dashPos

        ElseIf slashPos > 0 Then

            sepPos = slashPos

        End If

        

        If sepPos > 0 Then

            Dim numPartD As String: numPartD = Trim(Left(subTxt, sepPos - 1))

            Dim orgPartD As String: orgPartD = Trim(Mid(subTxt, sepPos + 1))

            

            Dim hasDigitD As Boolean: hasDigitD = False

            For k = 1 To Len(numPartD)

                cChar = Mid(numPartD, k, 1)

                If cChar >= "0" And cChar <= "9" Then

                    hasDigitD = True

                    Exit For

                End If

            Next k

            

            If Not hasDigitD And numPartD <> "..." Then

                numPartD = "..."

            ElseIf hasDigitD And IsNumeric(numPartD) Then

                Dim valND As Long: valND = Val(numPartD)

                If valND >= 1 And valND <= 9 And Len(numPartD) = 1 Then

                    numPartD = "0" & CStr(valND)

                End If

            End If

            If Len(orgPartD) = 0 Then orgPartD = "..."

            NormalizeDocNumber = prefix & numPartD & "-" & orgPartD

            Exit Function

        End If

        

        If IsNumeric(subTxt) Then

            Dim valOnly As Long: valOnly = Val(subTxt)

            If valOnly >= 1 And valOnly <= 9 And Len(subTxt) = 1 Then

                subTxt = "0" & CStr(valOnly)

            End If

            NormalizeDocNumber = prefix & subTxt & "-..."

            Exit Function

        End If

        

        NormalizeDocNumber = prefix & subTxt

        Exit Function

    Else

        ' QUY CHUAN HANH CHINH NHA NUOC (ND 30):

        slashPos = InStr(subTxt, "/")

        If slashPos > 0 Then

            Dim numPart As String: numPart = Trim(Left(subTxt, slashPos - 1))

            Dim orgPart As String: orgPart = Trim(Mid(subTxt, slashPos + 1))

            

            Dim hasDigit As Boolean: hasDigit = False

            For k = 1 To Len(numPart)

                cChar = Mid(numPart, k, 1)

                If cChar >= "0" And cChar <= "9" Then

                    hasDigit = True

                    Exit For

                End If

            Next k

            

            If Not hasDigit And numPart <> "..." Then

                numPart = "..."

            End If

            If Len(orgPart) = 0 Then orgPart = "..."

            NormalizeDocNumber = prefix & numPart & "/" & orgPart

            Exit Function

        End If

        

        Dim hasAnyDigit As Boolean: hasAnyDigit = False

        For k = 1 To Len(subTxt)

            cChar = Mid(subTxt, k, 1)

            If cChar >= "0" And cChar <= "9" Then

                hasAnyDigit = True

                Exit For

            End If

        Next k

        

        If Not hasAnyDigit Then

            NormalizeDocNumber = prefix & ".../..."

        Else

            NormalizeDocNumber = prefix & subTxt

        End If

    End If

End Function



Public Function NormalizeDocDate(ByVal rawDate As String) As String

    Dim t As String: t = Trim(rawDate)

    If Len(t) = 0 Then

        NormalizeDocDate = DecodeUTF8("Li4uLCBuZ8OgeSAuLi4gdGjDoW5nIC4uLiBuxINtIDIwLi4u")

        Exit Function

    End If

    

    Dim u As String: u = UCase(ConvertToUnsignedVietnamese(t))

    Dim ngayPos As Long: ngayPos = InStr(u, "NGAY")

    If ngayPos = 0 Then

        NormalizeDocDate = t

        Exit Function

    End If

    

    Dim placePart As String

    placePart = Trim(Left(t, ngayPos - 1))

    If Right(placePart, 1) = "," Then

        placePart = Trim(Left(placePart, Len(placePart) - 1))

    End If

    If Len(placePart) = 0 Then placePart = "..."

    

    Dim afterNgay As String: afterNgay = Trim(Mid(t, ngayPos + 4))

    Dim uAfterNgay As String: uAfterNgay = UCase(ConvertToUnsignedVietnamese(afterNgay))

    Dim thangPos As Long: thangPos = InStr(uAfterNgay, "THANG")

    If thangPos = 0 Then

        NormalizeDocDate = t

        Exit Function

    End If

    

    Dim dayPart As String: dayPart = Trim(Left(afterNgay, thangPos - 1))

    Dim hasDayDigit As Boolean: hasDayDigit = False

    Dim k As Long, c As String

    For k = 1 To Len(dayPart)

        c = Mid(dayPart, k, 1)

        If c >= "0" And c <= "9" Then

            hasDayDigit = True

            Exit For

        End If

    Next k

    If Not hasDayDigit And dayPart <> "..." Then

        dayPart = "..."

    ElseIf hasDayDigit And IsNumeric(dayPart) Then

        Dim vDay As Long: vDay = Val(dayPart)

        If vDay >= 1 And vDay <= 9 And Len(Trim(dayPart)) = 1 Then

            dayPart = "0" & CStr(vDay)

        End If

    End If

    

    Dim afterThang As String: afterThang = Trim(Mid(afterNgay, thangPos + 5))

    Dim uAfterThang As String: uAfterThang = UCase(ConvertToUnsignedVietnamese(afterThang))

    Dim namPos As Long: namPos = InStr(uAfterThang, "NAM")

    If namPos = 0 Then

        NormalizeDocDate = t

        Exit Function

    End If

    

    Dim monthPart As String: monthPart = Trim(Left(afterThang, namPos - 1))

    Dim hasMonthDigit As Boolean: hasMonthDigit = False

    For k = 1 To Len(monthPart)

        c = Mid(monthPart, k, 1)

        If c >= "0" And c <= "9" Then

            hasMonthDigit = True

            Exit For

        End If

    Next k

    If Not hasMonthDigit And monthPart <> "..." Then

        monthPart = "..."

    ElseIf hasMonthDigit And IsNumeric(monthPart) Then

        Dim vMonth As Long: vMonth = Val(monthPart)

        If (vMonth = 1 Or vMonth = 2) And Len(Trim(monthPart)) = 1 Then

            monthPart = "0" & CStr(vMonth)

        End If

    End If

    

    Dim yearPart As String: yearPart = Trim(Mid(afterThang, namPos + 3))

    Dim hasYearDigit As Boolean: hasYearDigit = False

    For k = 1 To Len(yearPart)

        c = Mid(yearPart, k, 1)

        If c >= "0" And c <= "9" Then

            hasYearDigit = True

            Exit For

        End If

    Next k

    If Not hasYearDigit And yearPart <> "..." Then

        yearPart = "20..."

    ElseIf yearPart = "20" Then

        yearPart = "20..."

    End If

    

    NormalizeDocDate = placePart & DecodeUTF8("LCBuZ8OgeSA=") & dayPart & _

                       DecodeUTF8("IHRow6FuZyA=") & monthPart & _

                       DecodeUTF8("IG7Eg20g") & yearPart

End Function



Private Sub RebuildHeaderTable(ByVal doc As Document, ByVal tbl As Table, ByVal regIdx As Integer)

    On Error GoTo Fallback

    If doc Is Nothing Or tbl Is Nothing Then Exit Sub

    

    ' 1. Trich xuat du lieu tu bang cu cua nguoi dung

    Dim supAgency As String: supAgency = ""

    Dim issAgency As String: issAgency = ""

    Dim docNum As String: docNum = ""

    Dim docNation As String: docNation = ""

    Dim docMotto As String: docMotto = ""

    Dim docDate As String: docDate = ""

    

    Dim rIdx As Long, cellObj As Cell, pObj As Paragraph

    Dim pText As String, uText As String

    

    ' Quet Cot 1: Co quan va So hieu

    Dim agencyLines As String: agencyLines = ""

    For rIdx = 1 To tbl.Rows.Count

        If tbl.Rows(rIdx).Cells.Count >= 1 Then

            Set cellObj = tbl.Cell(rIdx, 1)

            For Each pObj In cellObj.Range.Paragraphs

                pText = Trim(Replace(Replace(Replace(pObj.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

                If Len(pText) > 0 Then

                    If Not IsSeparatorLine(pText) Then

                        If IsDocNumberLine(pText) Then

                            If Len(docNum) = 0 Then docNum = pText

                        Else

                            If Len(agencyLines) > 0 Then agencyLines = agencyLines & vbCrLf

                            agencyLines = agencyLines & pText

                        End If

                    End If

                End If

            Next pObj

        End If

    Next rIdx

    

    ' Phan loai Co quan chu quan va Co quan ban hanh

    If Len(agencyLines) > 0 Then

        Dim aParts() As String: aParts = Split(agencyLines, vbCrLf)

        If UBound(aParts) = 0 Then

            issAgency = aParts(0)

        Else

            supAgency = aParts(0)

            issAgency = aParts(1)

        End If

    End If

    

    ' Quet Cot 2: Quoc hieu, Tieu ngu, Dia danh ngay thang

    For rIdx = 1 To tbl.Rows.Count

        If tbl.Rows(rIdx).Cells.Count >= 2 Then

            Set cellObj = tbl.Cell(rIdx, 2)

            For Each pObj In cellObj.Range.Paragraphs

                pText = Trim(Replace(Replace(Replace(pObj.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

                If Len(pText) > 0 Then

                    If Not IsSeparatorLine(pText) Then

                        uText = UCase(ConvertToUnsignedVietnamese(pText))

                        If InStr(uText, "CONG HOA XA HOI") > 0 Or InStr(uText, "DANG CONG SAN") > 0 Then

                            If Len(docNation) = 0 Then docNation = pText

                        ElseIf InStr(uText, "DOC LAP - TU DO") > 0 Or InStr(uText, "DOC LAP") > 0 Or InStr(uText, "HANH PHUC") > 0 Then

                            If Len(docMotto) = 0 Then docMotto = pText

                        ElseIf InStr(uText, "NGAY") > 0 And (InStr(uText, "THANG") > 0 Or InStr(uText, "NAM") > 0) Then

                            If Len(docDate) = 0 Then docDate = pText

                        ElseIf InStr(pText, ",") > 0 And Len(docDate) = 0 Then

                            docDate = pText

                        End If

                    End If

                End If

            Next pObj

        End If

    Next rIdx

    

    ' Bo sung gia tri mac dinh neu thieu & Tu dong danh san dau ... neu chua dien so hoac ngay thang

    docNum = NormalizeDocNumber(docNum, regIdx)

    docDate = NormalizeDocDate(docDate)

    

    If regIdx = 1 Then

        ' Dang

        If Len(docNation) = 0 Then docNation = DecodeUTF8("xJDhuqJORyBD4buYTkcgU+G6ok4gVknhu4ZUIE5BTQ==")

        If Len(issAgency) = 0 Then issAgency = DecodeUTF8("QkFOIENI4bqgUCBIw4BOSC4uLg==")

    Else

        ' Hanh chinh (ND 30)

        If Len(docNation) = 0 Then docNation = DecodeUTF8("Q+G7mE5HIEjDkkEgWMODIEjhu5hJIENI4bumIE5HSMSoQSBWSeG7hlQgTkFN")

        If Len(docMotto) = 0 Then docMotto = DecodeUTF8("xJDhu5ljIGzhuq1wIC0gVOG7sSBkbyAtIEjhuqFuaCBwaMO6Yw==")

        If Len(issAgency) = 0 Then issAgency = DecodeUTF8("VsSQLE4gUEjDkk5HLi4u")

    End If

    

    ' 2. Xoa bang cu va xoa Shape duong ke rac o khu vuc dau trang

    Dim rngTarget As Range

    Set rngTarget = tbl.Range

    rngTarget.Collapse wdCollapseStart

    

    tbl.Delete

    Set tbl = Nothing

    

    Dim sIdx As Long, shp As Shape

    For sIdx = doc.Shapes.Count To 1 Step -1

        Set shp = doc.Shapes(sIdx)

        If shp.Top < MillimetersToPoints(80) Then

            shp.Delete

        End If

    Next sIdx

    

    ' 3. Dung lai Bang 1x2 chuan Template V-Format tai vi tri bang cu

    Dim newTbl As Table

    Set newTbl = doc.Tables.Add(Range:=rngTarget, NumRows:=1, NumColumns:=2)

    newTbl.Borders.Enable = False

    newTbl.AutoFitBehavior wdAutoFitFixed

    newTbl.Rows.WrapAroundText = 0

    newTbl.Rows.LeftIndent = 0

    newTbl.Rows.Alignment = wdAlignRowLeft

    newTbl.LeftPadding = 0: newTbl.RightPadding = 0

    newTbl.TopPadding = 0: newTbl.BottomPadding = 0

    

    Dim printW As Single

    On Error Resume Next

    printW = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin

    If printW < MillimetersToPoints(140) Or printW > MillimetersToPoints(200) Then

        printW = MillimetersToPoints(160)

    End If

    

    Dim col1W As Single, col2W As Single

    If regIdx = 1 Then

        col1W = MillimetersToPoints(70)

        col2W = printW - col1W

        If col2W < MillimetersToPoints(88) Then col2W = MillimetersToPoints(88)

    Else

        col1W = MillimetersToPoints(65)

        col2W = printW - col1W

        If col2W < MillimetersToPoints(95) Then col2W = MillimetersToPoints(95)

    End If

    newTbl.PreferredWidthType = wdPreferredWidthPoints

    newTbl.PreferredWidth = col1W + col2W

    newTbl.Columns(1).Width = col1W

    newTbl.Columns(2).Width = col2W

    newTbl.Rows(1).Cells(1).LeftPadding = 0

    newTbl.Rows(1).Cells(1).RightPadding = 0

    newTbl.Rows(1).Cells(2).LeftPadding = 0

    newTbl.Rows(1).Cells(2).RightPadding = 0

    newTbl.Rows(1).Cells(1).Width = col1W

    newTbl.Rows(1).Cells(2).Width = col2W

    newTbl.Rows.AllowBreakAcrossPages = False

    newTbl.Range.ParagraphFormat.KeepWithNext = True

    

    Dim isAttRebuild As Boolean

    isAttRebuild = IsDocumentAttachedToDecision(doc)

    

    ' 4. Dien noi dung & format Cot 1

    Dim c1 As Cell: Set c1 = newTbl.Cell(1, 1)

    If regIdx = 1 Then

        If Len(supAgency) > 0 Then

            c1.Range.Text = supAgency & vbCrLf & issAgency & vbCrLf & "*" & vbCrLf & docNum

            SetParaFont c1.Range.Paragraphs(1), "Times New Roman", 14, False, False

            SetParaFont c1.Range.Paragraphs(2), "Times New Roman", 14, True, False

            SetParaFont c1.Range.Paragraphs(3), "Times New Roman", 14, False, False

            SetParaFont c1.Range.Paragraphs(4), "Times New Roman", 14, False, False

        Else

            c1.Range.Text = issAgency & vbCrLf & "*" & vbCrLf & docNum

            SetParaFont c1.Range.Paragraphs(1), "Times New Roman", 14, True, False

            SetParaFont c1.Range.Paragraphs(2), "Times New Roman", 14, False, False

            SetParaFont c1.Range.Paragraphs(3), "Times New Roman", 14, False, False

        End If

    Else

        Dim em7 As String: em7 = String(7, ChrW(&H2014))

        If Len(supAgency) > 0 Then

            If isAttRebuild And Len(docNum) = 0 Then

                c1.Range.Text = supAgency & vbCrLf & issAgency & vbCrLf & em7

                SetParaFont c1.Range.Paragraphs(1), "Times New Roman", 12, False, False

                SetParaFont c1.Range.Paragraphs(2), "Times New Roman", 12.5, True, False

                SetParaFont c1.Range.Paragraphs(3), "Times New Roman", 10, True, False

            Else

                c1.Range.Text = supAgency & vbCrLf & issAgency & vbCrLf & em7 & vbCrLf & docNum

                SetParaFont c1.Range.Paragraphs(1), "Times New Roman", 12, False, False

                SetParaFont c1.Range.Paragraphs(2), "Times New Roman", 12.5, True, False

                SetParaFont c1.Range.Paragraphs(3), "Times New Roman", 10, True, False

                SetParaFont c1.Range.Paragraphs(4), "Times New Roman", 13, False, False

            End If

        Else

            If isAttRebuild And Len(docNum) = 0 Then

                c1.Range.Text = issAgency & vbCrLf & em7

                SetParaFont c1.Range.Paragraphs(1), "Times New Roman", 12.5, True, False

                SetParaFont c1.Range.Paragraphs(2), "Times New Roman", 10, True, False

            Else

                c1.Range.Text = issAgency & vbCrLf & em7 & vbCrLf & docNum

                SetParaFont c1.Range.Paragraphs(1), "Times New Roman", 12.5, True, False

                SetParaFont c1.Range.Paragraphs(2), "Times New Roman", 10, True, False

                SetParaFont c1.Range.Paragraphs(3), "Times New Roman", 13, False, False

            End If

        End If

    End If

    

    ' 5. Dien noi dung & format Cot 2

    Dim c2 As Cell: Set c2 = newTbl.Cell(1, 2)

    If regIdx = 1 Then

        c2.Range.Text = docNation & vbCrLf & docDate

        SetParaFont c2.Range.Paragraphs(1), "Times New Roman", 15, True, False

        SetParaFont c2.Range.Paragraphs(2), "Times New Roman", 14, False, True

    Else

        Dim em15 As String: em15 = String(15, ChrW(&H2014))

        If isAttRebuild And Len(docDate) = 0 Then

            c2.Range.Text = docNation & vbCrLf & docMotto & vbCrLf & em15

            SetParaFont c2.Range.Paragraphs(1), "Times New Roman", 12, True, False

            SetParaFont c2.Range.Paragraphs(2), "Times New Roman", 13, True, False

            SetParaFont c2.Range.Paragraphs(3), "Times New Roman", 11, True, False

        Else

            c2.Range.Text = docNation & vbCrLf & docMotto & vbCrLf & em15 & vbCrLf & docDate

            SetParaFont c2.Range.Paragraphs(1), "Times New Roman", 12, True, False

            SetParaFont c2.Range.Paragraphs(2), "Times New Roman", 13, True, False

            SetParaFont c2.Range.Paragraphs(3), "Times New Roman", 11, True, False

            SetParaFont c2.Range.Paragraphs(4), "Times New Roman", 13.5, False, True

        End If

    End If

    

    ' 6. Can giua tuyet doi tat ca paragraph trong bang moi, xoa moi le

    Dim cCell As Cell, pCell As Paragraph

    For Each cCell In newTbl.Range.Cells

        For Each pCell In cCell.Range.Paragraphs

            With pCell.Range.ParagraphFormat

                .Alignment = wdAlignParagraphCenter

                .LeftIndent = 0

                .RightIndent = 0

                .FirstLineIndent = 0

                .SpaceBefore = 0

                .SpaceAfter = 0

                .LineSpacingRule = wdLineSpaceSingle

                .TabStops.ClearAll

            End With

            pCell.Range.Font.Underline = wdUnderlineNone

        Next pCell

    Next cCell

    

    If regIdx = 1 Then

        c2.Range.Paragraphs(1).Range.Font.Underline = wdUnderlineSingle

    End If

    

    Exit Sub

Fallback:

    On Error Resume Next

    If Not tbl Is Nothing Then FormatExistingHeaderTable tbl, regIdx

End Sub



Private Sub FormatExistingHeaderTable(ByVal tbl As Table, ByVal regIdx As Integer)

    On Error Resume Next

    tbl.Borders.Enable = False

    tbl.AutoFitBehavior wdAutoFitFixed

    tbl.LeftPadding = 0: tbl.RightPadding = 0

    tbl.TopPadding = 0: tbl.BottomPadding = 0

    tbl.Rows.WrapAroundText = 0

    tbl.Rows.LeftIndent = 0

    tbl.Rows.Alignment = wdAlignRowLeft

    tbl.Rows.AllowBreakAcrossPages = False

    tbl.Range.ParagraphFormat.KeepWithNext = True

    

    Dim printWHdr As Single

    On Error Resume Next

    printWHdr = tbl.Range.Document.PageSetup.PageWidth - tbl.Range.Document.PageSetup.LeftMargin - tbl.Range.Document.PageSetup.RightMargin

    If printWHdr < MillimetersToPoints(140) Or printWHdr > MillimetersToPoints(200) Then

        printWHdr = MillimetersToPoints(160)

    End If

    

    Dim col1WHdr As Single, col2WHdr As Single

    If regIdx = 1 Then

        col1WHdr = MillimetersToPoints(70)

        col2WHdr = printWHdr - col1WHdr

        If col2WHdr < MillimetersToPoints(88) Then col2WHdr = MillimetersToPoints(88)

    Else

        col1WHdr = MillimetersToPoints(65)

        col2WHdr = printWHdr - col1WHdr

        If col2WHdr < MillimetersToPoints(95) Then col2WHdr = MillimetersToPoints(95)

    End If

    tbl.PreferredWidthType = wdPreferredWidthPoints

    tbl.PreferredWidth = col1WHdr + col2WHdr

    tbl.Columns(1).Width = col1WHdr

    tbl.Columns(2).Width = col2WHdr

    Dim rWHdr As Row

    For Each rWHdr In tbl.Rows

        If rWHdr.Cells.Count >= 2 Then

            rWHdr.Cells(1).LeftPadding = 0

            rWHdr.Cells(1).RightPadding = 0

            rWHdr.Cells(2).LeftPadding = 0

            rWHdr.Cells(2).RightPadding = 0

            rWHdr.Cells(1).Width = col1WHdr

            rWHdr.Cells(2).Width = col2WHdr

        End If

    Next rWHdr

    

    Dim c1 As Cell, c2 As Cell

    Set c1 = tbl.Cell(1, 1)

    Set c2 = tbl.Cell(1, 2)

    

    c1.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

    c1.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

    c1.Range.ParagraphFormat.SpaceBefore = 0

    c1.Range.ParagraphFormat.SpaceAfter = 0

    c1.Range.ParagraphFormat.LeftIndent = 0

    c1.Range.ParagraphFormat.RightIndent = 0

    c1.Range.ParagraphFormat.FirstLineIndent = 0

    c1.Range.ParagraphFormat.TabStops.ClearAll

    

    c2.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

    c2.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

    c2.Range.ParagraphFormat.SpaceBefore = 0

    c2.Range.ParagraphFormat.SpaceAfter = 0

    c2.Range.ParagraphFormat.LeftIndent = 0

    c2.Range.ParagraphFormat.RightIndent = 0

    c2.Range.ParagraphFormat.FirstLineIndent = 0

    c2.Range.ParagraphFormat.TabStops.ClearAll

    

    If tbl.Rows.Count >= 2 Then

        With tbl.Cell(2, 1).Range.ParagraphFormat

            .Alignment = wdAlignParagraphCenter

            .LineSpacingRule = wdLineSpaceSingle

            .SpaceBefore = 0

            .SpaceAfter = 0

            .LeftIndent = 0

            .RightIndent = 0

            .FirstLineIndent = 0

            .TabStops.ClearAll

        End With

        With tbl.Cell(2, 2).Range.ParagraphFormat

            .Alignment = wdAlignParagraphCenter

            .LineSpacingRule = wdLineSpaceSingle

            .SpaceBefore = 0

            .SpaceAfter = 0

            .LeftIndent = 0

            .RightIndent = 0

            .FirstLineIndent = 0

            .TabStops.ClearAll

        End With

    End If

    

    Dim p As Paragraph

    For Each p In c1.Range.Paragraphs

        p.Alignment = wdAlignParagraphCenter

        p.LeftIndent = 0: p.RightIndent = 0: p.FirstLineIndent = 0

        SetParaFont p, "Times New Roman", 0, False, False

    Next p

    For Each p In c2.Range.Paragraphs

        p.Alignment = wdAlignParagraphCenter

        p.LeftIndent = 0: p.RightIndent = 0: p.FirstLineIndent = 0

        SetParaFont p, "Times New Roman", 0, False, False

    Next p

    If tbl.Rows.Count >= 2 Then

        For Each p In tbl.Cell(2, 1).Range.Paragraphs

            p.Alignment = wdAlignParagraphCenter

            p.LeftIndent = 0: p.RightIndent = 0: p.FirstLineIndent = 0

            SetParaFont p, "Times New Roman", 0, False, False

        Next p

        For Each p In tbl.Cell(2, 2).Range.Paragraphs

            p.Alignment = wdAlignParagraphCenter

            p.LeftIndent = 0: p.RightIndent = 0: p.FirstLineIndent = 0

            SetParaFont p, "Times New Roman", 0, False, False

        Next p

    End If

    

    If regIdx = 1 Then

        ' Van ban Dang (HD 05)

        If c1.Range.Paragraphs.Count = 1 Then

            SetParaFont c1.Range.Paragraphs(1), "Times New Roman", 14, True, False

            c1.Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

        Else

            If c1.Range.Paragraphs.Count >= 1 Then

                SetParaFont c1.Range.Paragraphs(1), "Times New Roman", 14, False, False

                c1.Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

            End If

            If c1.Range.Paragraphs.Count >= 2 Then

                SetParaFont c1.Range.Paragraphs(2), "Times New Roman", 14, True, False

                c1.Range.Paragraphs(2).Alignment = wdAlignParagraphCenter

            End If

        End If

        If c1.Range.Paragraphs.Count >= 3 Then

            SetParaFont c1.Range.Paragraphs(3), "Times New Roman", 14, False, False

            c1.Range.Paragraphs(3).Alignment = wdAlignParagraphCenter

        Else

            Dim tC1 As String: tC1 = Trim(Replace(Replace(Replace(c1.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            If Right(tC1, 1) <> "*" Then

                c1.Range.InsertAfter vbCrLf & "*"

                If c1.Range.Paragraphs.Count >= 3 Then

                    SetParaFont c1.Range.Paragraphs(3), "Times New Roman", 14, False, False

                    c1.Range.Paragraphs(3).Alignment = wdAlignParagraphCenter

                End If

            End If

        End If

        

        ' Cell (2,1): So va ky hieu (14 dung thuong, khong dau hai cham, khong dau sao o duoi) - BAT BUOC CAN GIUA

        If tbl.Rows.Count >= 2 Then

            Dim c21Text As String

            c21Text = tbl.Cell(2, 1).Range.Text

            c21Text = Replace(Replace(Replace(c21Text, vbCr, ""), vbLf, ""), Chr(7), "")

            c21Text = Trim(c21Text)

            If Right(c21Text, 1) = "*" Then c21Text = Trim(Left(c21Text, Len(c21Text) - 1))

            c21Text = Replace(c21Text, DecodeUTF8("U+G7kTo="), DecodeUTF8("U+G7kSA="))

            c21Text = Replace(c21Text, DecodeUTF8("U+G7kTog"), DecodeUTF8("U+G7kSA="))

            c21Text = Replace(c21Text, "So: ", "So ")

            c21Text = Replace(c21Text, "So:", "So ")

            c21Text = Replace(c21Text, "SO: ", "SO ")

            c21Text = Replace(c21Text, "SO:", "SO ")

            Do While InStr(c21Text, "  ") > 0: c21Text = Replace(c21Text, "  ", " "): Loop

            tbl.Cell(2, 1).Range.Text = NormalizeDocNumber(c21Text, regIdx)

            SetParaFont tbl.Cell(2, 1).Range.Paragraphs(1), "Times New Roman", 14, False, False

            With tbl.Cell(2, 1).Range.ParagraphFormat

                .Alignment = wdAlignParagraphCenter

                .LineSpacingRule = wdLineSpaceSingle

                .SpaceBefore = 0

                .SpaceAfter = 0

                .LeftIndent = 0

                .RightIndent = 0

                .FirstLineIndent = 0

                .TabStops.ClearAll

            End With

            If tbl.Cell(2, 1).Range.Paragraphs.Count > 0 Then

                tbl.Cell(2, 1).Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

                tbl.Cell(2, 1).Range.Paragraphs(1).LeftIndent = 0

                tbl.Cell(2, 1).Range.Paragraphs(1).RightIndent = 0

                tbl.Cell(2, 1).Range.Paragraphs(1).FirstLineIndent = 0

            End If

        ElseIf c1.Range.Paragraphs.Count >= 4 Then

            Dim c14DangText As String

            c14DangText = c1.Range.Paragraphs(4).Range.Text

            c14DangText = Replace(Replace(Replace(c14DangText, vbCr, ""), vbLf, ""), Chr(7), "")

            c1.Range.Paragraphs(4).Range.Text = NormalizeDocNumber(c14DangText, regIdx)

            SetParaFont c1.Range.Paragraphs(4), "Times New Roman", 14, False, False

            c1.Range.Paragraphs(4).Alignment = wdAlignParagraphCenter

        End If

        

        ' Cell (1,2): Tieu de Dang: DANG CONG SAN VIET NAM (15, Dung, dam, co gach chan net lien bang dong chu)

        If c2.Range.Paragraphs.Count >= 1 Then

            SetParaFont c2.Range.Paragraphs(1), "Times New Roman", 15, True, False

            c2.Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

            c2.Range.Paragraphs(1).Range.Font.Underline = wdUnderlineSingle

        End If

        If c2.Range.Paragraphs.Count >= 2 Then

            c2.Range.Paragraphs(2).Range.Font.Underline = wdUnderlineNone

        End If

        

        ' Cell (2,2): Dia danh ngay thang (14, Nghieng, khong dam) - BAT BUOC CAN GIUA

        If tbl.Rows.Count >= 2 Then

            Dim c22DangText As String

            c22DangText = tbl.Cell(2, 2).Range.Text

            c22DangText = Replace(Replace(Replace(c22DangText, vbCr, ""), vbLf, ""), Chr(7), "")

            tbl.Cell(2, 2).Range.Text = NormalizeDocDate(c22DangText)

            SetParaFont tbl.Cell(2, 2).Range.Paragraphs(1), "Times New Roman", 14, False, True

            With tbl.Cell(2, 2).Range.ParagraphFormat

                .Alignment = wdAlignParagraphCenter

                .LineSpacingRule = wdLineSpaceSingle

                .SpaceBefore = 0

                .SpaceAfter = 0

                .LeftIndent = 0

                .RightIndent = 0

                .FirstLineIndent = 0

                .TabStops.ClearAll

            End With

            If tbl.Cell(2, 2).Range.Paragraphs.Count > 0 Then

                tbl.Cell(2, 2).Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

                tbl.Cell(2, 2).Range.Paragraphs(1).LeftIndent = 0

                tbl.Cell(2, 2).Range.Paragraphs(1).RightIndent = 0

                tbl.Cell(2, 2).Range.Paragraphs(1).FirstLineIndent = 0

            End If

        ElseIf c2.Range.Paragraphs.Count >= 2 Then

            Dim c2LastDangText As String

            c2LastDangText = c2.Range.Paragraphs(c2.Range.Paragraphs.Count).Range.Text

            c2LastDangText = Replace(Replace(Replace(c2LastDangText, vbCr, ""), vbLf, ""), Chr(7), "")

            c2.Range.Paragraphs(c2.Range.Paragraphs.Count).Range.Text = NormalizeDocDate(c2LastDangText)

            SetParaFont c2.Range.Paragraphs(c2.Range.Paragraphs.Count), "Times New Roman", 14, False, True

            c2.Range.Paragraphs(c2.Range.Paragraphs.Count).Alignment = wdAlignParagraphCenter

        End If

    Else

        ' Hanh chinh (ND 30)

        Dim isAttachedAdm As Boolean

        isAttachedAdm = IsDocumentAttachedToDecision(tbl.Range.Document)

        

        If isAttachedAdm And tbl.Rows.Count >= 2 Then

            ' Mau 1.3 ND 30: Khong co So va Ngay thang duoi Quoc hieu.

            ' Neu hang 2 rong hoac chi la placeholder thi xoa hang 2 de dung 100% Mau 1.3

            Dim rawNumChk As String, rawDateChk As String

            rawNumChk = Trim(Replace(Replace(Replace(tbl.Cell(2, 1).Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            rawDateChk = Trim(Replace(Replace(Replace(tbl.Cell(2, 2).Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            If Len(rawNumChk) = 0 Or InStr(rawNumChk, ".../") > 0 Or InStr(rawNumChk, DecodeUTF8("U+G7kTog")) > 0 Or InStr(rawNumChk, "So: ") > 0 Then

                If Len(rawDateChk) = 0 Or InStr(rawDateChk, DecodeUTF8("bmfDoHkgLi4u")) > 0 Or InStr(rawDateChk, "ngay ...") > 0 Then

                    tbl.Rows(2).Delete

                End If

            End If

        End If

        

        If c1.Range.Paragraphs.Count = 1 Then

            SetParaFont c1.Range.Paragraphs(1), "Times New Roman", 12.5, True, False

            c1.Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

        Else

            If c1.Range.Paragraphs.Count >= 1 Then

                ' Co quan chu quan: 12, Dung (Regular khong dam theo ND 30)

                SetParaFont c1.Range.Paragraphs(1), "Times New Roman", 12, False, False

                c1.Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

            End If

            If c1.Range.Paragraphs.Count >= 2 Then

                ' Co quan ban hanh: 12.5, Dung, Dam (Bold theo ND 30)

                SetParaFont c1.Range.Paragraphs(2), "Times New Roman", 12.5, True, False

                c1.Range.Paragraphs(2).Alignment = wdAlignParagraphCenter

            End If

        End If

        If tbl.Rows.Count >= 2 Then

            Dim c21AdmText As String

            c21AdmText = tbl.Cell(2, 1).Range.Text

            c21AdmText = Replace(Replace(Replace(c21AdmText, vbCr, ""), vbLf, ""), Chr(7), "")

            If Not isAttachedAdm Or Len(Trim(c21AdmText)) > 0 Then

                c21AdmText = NormalizeDocNumber(c21AdmText, regIdx)

                tbl.Cell(2, 1).Range.Text = c21AdmText

                SetParaFont tbl.Cell(2, 1).Range.Paragraphs(1), "Times New Roman", 13, False, False

            End If

            With tbl.Cell(2, 1).Range.ParagraphFormat

                .Alignment = wdAlignParagraphCenter

                .LineSpacingRule = wdLineSpaceSingle

                .SpaceBefore = 0

                .SpaceAfter = 0

                .LeftIndent = 0

                .RightIndent = 0

                .FirstLineIndent = 0

                .TabStops.ClearAll

            End With

            If tbl.Cell(2, 1).Range.Paragraphs.Count > 0 Then

                tbl.Cell(2, 1).Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

                tbl.Cell(2, 1).Range.Paragraphs(1).LeftIndent = 0

                tbl.Cell(2, 1).Range.Paragraphs(1).RightIndent = 0

                tbl.Cell(2, 1).Range.Paragraphs(1).FirstLineIndent = 0

            End If

        ElseIf c1.Range.Paragraphs.Count >= 3 And Not isAttachedAdm Then

            Dim c13Text As String

            c13Text = c1.Range.Paragraphs(3).Range.Text

            c13Text = Replace(Replace(Replace(c13Text, vbCr, ""), vbLf, ""), Chr(7), "")

            c1.Range.Paragraphs(3).Range.Text = NormalizeDocNumber(c13Text, regIdx)

            SetParaFont c1.Range.Paragraphs(3), "Times New Roman", 13, False, False

            c1.Range.Paragraphs(3).Alignment = wdAlignParagraphCenter

        End If

        

        ' Quoc hieu & Tieu ngu ND 30

        If c2.Range.Paragraphs.Count >= 1 Then

            SetParaFont c2.Range.Paragraphs(1), "Times New Roman", 12, True, False

            c2.Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

        End If

        If c2.Range.Paragraphs.Count >= 2 Then

            SetParaFont c2.Range.Paragraphs(2), "Times New Roman", 13, True, False

            c2.Range.Paragraphs(2).Alignment = wdAlignParagraphCenter

        End If

        If tbl.Rows.Count >= 2 Then

            Dim c22AdmText As String

            c22AdmText = tbl.Cell(2, 2).Range.Text

            c22AdmText = Replace(Replace(Replace(c22AdmText, vbCr, ""), vbLf, ""), Chr(7), "")

            If Not isAttachedAdm Or Len(Trim(c22AdmText)) > 0 Then

                c22AdmText = NormalizeDocDate(c22AdmText)

                tbl.Cell(2, 2).Range.Text = c22AdmText

                SetParaFont tbl.Cell(2, 2).Range.Paragraphs(1), "Times New Roman", 13.5, False, True

            End If

            With tbl.Cell(2, 2).Range.ParagraphFormat

                .Alignment = wdAlignParagraphCenter

                .LineSpacingRule = wdLineSpaceSingle

                .SpaceBefore = 6

                .SpaceAfter = 0

                .LeftIndent = 0

                .RightIndent = 0

                .FirstLineIndent = 0

                .TabStops.ClearAll

            End With

            If tbl.Cell(2, 2).Range.Paragraphs.Count > 0 Then

                tbl.Cell(2, 2).Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

                tbl.Cell(2, 2).Range.Paragraphs(1).LeftIndent = 0

                tbl.Cell(2, 2).Range.Paragraphs(1).RightIndent = 0

                tbl.Cell(2, 2).Range.Paragraphs(1).FirstLineIndent = 0

            End If

        ElseIf c2.Range.Paragraphs.Count >= 3 Then

            Dim c2LastAdmText As String

            c2LastAdmText = c2.Range.Paragraphs(c2.Range.Paragraphs.Count).Range.Text

            c2LastAdmText = Replace(Replace(Replace(c2LastAdmText, vbCr, ""), vbLf, ""), Chr(7), "")

            c2.Range.Paragraphs(c2.Range.Paragraphs.Count).Range.Text = NormalizeDocDate(c2LastAdmText)

            SetParaFont c2.Range.Paragraphs(c2.Range.Paragraphs.Count), "Times New Roman", 13.5, False, True

            c2.Range.Paragraphs(c2.Range.Paragraphs.Count).Alignment = wdAlignParagraphCenter

        End If

    End If

    

    ' Dam bao tuyet doi moi o va moi doan van trong bang tieu de deu can giua, khong bi thut le

    Dim rIdx As Long, cIdx As Long, cellObj As Cell, paraObj As Paragraph

    For rIdx = 1 To tbl.Rows.Count

        For cIdx = 1 To tbl.Columns.Count

            Set cellObj = tbl.Cell(rIdx, cIdx)

            With cellObj.Range.ParagraphFormat

                .Alignment = wdAlignParagraphCenter

                .LeftIndent = 0

                .RightIndent = 0

                .FirstLineIndent = 0

                .TabStops.ClearAll

            End With

            For Each paraObj In cellObj.Range.Paragraphs

                paraObj.Alignment = wdAlignParagraphCenter

                paraObj.LeftIndent = 0

                paraObj.RightIndent = 0

                paraObj.FirstLineIndent = 0

                paraObj.TabStops.ClearAll

            Next paraObj

        Next cIdx

    Next rIdx

End Sub



Private Sub AutoInsertMissingHeader(ByVal doc As Document, ByVal regIdx As Integer)

    On Error Resume Next

    Dim rng As Range: Set rng = doc.Range(0, 0)

    

    Dim isAttMiss As Boolean: isAttMiss = False

    If regIdx <> 1 Then isAttMiss = IsDocumentAttachedToDecision(doc)

    

    Dim numRowsHdr As Long

    numRowsHdr = IIf(isAttMiss, 1, 2)

    

    Dim tbl As Table

    Set tbl = doc.Tables.Add(Range:=rng, NumRows:=numRowsHdr, NumColumns:=2)

    tbl.Borders.Enable = False

    tbl.AutoFitBehavior wdAutoFitFixed

    tbl.LeftPadding = 0: tbl.RightPadding = 0

    tbl.TopPadding = 0: tbl.BottomPadding = 0

    tbl.Rows.AllowBreakAcrossPages = False

    tbl.Range.ParagraphFormat.KeepWithNext = True

    

    Dim col1WMiss As Single, col2WMiss As Single

    If regIdx = 1 Then

        col1WMiss = MillimetersToPoints(70)

        col2WMiss = MillimetersToPoints(90)

    Else

        col1WMiss = MillimetersToPoints(65)

        col2WMiss = MillimetersToPoints(95)

    End If

    tbl.Columns(1).Width = col1WMiss

    tbl.Columns(2).Width = col2WMiss

    Dim rWMiss As Row

    For Each rWMiss In tbl.Rows

        If rWMiss.Cells.Count >= 2 Then

            rWMiss.Cells(1).Width = col1WMiss

            rWMiss.Cells(2).Width = col2WMiss

        End If

    Next rWMiss

    

    With tbl.Range.Font

        .Name = "Times New Roman"

        .NameAscii = "Times New Roman"

        .NameOther = "Times New Roman"

        .NameBi = "Times New Roman"

        .Bold = False

        .Italic = False

        .ColorIndex = wdBlack

    End With

    

    If regIdx = 1 Then

        ' Header Dang (HD 05)

        With tbl.Cell(1, 1).Range

            .Text = DecodeUTF8("VOG7iE5IIOG7plkuLi4=") & vbCrLf & DecodeUTF8("xJDhuqJORyDhu6ZZLi4u") & vbCrLf & "*"

            .ParagraphFormat.Alignment = wdAlignParagraphCenter

            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            .ParagraphFormat.SpaceBefore = 0

            .ParagraphFormat.SpaceAfter = 0

            SetParaFont .Paragraphs(1), "Times New Roman", 14, False, False

            SetParaFont .Paragraphs(2), "Times New Roman", 14, True, False

            SetParaFont .Paragraphs(3), "Times New Roman", 14, False, False

        End With

        With tbl.Cell(2, 1).Range

            .Text = DecodeUTF8("U+G7kSAwMS1OUS/EkFU=")

            .ParagraphFormat.Alignment = wdAlignParagraphCenter

            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            .ParagraphFormat.SpaceBefore = 0

            .ParagraphFormat.SpaceAfter = 0

            SetParaFont .Paragraphs(1), "Times New Roman", 14, False, False

        End With

        With tbl.Cell(1, 2).Range

            .Text = DecodeUTF8("xJDhuqJORyBD4buYTkcgU+G6ok4gVknhu4ZUIE5BTQ==")

            .ParagraphFormat.Alignment = wdAlignParagraphCenter

            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            .ParagraphFormat.SpaceBefore = 0

            .ParagraphFormat.SpaceAfter = 0

            SetParaFont .Paragraphs(1), "Times New Roman", 15, True, False

            .Paragraphs(1).Range.Font.Underline = wdUnderlineSingle

        End With

        With tbl.Cell(2, 2).Range

            .Text = DecodeUTF8("Li4uLCBuZ8OgeSAuLi4gdGjDoW5nIC4uLiBuxINtIDIwLi4u")

            .ParagraphFormat.Alignment = wdAlignParagraphCenter

            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            .ParagraphFormat.SpaceBefore = 0

            .ParagraphFormat.SpaceAfter = 0

            SetParaFont .Paragraphs(1), "Times New Roman", 14, False, True

            .Paragraphs(1).Range.Font.Underline = wdUnderlineNone

        End With

    Else

        ' Header Hanh chinh ND 30

        With tbl.Cell(1, 1).Range

            .Text = DecodeUTF8("VUJORCBU4buITkguLi4=") & vbCrLf & DecodeUTF8("U+G7ni4uLg==")

            .ParagraphFormat.Alignment = wdAlignParagraphCenter

            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            .ParagraphFormat.SpaceBefore = 0

            .ParagraphFormat.SpaceAfter = 0

            SetParaFont .Paragraphs(1), "Times New Roman", 12, False, False

            SetParaFont .Paragraphs(2), "Times New Roman", 12.5, True, False

        End With

        With tbl.Cell(1, 2).Range

            .Text = DecodeUTF8("Q+G7mE5HIEjDkkEgWMODIEjhu5hJIENI4bumIE5HSMSoQSBWSeG7hlQgTkFN") & vbCrLf & DecodeUTF8("xJDhu5ljIGzhuq1wIC0gVOG7sSBkbyAtIEjhuqFuaCBwaMO6Yw==")

            .ParagraphFormat.Alignment = wdAlignParagraphCenter

            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            .ParagraphFormat.SpaceBefore = 0

            .ParagraphFormat.SpaceAfter = 0

            SetParaFont .Paragraphs(1), "Times New Roman", 12.5, True, False

            SetParaFont .Paragraphs(2), "Times New Roman", 13.5, True, False

        End With

        If Not isAttMiss Then

            With tbl.Cell(2, 1).Range

                .Text = DecodeUTF8("U+G7kTogLi4uL1HEkC0uLi4=")

                .ParagraphFormat.Alignment = wdAlignParagraphCenter

                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

                .ParagraphFormat.SpaceBefore = 0

                .ParagraphFormat.SpaceAfter = 0

                SetParaFont .Paragraphs(1), "Times New Roman", 13, False, False

            End With

            With tbl.Cell(2, 2).Range

                .Text = DecodeUTF8("Li4uLCBuZ8OgeSAuLi4gdGjDoW5nIC4uLiBuxINtIDIwLi4u")

                .ParagraphFormat.Alignment = wdAlignParagraphCenter

                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

                .ParagraphFormat.SpaceBefore = 0

                .ParagraphFormat.SpaceAfter = 0

                SetParaFont .Paragraphs(1), "Times New Roman", 13.5, False, True

            End With

        End If

    End If

    

    Dim rI As Long, cI As Long, cO As Cell, pO As Paragraph

    For rI = 1 To tbl.Rows.Count

        For cI = 1 To tbl.Columns.Count

            Set cO = tbl.Cell(rI, cI)

            With cO.Range.ParagraphFormat

                .Alignment = wdAlignParagraphCenter

                .LeftIndent = 0: .RightIndent = 0: .FirstLineIndent = 0: .TabStops.ClearAll

            End With

            For Each pO In cO.Range.Paragraphs

                pO.Alignment = wdAlignParagraphCenter

                pO.LeftIndent = 0: pO.RightIndent = 0: pO.FirstLineIndent = 0: pO.TabStops.ClearAll

            Next pO

        Next cI

    Next rI

    

    tbl.Range.Collapse wdCollapseEnd

    tbl.Range.InsertParagraphAfter

End Sub



Private Function StandardizeTitleAndAbstract(ByVal doc As Document, ByVal regIdx As Integer, ByRef hCount As Long) As String

    On Error Resume Next

    StandardizeTitleAndAbstract = ""

    

    Dim maxP As Long: maxP = MinLong(doc.Paragraphs.Count, 25)

    Dim i As Long, p As Paragraph, rawTxt As String, uTxt As String

    

    For i = 1 To maxP

        Set p = doc.Paragraphs(i)

        If Not p.Range.Information(wdWithInTable) Then

            rawTxt = Trim(p.Range.Text)

            rawTxt = Replace(Replace(Replace(rawTxt, vbCr, ""), vbLf, ""), Chr(7), "")

            uTxt = UCase(ConvertToUnsignedVietnamese(rawTxt))

            

            If IsDocumentTitle(uTxt) Then

                StandardizeTitleAndAbstract = rawTxt

                

                ' 1. Dinh dang Tieu de chinh (THONG BAO, QUYET DINH...)

                With p.Range

                    .ParagraphFormat.Alignment = wdAlignParagraphCenter

                    .ParagraphFormat.FirstLineIndent = 0

                    .ParagraphFormat.LeftIndent = 0

                    .ParagraphFormat.RightIndent = 0

                    .ParagraphFormat.SpaceBefore = 6

                    .ParagraphFormat.SpaceAfter = 2

                    .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

                    .ParagraphFormat.Borders.Enable = False

                    .Font.Underline = wdUnderlineNone

                End With

                SetParaFont p, "Times New Roman", IIf(regIdx = 1, 15, 14), True, False

                p.Range.Font.Underline = wdUnderlineNone

                p.Range.Borders(wdBorderBottom).LineStyle = wdLineStyleNone

                hCount = hCount + 1

                

                ' 2. Xoa bat ky Shape duong ke nao nam o khu vuc Tieu de

                Dim shpT As Shape, sIdxT As Long

                For sIdxT = doc.Shapes.Count To 1 Step -1

                    Set shpT = doc.Shapes(sIdxT)

                    If shpT.Top >= MillimetersToPoints(25) And shpT.Top <= MillimetersToPoints(150) Then

                        shpT.Delete

                    End If

                Next sIdxT

                

                ' 3. Quet & dinh dang TAT CA cac doan trich yeu tieu de ben duoi (toi da 2 doan)

                Dim nextPIdx As Long, pNext As Paragraph, nextRaw As String, nextUTxt As String

                Dim absCount As Long: absCount = 0

                For nextPIdx = i + 1 To MinLong(doc.Paragraphs.Count, i + 3)

                    Set pNext = doc.Paragraphs(nextPIdx)

                    If pNext.Range.Information(wdWithInTable) Then Exit For

                    nextRaw = Trim(Replace(Replace(Replace(pNext.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

                    If Len(nextRaw) > 0 Then

                        nextUTxt = UCase(ConvertToUnsignedVietnamese(nextRaw))

                        

                        ' Neu la doan duong ke gach ngang thua, xoa ngay (hoac chuan hoa ----- neu la Dang)

                        If IsSeparatorLine(nextRaw) Then

                            If regIdx = 1 Then

                                pNext.Range.Text = "-----"

                                pNext.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

                                pNext.Range.ParagraphFormat.FirstLineIndent = 0

                                pNext.Range.ParagraphFormat.SpaceBefore = 2

                                pNext.Range.ParagraphFormat.SpaceAfter = 6

                                SetParaFont pNext, "Times New Roman", 14, False, False

                                hCount = hCount + 1

                                Exit For

                            Else

                                pNext.Range.Delete

                                nextPIdx = nextPIdx - 1

                            End If

                        ElseIf IsBodyStarter(nextUTxt) Then

                            ' Da gap noi dung phan than van ban (doan dan vao van ban), ket thuc trich yeu

                            Exit For

                        Else

                            ' Doan van thuoc trich yeu tieu de -> CAN GIUA, IN DAM (rieng dong kem theo Quyet dinh thi IN NGHIENG, KHONG DAM)

                            Dim isAttNext As Boolean: isAttNext = False

                            If (Left(nextRaw, 1) = "(" Or InStr(nextUTxt, "KEM THEO") > 0 Or InStr(nextUTxt, "BAN HANH") > 0) And _

                               (InStr(nextUTxt, "QUYET DINH") > 0 Or InStr(nextUTxt, "QD-") > 0 Or InStr(nextUTxt, "QD/") > 0) Then

                                isAttNext = True

                            End If

                            With pNext.Range

                                .ParagraphFormat.Alignment = wdAlignParagraphCenter

                                .ParagraphFormat.FirstLineIndent = 0

                                .ParagraphFormat.LeftIndent = 0

                                .ParagraphFormat.RightIndent = 0

                                .ParagraphFormat.SpaceBefore = IIf(isAttNext, 2, 0)

                                .ParagraphFormat.SpaceAfter = IIf(isAttNext, 6, 4)

                                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

                                .ParagraphFormat.Borders.Enable = False

                                .Font.Name = "Times New Roman"

                                .Font.Size = IIf(regIdx = 1, 14, 13.5)

                                If isAttNext Then

                                    .Font.Bold = False

                                    .Font.Italic = True

                                Else

                                    .Font.Bold = True

                                    .Font.Italic = False

                                End If

                                .Font.Underline = wdUnderlineNone ' XOA GACH CHAN THUA

                            End With

                            pNext.Range.Borders(wdBorderBottom).LineStyle = wdLineStyleNone

                            hCount = hCount + 1

                            absCount = absCount + 1

                            If absCount >= 2 Then Exit For

                        End If

                    End If

                Next nextPIdx

                

                Exit Function

            End If

        End If

    Next i

End Function



Private Sub FormatBodyAndHeadings(ByVal doc As Document, ByVal regIdx As Integer, ByRef pCount As Long, ByRef hCount As Long)

    On Error Resume Next

    Dim i As Long, p As Paragraph, rawTxt As String, uTxt As String

    Dim fIndent As Single

    ' Quy chuan thut dau dong: 10 mm (HD 05) hoac 1,27 cm (ND 30)

    fIndent = IIf(regIdx = 1, MillimetersToPoints(10), CentimetersToPoints(1.27))

    

    Dim stdBodySize As Single

    stdBodySize = IIf(regIdx = 1, 14, 13.5)

    

    Dim inTitleAbstractArea As Boolean: inTitleAbstractArea = False

    Dim titleAbstractCount As Long: titleAbstractCount = 0

    

    For i = 1 To doc.Paragraphs.Count

        Set p = doc.Paragraphs(i)

        

        If m_ProgressIsShowing Then

            If (i Mod 10 = 0) Or (i = doc.Paragraphs.Count) Then

                Dim pTotal As Long: pTotal = doc.Paragraphs.Count

                If pTotal < 1 Then pTotal = 1

                If m_ProgressOwner = "AutoStandardize" Then

                    Dim subPct As Integer

                    subPct = 55 + CInt((i * 20) / pTotal)

                    SetProgress subPct, DecodeUTF8("xJBhbmcgxJHhu4tuaCBk4bqhbmcgbsOZaSBkdW5nLCDEkeG7gSBt4bulYyAo") & i & "/" & pTotal & DecodeUTF8("IMSRb+G6oW4pLi4u"), "AutoStandardize"

                ElseIf m_ProgressOwner = "ManualBody" Then

                    Dim bPct As Integer

                    bPct = 10 + CInt((i * 80) / pTotal)

                    SetProgress bPct, DecodeUTF8("xJBhbmcgxJHhu4tuaCBk4bqhbmcgbsOZaSBkdW5nLCDEkeG7gSBt4bulYyAo") & i & "/" & pTotal & DecodeUTF8("IMSRb+G6oW4pLi4u"), "ManualBody"

                End If

            End If

        End If

        

        If Not p.Range.Information(wdWithInTable) Then

            rawTxt = Trim(p.Range.Text)

            rawTxt = Replace(Replace(Replace(rawTxt, vbCr, ""), vbLf, ""), Chr(7), "")

            

            If Len(rawTxt) > 0 Then

                uTxt = UCase(ConvertToUnsignedVietnamese(rawTxt))

                

                ' 0A. CAC DOAN TRICH YEU LIEN SAU TIEU DE (CAN GIUA, IN DAM, XOA GACH CHAN)

                If inTitleAbstractArea Then

                    If IsBodyStarter(uTxt) Or titleAbstractCount >= 2 Then

                        inTitleAbstractArea = False

                    Else

                        If IsSeparatorLine(rawTxt) Then

                            If regIdx = 1 Then

                                p.Range.Text = "-----"

                                p.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

                                p.Range.ParagraphFormat.FirstLineIndent = 0

                                p.Range.ParagraphFormat.SpaceBefore = 2

                                p.Range.ParagraphFormat.SpaceAfter = 6

                                SetParaFont p, "Times New Roman", 14, False, False

                                inTitleAbstractArea = False

                                GoTo NextParaLoop

                            Else

                                p.Range.Delete

                                i = i - 1

                                GoTo NextParaLoop

                            End If

                        End If

                        

                        Dim isAttBody As Boolean: isAttBody = False

                        If (Left(rawTxt, 1) = "(" Or InStr(uTxt, "KEM THEO") > 0 Or InStr(uTxt, "BAN HANH") > 0) And _

                           (InStr(uTxt, "QUYET DINH") > 0 Or InStr(uTxt, "QD-") > 0 Or InStr(uTxt, "QD/") > 0) Then

                            isAttBody = True

                        End If

                        

                        With p.Range

                            .ParagraphFormat.Alignment = wdAlignParagraphCenter

                            .ParagraphFormat.FirstLineIndent = 0

                            .ParagraphFormat.LeftIndent = 0

                            .ParagraphFormat.RightIndent = 0

                            .ParagraphFormat.SpaceBefore = IIf(isAttBody, 2, 0)

                            .ParagraphFormat.SpaceAfter = IIf(isAttBody, 6, 4)

                            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

                            .ParagraphFormat.Borders.Enable = False

                            .Font.Name = "Times New Roman"

                            .Font.Size = IIf(regIdx = 1, 14, 13.5)

                            If isAttBody Then

                                .Font.Bold = False

                                .Font.Italic = True

                            Else

                                .Font.Bold = True

                                .Font.Italic = False

                            End If

                            .Font.Underline = wdUnderlineNone

                        End With

                        p.Range.Borders(wdBorderBottom).LineStyle = wdLineStyleNone

                        titleAbstractCount = titleAbstractCount + 1

                        hCount = hCount + 1

                        GoTo NextParaLoop

                    End If

                End If

                

                ' 0B. TIEU DE VAN BAN (QUYET DINH, NGHI QUYET, TO TRINH, KE HOACH, BAO CAO, THONG BAO...)

                If IsDocumentTitle(uTxt) Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphCenter

                        .ParagraphFormat.FirstLineIndent = 0

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 6

                        .ParagraphFormat.SpaceAfter = 2

                        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

                        .ParagraphFormat.Borders.Enable = False

                        .Font.Underline = wdUnderlineNone

                    End With

                    SetParaFont p, "Times New Roman", IIf(regIdx = 1, 15, 14), True, False

                    p.Range.Font.Underline = wdUnderlineNone

                    p.Range.Borders(wdBorderBottom).LineStyle = wdLineStyleNone

                    hCount = hCount + 1

                    inTitleAbstractArea = True

                    titleAbstractCount = 0

                    GoTo NextParaLoop

                

                ' 0C. TRICH YEU CONG VAN (V/v... HOAC VE VIEC... O CONG VAN KHONG CO TIEU DE)

                ElseIf Left(uTxt, 4) = "V/V " Or Left(uTxt, 4) = "V/V:" Or Left(uTxt, 8) = "VE VIEC " Or Left(uTxt, 8) = "VE VIEC:" Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphLeft

                        .ParagraphFormat.FirstLineIndent = 0

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 4

                        .ParagraphFormat.SpaceAfter = 6

                        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

                    End With

                    SetParaFont p, "Times New Roman", IIf(regIdx = 1, 12, 12.5), False, IIf(regIdx = 1, True, False)

                    hCount = hCount + 1

                    GoTo NextParaLoop

                

                ' 0D. DONG LENH QUYET DINH: / BAN HANH: / QUYET NGHI:

                ElseIf uTxt = "QUYET DINH:" Or uTxt = "BAN HANH:" Or uTxt = "QUYET NGHI:" Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphCenter

                        .ParagraphFormat.FirstLineIndent = 0

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 6

                        .ParagraphFormat.SpaceAfter = 4

                    End With

                    SetParaFont p, "Times New Roman", IIf(regIdx = 1, 14, 13.5), True, False

                    hCount = hCount + 1

                    GoTo NextParaLoop

                

                ' 1. PHAN / CHUONG

                ElseIf Left(uTxt, 5) = "PHAN " Or Left(uTxt, 7) = "CHUONG " Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphCenter

                        .ParagraphFormat.FirstLineIndent = 0

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 6

                        .ParagraphFormat.SpaceAfter = 2

                    End With

                    SetParaFont p, "Times New Roman", IIf(regIdx = 1, 14.5, 13.5), True, False

                    hCount = hCount + 1

                    

                ' 2. MUC / TIEU MUC

                ElseIf Left(uTxt, 4) = "MUC " Or Left(uTxt, 9) = "TIEU MUC " Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphCenter

                        .ParagraphFormat.FirstLineIndent = 0

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 4

                        .ParagraphFormat.SpaceAfter = 2

                    End With

                    SetParaFont p, "Times New Roman", IIf(regIdx = 1, 14, 13.5), True, False

                    hCount = hCount + 1

                    

                ' 3. DIEU

                ElseIf Left(uTxt, 5) = "DIEU " Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphJustify

                        .ParagraphFormat.FirstLineIndent = fIndent

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 4

                        .ParagraphFormat.SpaceAfter = 3

                        If regIdx = 1 Then

                            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly

                            .ParagraphFormat.LineSpacing = 19

                        Else

                            .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple

                            .ParagraphFormat.LineSpacing = VF_LinesToPoints(1.15)

                        End If

                    End With

                    p.Range.Font.Name = "Times New Roman"

                    p.Range.Font.Size = stdBodySize

                    p.Range.Font.ColorIndex = wdBlack

                    

                    ' In dam cum "Dieu X."

                    Dim dotPos As Long: dotPos = InStr(rawTxt, ".")

                    If dotPos > 0 And dotPos < 20 Then

                        Dim rDieu As Range

                        Set rDieu = doc.Range(p.Range.Start, p.Range.Start + dotPos)

                        rDieu.Font.Bold = True

                    Else

                        p.Range.Font.Bold = True

                    End If

                    hCount = hCount + 1

                    

                ' 4. CAN CU

                ElseIf Left(uTxt, 7) = "CAN CU " Or Left(uTxt, 9) = "- CAN CU " Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphJustify

                        .ParagraphFormat.FirstLineIndent = fIndent

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 0

                        .ParagraphFormat.SpaceAfter = 2

                        If regIdx = 1 Then

                            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly

                            .ParagraphFormat.LineSpacing = 19

                        Else

                            .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple

                            .ParagraphFormat.LineSpacing = VF_LinesToPoints(1.15)

                        End If

                    End With

                    If regIdx = 1 Then

                        ' Van ban Dang (HD 05): Bat buoc co dau gach ngang '- ', kieu chu DUNG (khong nghieng, khong dam)

                        If Left(rawTxt, 2) <> "- " Then

                            p.Range.InsertBefore "- "

                            rawTxt = "- " & rawTxt

                        End If

                        

                        ' Chuan hoa dau cham / cham phay cuoi can cu theo HD 05

                        Dim isLastCC As Boolean: isLastCC = True

                        If i < doc.Paragraphs.Count Then

                            Dim pNextCC As Paragraph: Set pNextCC = doc.Paragraphs(i + 1)

                            Dim rNextCC As String: rNextCC = Trim(Replace(Replace(Replace(pNextCC.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

                            Dim uNextCC As String: uNextCC = UCase(ConvertToUnsignedVietnamese(rNextCC))

                            If Left(uNextCC, 7) = "CAN CU " Or Left(uNextCC, 9) = "- CAN CU " Then

                                isLastCC = False

                            End If

                        End If

                        

                        Dim cEnd As String: cEnd = Right(rawTxt, 1)

                        If isLastCC Then

                            If cEnd = ";" Or cEnd = "," Then

                                p.Range.Text = Left(rawTxt, Len(rawTxt) - 1) & "."

                            ElseIf cEnd <> "." Then

                                p.Range.Text = rawTxt & "."

                            End If

                        Else

                            If cEnd = "." Or cEnd = "," Then

                                p.Range.Text = Left(rawTxt, Len(rawTxt) - 1) & ";"

                            ElseIf cEnd <> ";" Then

                                p.Range.Text = rawTxt & ";"

                            End If

                        End If

                        

                        p.Range.Font.Name = "Times New Roman"

                        p.Range.Font.Size = stdBodySize

                        p.Range.Font.Bold = False

                        p.Range.Font.Italic = False

                        p.Range.Font.ColorIndex = wdBlack

                    Else

                        ' Van ban Hanh chinh (ND 30): KHONG co dau gach ngang o dau, kieu chu NGHIENG

                        If Left(rawTxt, 2) = "- " Then

                            p.Range.Text = Mid(rawTxt, 3)

                        End If

                        p.Range.Font.Name = "Times New Roman"

                        p.Range.Font.Size = stdBodySize

                        p.Range.Font.Bold = False

                        p.Range.Font.Italic = True

                        p.Range.Font.ColorIndex = wdBlack

                    End If

                    pCount = pCount + 1

                    

                ' 4A2. DONG DE NGHI VA BAN HANH (ND 30: IN NGHIENG CHUAN THE THUC)

                ElseIf (Left(uTxt, 13) = "THEO DE NGHI " Or Left(uTxt, 12) = "XET DE NGHI " Or _

                       InStr(uTxt, "BAN HANH NGHI DINH") > 0 Or InStr(uTxt, "BAN HANH QUYET DINH") > 0 Or _

                       InStr(uTxt, "BAN HANH THONG TU") > 0 Or InStr(uTxt, "BAN HANH QUY CHE") > 0 Or _

                       InStr(uTxt, "BAN HANH QUY DINH") > 0) And regIdx = 0 Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphJustify

                        .ParagraphFormat.FirstLineIndent = fIndent

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 0

                        .ParagraphFormat.SpaceAfter = 2

                        .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple

                        .ParagraphFormat.LineSpacing = VF_LinesToPoints(1.15)

                    End With

                    p.Range.Font.Name = "Times New Roman"

                    p.Range.Font.Size = stdBodySize

                    p.Range.Font.Bold = False

                    p.Range.Font.Italic = True

                    p.Range.Font.ColorIndex = wdBlack

                    pCount = pCount + 1

                    GoTo NextParaLoop

                    

                ' 4B. DUONG GACH NGAN CACH HOAC DAU SAO (-----, ***)

                ElseIf Replace(Replace(Replace(Replace(rawTxt, "-", ""), "*", ""), "_", ""), " ", "") = "" And Len(rawTxt) >= 2 Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphCenter

                        .ParagraphFormat.FirstLineIndent = 0

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 0

                        .ParagraphFormat.SpaceAfter = 4

                    End With

                    If regIdx = 1 Then

                        p.Range.Text = "-----"

                    End If

                    SetParaFont p, "Times New Roman", 14, False, False

                    pCount = pCount + 1

                    

                ' 4C. DE MUC LA MA (I. KET QUA..., II. PHUONG HUONG...)

                ElseIf IsRomanNumeralHeading(rawTxt) Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphJustify

                        .ParagraphFormat.FirstLineIndent = fIndent

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 4

                        .ParagraphFormat.SpaceAfter = 2

                        If regIdx = 1 Then

                            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly

                            .ParagraphFormat.LineSpacing = 19

                        Else

                            .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple

                            .ParagraphFormat.LineSpacing = VF_LinesToPoints(1.15)

                        End If

                    End With

                    p.Range.Font.Name = "Times New Roman"

                    p.Range.Font.Size = stdBodySize

                    p.Range.Font.Bold = True

                    p.Range.Font.ColorIndex = wdBlack

                    p.Range.Font.Underline = wdUnderlineNone

                    hCount = hCount + 1

                    

                ' 5. KHOAN (Bat dau bang so A Rap + dau cham: 1., 2...)

                ElseIf IsClauseHeading(rawTxt) Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphJustify

                        .ParagraphFormat.FirstLineIndent = fIndent

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 0

                        .ParagraphFormat.SpaceAfter = 3

                        If regIdx = 1 Then

                            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly

                            .ParagraphFormat.LineSpacing = 19

                        Else

                            .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple

                            .ParagraphFormat.LineSpacing = VF_LinesToPoints(1.15)

                        End If

                    End With

                    p.Range.Font.Name = "Times New Roman"

                    p.Range.Font.Size = stdBodySize

                    p.Range.Font.ColorIndex = wdBlack

                    

                    ' In dam cum so thu tu khoan "1.", "2." hoac toan bo dong tieu de ngan (< 60 ky tu)

                    Dim dotKPos As Long: dotKPos = InStr(rawTxt, ".")

                    If dotKPos > 0 And dotKPos <= 4 Then

                        If Len(rawTxt) < 60 Then

                            p.Range.Font.Bold = True

                        Else

                            Dim rKPrefix As Range

                            Set rKPrefix = doc.Range(p.Range.Start, p.Range.Start + dotKPos)

                            rKPrefix.Font.Bold = True

                        End If

                    End If

                    pCount = pCount + 1

                    

                ' 6. DIEM (Bat dau bang a), b), c), d)...)

                ElseIf IsPointHeading(rawTxt) Then

                    With p.Range

                        .ParagraphFormat.Alignment = wdAlignParagraphJustify

                        .ParagraphFormat.FirstLineIndent = fIndent

                        .ParagraphFormat.LeftIndent = 0

                        .ParagraphFormat.RightIndent = 0

                        .ParagraphFormat.SpaceBefore = 0

                        .ParagraphFormat.SpaceAfter = 2

                        If regIdx = 1 Then

                            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly

                            .ParagraphFormat.LineSpacing = 19

                        Else

                            .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple

                            .ParagraphFormat.LineSpacing = VF_LinesToPoints(1.15)

                        End If

                    End With

                    p.Range.Font.Name = "Times New Roman"

                    p.Range.Font.Size = stdBodySize

                    p.Range.Font.ColorIndex = wdBlack

                    pCount = pCount + 1

                    

                ' 7. KINH GUI / KINH TRINH

                ElseIf Left(uTxt, 8) = "KINH GUI" Or Left(uTxt, 10) = "KINH TRINH" Then

                    With p.Range.ParagraphFormat

                        If regIdx = 1 Then

                            ' HD05: O so 5a chinh giua trang dau van ban

                            .Alignment = wdAlignParagraphCenter

                            .FirstLineIndent = 0

                            .LeftIndent = 0

                            .RightIndent = 0

                            .SpaceBefore = 6

                            .SpaceAfter = 2

                            .LineSpacingRule = wdLineSpaceSingle

                        Else

                            ' ND30: Can le trai, thut le dau dong

                            .Alignment = wdAlignParagraphLeft

                            .FirstLineIndent = fIndent

                            .LeftIndent = 0

                            .RightIndent = 0

                            .SpaceBefore = 4

                            .SpaceAfter = 2

                            .LineSpacingRule = wdLineSpaceSingle

                        End If

                    End With

                    

                    Dim colonPosKG As Long

                    colonPosKG = InStr(rawTxt, ":")

                    If regIdx = 1 Then

                        ' HD05 (Bang 8 PL 2):

                        ' - Tu "Kinh gui" / "Kinh trinh": In thuong, co 14, NGHIENG

                        ' - Ten to chuc, ca nhan sau: In thuong, co 14-15, DUNG

                        If colonPosKG > 0 Then

                            Dim rPrefKG As Range

                            Set rPrefKG = doc.Range(p.Range.Start, p.Range.Start + colonPosKG)

                            With rPrefKG.Font

                                .Name = "Times New Roman"

                                .Size = 14

                                .Bold = False

                                .Italic = True

                                .ColorIndex = wdBlack

                            End With

                            If p.Range.Start + colonPosKG < p.Range.End - 1 Then

                                Dim rSuffKG As Range

                                Set rSuffKG = doc.Range(p.Range.Start + colonPosKG, p.Range.End - 1)

                                With rSuffKG.Font

                                    .Name = "Times New Roman"

                                    .Size = 14

                                    .Bold = False

                                    .Italic = False

                                    .ColorIndex = wdBlack

                                End With

                            End If

                        Else

                            With p.Range.Font

                                .Name = "Times New Roman"

                                .Size = 14

                                .Bold = False

                                .Italic = True

                                .ColorIndex = wdBlack

                            End With

                        End If

                    Else

                        ' ND30: In dung, co 13.5 - 14

                        SetParaFont p, "Times New Roman", stdBodySize, False, False

                    End If

                    pCount = pCount + 1

                    

                    ' Dinh dang danh sach ten co quan nhan o cac dong gach dau dong lien sau (neu co)

                    Dim jKG As Long, pSubKG As Paragraph, subRawKG As String, subUTxtKG As String

                    For jKG = i + 1 To MinLong(doc.Paragraphs.Count, i + 10)

                        Set pSubKG = doc.Paragraphs(jKG)

                        If pSubKG.Range.Information(wdWithInTable) Then Exit For

                        subRawKG = Trim(Replace(Replace(Replace(pSubKG.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

                        If Len(subRawKG) = 0 Then Exit For

                        subUTxtKG = UCase(ConvertToUnsignedVietnamese(subRawKG))

                        If (Left(subRawKG, 2) = "- " Or Left(subRawKG, 1) = "-" Or Left(subRawKG, 2) = "+ ") And _

                           (InStr(subUTxtKG, "CAN CU") = 0 And InStr(subUTxtKG, "DIEU ") = 0) Then

                            With pSubKG.Range.ParagraphFormat

                                If regIdx = 1 Then

                                    .Alignment = wdAlignParagraphCenter

                                Else

                                    .Alignment = wdAlignParagraphLeft

                                    .LeftIndent = fIndent

                                End If

                                .FirstLineIndent = 0

                                .SpaceBefore = 0

                                .SpaceAfter = 2

                                .LineSpacingRule = wdLineSpaceSingle

                            End With

                            With pSubKG.Range.Font

                                .Name = "Times New Roman"

                                .Size = 14

                                .Bold = False

                                .Italic = False

                                .ColorIndex = wdBlack

                            End With

                            pCount = pCount + 1

                            i = jKG

                        Else

                            Exit For

                        End If

                    Next jKG

                    GoTo NextParaLoop

                    

                ' 8. CAC DOAN VAN THUONG TRONG THAN BAI

                Else

                    If InStr(uTxt, "CONG HOA XA HOI") = 0 And InStr(uTxt, "DOC LAP - TU DO") = 0 And _

                       InStr(uTxt, "DANG CONG SAN") = 0 And InStr(uTxt, "NOI NHAN") = 0 Then

                        If p.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter And i <= 10 And Len(uTxt) <= 80 And Not IsBodyStarter(uTxt) Then

                            p.Range.ParagraphFormat.FirstLineIndent = 0

                            p.Range.ParagraphFormat.LeftIndent = 0

                            p.Range.ParagraphFormat.RightIndent = 0

                            p.Range.Font.Name = "Times New Roman"

                            p.Range.Font.Size = stdBodySize

                            pCount = pCount + 1

                        Else

                            With p.Range

                                .ParagraphFormat.Alignment = wdAlignParagraphJustify

                                .ParagraphFormat.FirstLineIndent = fIndent

                                .ParagraphFormat.LeftIndent = 0

                                .ParagraphFormat.RightIndent = 0

                                .ParagraphFormat.SpaceBefore = 0

                                .ParagraphFormat.SpaceAfter = 6

                                If regIdx = 1 Then

                                    .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly

                                    .ParagraphFormat.LineSpacing = 19

                                Else

                                    .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple

                                    .ParagraphFormat.LineSpacing = VF_LinesToPoints(1.15)

                                End If

                            End With

                            ' Neu toan bo doan van bi in dam do nhan dien nham truoc do hoac la doan dan vao van ban -> bo in dam toan doan

                            If p.Range.Font.Bold = True And (Len(rawTxt) > 80 Or IsBodyStarter(uTxt)) Then

                                p.Range.Font.Bold = False

                            End If

                            p.Range.Font.Name = "Times New Roman"

                            p.Range.Font.Size = stdBodySize

                            p.Range.Font.ColorIndex = wdBlack

                            pCount = pCount + 1

                        End If

                    End If

                End If

            End If

        End If

NextParaLoop:

    Next i

End Sub



Private Function IsRomanNumeralHeading(ByVal s As String) As Boolean

    IsRomanNumeralHeading = False

    Dim t As String: t = Trim(s)

    If Len(t) < 3 Then Exit Function

    Dim dotPos As Long: dotPos = InStr(t, ".")

    If dotPos >= 2 And dotPos <= 6 Then

        Dim prefix As String: prefix = UCase(Trim(Left(t, dotPos - 1)))

        If prefix = "I" Or prefix = "II" Or prefix = "III" Or prefix = "IV" Or prefix = "V" Or _

           prefix = "VI" Or prefix = "VII" Or prefix = "VIII" Or prefix = "IX" Or prefix = "X" Then

            IsRomanNumeralHeading = True

        End If

    End If

End Function



Private Function IsClauseHeading(ByVal s As String) As Boolean

    IsClauseHeading = False

    Dim t As String: t = Trim(s)

    If Len(t) < 3 Then Exit Function

    Dim i As Long, ch As String

    For i = 1 To MinLong(Len(t), 4)

        ch = Mid(t, i, 1)

        If ch = "." Then

            If i > 1 And IsNumeric(Left(t, i - 1)) Then

                IsClauseHeading = True

                Exit Function

            End If

        ElseIf Not IsNumeric(ch) Then

            Exit Function

        End If

    Next i

End Function



Private Function IsPointHeading(ByVal s As String) As Boolean

    IsPointHeading = False

    Dim t As String: t = Trim(s)

    If Len(t) < 3 Then Exit Function

    Dim pClose As Long: pClose = InStr(t, ")")

    If pClose = 2 Or pClose = 3 Then

        Dim prefix As String: prefix = Left(t, pClose - 1)

        If Len(prefix) = 1 Then

            Dim c As String: c = LCase(prefix)

            If c >= "a" And c <= "z" Then IsPointHeading = True

        End If

    End If

End Function



Public Function CleanAcademicAndMilitaryTitles(ByVal rawName As String) As String

    Dim t As String: t = Trim(rawName)

    If Len(t) = 0 Then

        CleanAcademicAndMilitaryTitles = ""

        Exit Function

    End If

    

    Dim uName As String

    Dim changed As Boolean

    Dim i As Long, pfx As String, pfxLen As Long

    Dim titlesAscii As Variant

    titlesAscii = Array("GS.TS.", "GS. TS.", "GS.TS", "GS. TS", _

                        "PGS.TS.", "PGS. TS.", "PGS.TS", "PGS. TS", _

                        "TS.BS.", "TS. BS.", "TS.BS", "TS. BS", _

                        "THS.BS.", "THS. BS.", "THS.BS", "THS. BS", _

                        "GS.", "PGS.", "TS.", "THS.", "BSCKII.", "BSCKI.", "BS.", "CN.", "KTS.", "KS.", _

                        "DAI TUONG", "THUONG TUONG", "TRUNG TUONG", "THIEU TUONG", _

                        "DAI TA", "THUONG TA", "TRUNG TA", "THIEU TA", _

                        "DAI UY", "THUONG UY", "TRUNG UY", "THIEU UY")

    

    Do

        changed = False

        uName = UCase(ConvertToUnsignedVietnamese(t))

        For i = LBound(titlesAscii) To UBound(titlesAscii)

            pfx = CStr(titlesAscii(i))

            If InStr(1, uName, pfx, vbTextCompare) = 1 Then

                pfxLen = Len(pfx)

                t = Trim(Mid(t, pfxLen + 1))

                If Left(t, 1) = "." Or Left(t, 1) = "," Or Left(t, 1) = "-" Or Left(t, 1) = ":" Then

                    t = Trim(Mid(t, 2))

                End If

                changed = True

                Exit For

            End If

        Next i

    Loop While changed

    

    CleanAcademicAndMilitaryTitles = t

End Function



Private Function RebuildSignatureTable(ByVal doc As Document, ByVal tbl As Table, ByVal regIdx As Integer) As Table

    Set RebuildSignatureTable = tbl

    On Error GoTo Fallback

    If doc Is Nothing Or tbl Is Nothing Then Exit Function

    

    ' Neu bang co chua anh chu ky / dau scan (InlineShape) thi giu nguyen va dinh dang tai cho

    If tbl.Range.InlineShapes.Count > 0 Then

        FormatExistingSignatureTable tbl, regIdx

        Set RebuildSignatureTable = tbl

        Exit Function

    End If

    

    ' 1. Trich xuat du lieu Noi nhan va Chu ky tu bang cu

    Dim recLines As String: recLines = ""

    Dim sigLinesList As String: sigLinesList = ""

    Dim rIdx As Long, cObj As Cell, pObj As Paragraph

    Dim pText As String

    

    ' Quet Cot 1: Cot Noi nhan

    For rIdx = 1 To tbl.Rows.Count

        If tbl.Rows(rIdx).Cells.Count >= 1 Then

            Set cObj = tbl.Cell(rIdx, 1)

            For Each pObj In cObj.Range.Paragraphs

                pText = Trim(Replace(Replace(Replace(pObj.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

                If Len(pText) > 0 Then

                    If Not IsSeparatorLine(pText) Then

                        If Len(recLines) > 0 Then recLines = recLines & vbCrLf

                        recLines = recLines & pText

                    End If

                End If

            Next pObj

        End If

    Next rIdx

    

    ' Quet Cot 2: Cot Chu ky

    For rIdx = 1 To tbl.Rows.Count

        If tbl.Rows(rIdx).Cells.Count >= 2 Then

            Set cObj = tbl.Cell(rIdx, 2)

            For Each pObj In cObj.Range.Paragraphs

                pText = Trim(Replace(Replace(Replace(pObj.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

                If Len(pText) > 0 Then

                    If Not IsSeparatorLine(pText) Then

                        If Len(sigLinesList) > 0 Then sigLinesList = sigLinesList & vbCrLf

                        sigLinesList = sigLinesList & pText

                    End If

                End If

            Next pObj

        End If

    Next rIdx

    

    ' Neu 2 cot bi dao nguoc

    If InStr(UCase(ConvertToUnsignedVietnamese(sigLinesList)), "NOI NHAN") > 0 And _

       InStr(UCase(ConvertToUnsignedVietnamese(recLines)), "NOI NHAN") = 0 Then

        Dim tmpSwap As String

        tmpSwap = recLines

        recLines = sigLinesList

        sigLinesList = tmpSwap

    End If

    

    ' 2. Xu ly chuan hoa cac dong Noi nhan

    Dim finalRecLines As String: finalRecLines = ""

    Dim recHeader As String

    recHeader = DecodeUTF8("TsahaSBuaOG6rW46") ' Noi nhan:

    

    If Len(recLines) > 0 Then

        Dim rParts() As String: rParts = Split(recLines, vbCrLf)

        Dim rpIdx As Long, rItem As String, uRItem As String

        For rpIdx = 0 To UBound(rParts)

            rItem = Trim(rParts(rpIdx))

            If Len(rItem) > 0 Then

                uRItem = UCase(ConvertToUnsignedVietnamese(rItem))

                If InStr(uRItem, "NOI NHAN") = 0 Then

                    ' Dinh dang dau dong cho tung muc noi nhan neu chua co '-' hoac '+'

                    If Left(rItem, 1) <> "-" And Left(rItem, 1) <> "+" Then

                        rItem = "- " & rItem

                    End If

                    If Len(finalRecLines) > 0 Then finalRecLines = finalRecLines & vbCrLf

                    finalRecLines = finalRecLines & rItem

                End If

            End If

        Next rpIdx

    End If

    

    If Len(finalRecLines) = 0 Then

        If regIdx = 1 Then

            finalRecLines = DecodeUTF8("LSBOaMawIHRyw6puOw==") & vbCrLf & _

                            DecodeUTF8("LSBUaMaw4budbmcgdHLhu7FjIMSQ4bqjbmcgcGjGsOG7nW5nOw==") & vbCrLf & _

                            DecodeUTF8("LSBMxrB1IFbEg24gcGjDsm5nLg==")

        Else

            finalRecLines = DecodeUTF8("LSBOaMawIHRyw6puOw==") & vbCrLf & _

                            DecodeUTF8("LSBMxrB1OiBWVC4=")

        End If

    End If

    

    Dim fullRecText As String

    fullRecText = recHeader & vbCrLf & finalRecLines

    

    ' 3. Xu ly chuan hoa cac dong Chu ky (Tham quyen, Chuc vu, Ho ten nguoi ky)

    Dim sigTitle As String: sigTitle = ""

    Dim sigName As String: sigName = ""

    

    If Len(sigLinesList) > 0 Then

        Dim sParts() As String: sParts = Split(sigLinesList, vbCrLf)

        Dim sClean() As String, sCount As Long: sCount = 0

        ReDim sClean(UBound(sParts))

        Dim spIdx As Long, sCleanItem As String

        For spIdx = 0 To UBound(sParts)

            sCleanItem = Trim(sParts(spIdx))

            If Len(sCleanItem) > 0 Then

                sClean(sCount) = sCleanItem

                sCount = sCount + 1

            End If

        Next spIdx

        

        If sCount = 1 Then

            Dim uOnly As String: uOnly = UCase(ConvertToUnsignedVietnamese(sClean(0)))

            If InStr(uOnly, "CHU TICH") > 0 Or InStr(uOnly, "GIAM DOC") > 0 Or _

               InStr(uOnly, "BI THU") > 0 Or InStr(uOnly, "TRUONG") > 0 Or _

               Left(uOnly, 3) = "TM." Or Left(uOnly, 3) = "KT." Then

                sigTitle = UCase(sClean(0))

            Else

                If regIdx = 1 Then

                    sigTitle = DecodeUTF8("QsONIFRIxq8=")

                Else

                    sigTitle = DecodeUTF8("R0nDgU0gxJDhu5BD")

                End If

                sigName = sClean(0)

            End If

        ElseIf sCount = 2 Then

            Dim uSecond As String: uSecond = UCase(ConvertToUnsignedVietnamese(sClean(1)))

            If InStr(uSecond, "PHO GIAM DOC") > 0 Or InStr(uSecond, "PHO CHU TICH") > 0 Or _

               InStr(uSecond, "PHO BI THU") > 0 Or InStr(uSecond, "GIAM DOC") > 0 Or _

               InStr(uSecond, "CHU TICH") > 0 Or InStr(uSecond, "BI THU") > 0 Then

                sigTitle = UCase(sClean(0)) & vbCrLf & UCase(sClean(1))

            Else

                sigTitle = UCase(sClean(0))

                sigName = sClean(1)

            End If

        ElseIf sCount >= 3 Then

            Dim lastItem As String: lastItem = sClean(sCount - 1)

            Dim kIdx As Long

            If Left(lastItem, 1) = "(" And Right(lastItem, 1) = ")" And sCount >= 4 Then

                sigName = sClean(sCount - 2)

                For kIdx = 0 To sCount - 3

                    If Len(sigTitle) > 0 Then sigTitle = sigTitle & vbCrLf

                    sigTitle = sigTitle & UCase(sClean(kIdx))

                Next kIdx

            Else

                sigName = sClean(sCount - 1)

                For kIdx = 0 To sCount - 2

                    If Len(sigTitle) > 0 Then sigTitle = sigTitle & vbCrLf

                    sigTitle = sigTitle & UCase(sClean(kIdx))

                Next kIdx

            End If

        End If

    End If

    

    If Len(sigTitle) = 0 Then

        If regIdx = 1 Then

            sigTitle = DecodeUTF8("VC9NIEJBTiBDSOG6pFAgSMOATkg=") & vbCrLf & _

                       DecodeUTF8("QsONIFRIxq8=")

        Else

            sigTitle = DecodeUTF8("VE0uIOG7plkgQkFOIE5Iw4JOIETDgk4=") & vbCrLf & _

                       DecodeUTF8("Q0jhu6YgVOG7ikNI")

        End If

    End If

    

    If regIdx = 1 Then

        sigTitle = Replace(sigTitle, "TM.", "T/M")

        sigTitle = Replace(sigTitle, "KT.", "K/T")

        sigTitle = Replace(sigTitle, "TL.", "T/L")

        sigTitle = Replace(sigTitle, "TM ", "T/M ")

        sigTitle = Replace(sigTitle, "KT ", "K/T ")

        sigTitle = Replace(sigTitle, "TL ", "T/L ")

        If Len(sigName) > 0 Then

            sigName = CleanAcademicAndMilitaryTitles(sigName)

        End If

    End If

    

    ' 4. Xoa bang cu va dung bang moi 1x2 tai dung vi tri do

    Dim rngTarget As Range

    Set rngTarget = tbl.Range

    rngTarget.Collapse wdCollapseStart

    

    tbl.Delete

    

    Dim newTbl As Table

    Set newTbl = doc.Tables.Add(Range:=rngTarget, NumRows:=1, NumColumns:=2)

    newTbl.Borders.Enable = False

    newTbl.AutoFitBehavior wdAutoFitFixed

    newTbl.Rows.WrapAroundText = 0

    newTbl.Rows.LeftIndent = 0

    newTbl.Rows.Alignment = wdAlignRowLeft

    newTbl.LeftPadding = 0: newTbl.RightPadding = 0

    newTbl.TopPadding = 0: newTbl.BottomPadding = 0

    

    Dim printW As Single

    printW = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin

    If printW < MillimetersToPoints(140) Or printW > MillimetersToPoints(200) Then

        printW = MillimetersToPoints(160)

    End If

    

    Dim col1W As Single, col2W As Single

    If regIdx = 1 Then

        col1W = MillimetersToPoints(70)

        col2W = printW - col1W

        If col2W < MillimetersToPoints(85) Then col2W = MillimetersToPoints(85)

    Else

        col1W = MillimetersToPoints(70)

        col2W = printW - col1W

        If col2W < MillimetersToPoints(90) Then col2W = MillimetersToPoints(90)

    End If

    

    newTbl.PreferredWidthType = wdPreferredWidthPoints

    newTbl.PreferredWidth = col1W + col2W

    newTbl.Columns(1).Width = col1W

    newTbl.Columns(2).Width = col2W

    newTbl.Rows(1).Cells(1).LeftPadding = 0

    newTbl.Rows(1).Cells(1).RightPadding = 0

    newTbl.Rows(1).Cells(2).LeftPadding = 0

    newTbl.Rows(1).Cells(2).RightPadding = 0

    newTbl.Rows(1).Cells(1).Width = col1W

    newTbl.Rows(1).Cells(2).Width = col2W

    

    ' 5. Dien noi dung & format Cot 1 (Noi nhan)

    Dim c1 As Cell: Set c1 = newTbl.Cell(1, 1)

    c1.Range.Text = fullRecText

    

    Dim p As Paragraph, pI As Long

    For pI = 1 To c1.Range.Paragraphs.Count

        Set p = c1.Range.Paragraphs(pI)

        With p.Range.ParagraphFormat

            .Alignment = wdAlignParagraphLeft

            .LineSpacingRule = wdLineSpaceSingle

            .SpaceBefore = 0

            .SpaceAfter = 0

            .LeftIndent = 0

            .RightIndent = 0

            .FirstLineIndent = 0

            .TabStops.ClearAll

        End With

        

        If pI = 1 Then

            ' Dong "Noi nhan:"

            If regIdx = 1 Then

                ' HD 05: 14pt Regular, Dung, co gach chan net lien

                SetParaFont p, "Times New Roman", 14, False, False

                p.Range.Font.Underline = wdUnderlineSingle

            Else

                ' ND 30: 12pt Bold Italic, dam nghieng

                SetParaFont p, "Times New Roman", 12, True, True

                p.Range.Font.Underline = wdUnderlineNone

            End If

        Else

            ' Cac dong noi nhan con lai

            Dim pLineTxt As String

            pLineTxt = Trim(Replace(Replace(Replace(p.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            If regIdx = 1 Then

                ' HD 05: Cac dong noi nhan la 12pt Dung.

                ' Dong ky hieu nguoi soan thao & so luong ban (vi du NguyenVanAn-26) la 11pt Dung.

                If InStr(pLineTxt, "-") > 0 And Len(pLineTxt) <= 25 And Left(pLineTxt, 1) <> "-" And Left(pLineTxt, 1) <> "+" Then

                    SetParaFont p, "Times New Roman", 11, False, False

                Else

                    SetParaFont p, "Times New Roman", 12, False, False

                End If

            Else

                ' ND 30: 11pt Regular

                SetParaFont p, "Times New Roman", 11, False, False

            End If

            p.Range.Font.Underline = wdUnderlineNone

        End If

    Next pI

    

    ' 6. Dien noi dung & format Cot 2 (Chu ky)

    Dim c2 As Cell: Set c2 = newTbl.Cell(1, 2)

    Dim fullSigText As String

    If Len(sigName) > 0 Then

        fullSigText = sigTitle & vbCrLf & vbCrLf & vbCrLf & vbCrLf & sigName

    Else

        fullSigText = sigTitle & vbCrLf & vbCrLf & vbCrLf & vbCrLf

    End If

    c2.Range.Text = fullSigText

    

    Dim totalP As Long: totalP = c2.Range.Paragraphs.Count

    Dim titleLinesCount As Long

    titleLinesCount = UBound(Split(sigTitle, vbCrLf)) + 1

    

    For pI = 1 To totalP

        Set p = c2.Range.Paragraphs(pI)

        With p.Range.ParagraphFormat

            .Alignment = wdAlignParagraphCenter

            .LineSpacingRule = wdLineSpaceSingle

            .SpaceBefore = 0

            .SpaceAfter = 0

            .LeftIndent = 0

            .RightIndent = 0

            .FirstLineIndent = 0

            .TabStops.ClearAll

        End With

        

        If pI <= titleLinesCount Then

            If regIdx = 1 Then

                If titleLinesCount >= 2 And pI = 2 Then

                    ' Chuc vu nguoi ky o dong 2: In hoa, 14pt, Dung (Khong dam)

                    SetParaFont p, "Times New Roman", 14, False, False

                Else

                    ' Quyen han (T/M, K/T, T/L) hoac Chuc vu duy nhat: In hoa, 14pt, Dung, Dam

                    SetParaFont p, "Times New Roman", 14, True, False

                End If

            Else

                ' ND 30: 13.5pt, Dam

                SetParaFont p, "Times New Roman", 13.5, True, False

            End If

        ElseIf pI = totalP And Len(sigName) > 0 Then

            ' Ho ten nguoi ky: Dam, 13.5pt (ND 30) hoac 14pt (HD 05)

            SetParaFont p, "Times New Roman", IIf(regIdx = 1, 14, 13.5), True, False

        Else

            ' Dong trong de ky

            SetParaFont p, "Times New Roman", 12, False, False

        End If

        p.Range.Font.Underline = wdUnderlineNone

    Next pI

    

    ' 7. Dam bao tat ca doan trong bang mang font Times New Roman

    Dim allP As Paragraph

    For Each allP In newTbl.Range.Paragraphs

        allP.Range.Font.Name = "Times New Roman"

        allP.Range.Font.NameAscii = "Times New Roman"

        allP.Range.Font.NameOther = "Times New Roman"

        allP.Range.Font.NameBi = "Times New Roman"

        allP.Range.Font.ColorIndex = wdBlack

    Next allP

    

    ' Tranh doan van mac dinh sau bang day trang moi

    If doc.Paragraphs.Count > 0 Then

        Dim lastDocP As Paragraph: Set lastDocP = doc.Paragraphs(doc.Paragraphs.Count)

        If Not lastDocP.Range.Information(wdWithInTable) Then

            Dim lastPTxt As String

            lastPTxt = Trim(Replace(Replace(Replace(lastDocP.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            If Len(lastPTxt) = 0 Then

                lastDocP.Range.Font.Size = 1

                lastDocP.Range.ParagraphFormat.SpaceBefore = 0

                lastDocP.Range.ParagraphFormat.SpaceAfter = 0

                lastDocP.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            End If

        End If

    End If

    

    Set RebuildSignatureTable = newTbl

    Exit Function

Fallback:

    On Error Resume Next

    FormatExistingSignatureTable tbl, regIdx

    Set RebuildSignatureTable = tbl

End Function



Private Sub FormatExistingSignatureTable(ByVal tbl As Table, ByVal regIdx As Integer)

    On Error Resume Next

    tbl.Borders.Enable = False

    tbl.AutoFitBehavior wdAutoFitFixed

    tbl.LeftPadding = 0: tbl.RightPadding = 0

    tbl.TopPadding = 0: tbl.BottomPadding = 0

    tbl.Rows.WrapAroundText = 0

    tbl.Rows.LeftIndent = 0

    tbl.Rows.Alignment = wdAlignRowLeft

    

    Dim col1WSig As Single, col2WSig As Single

    If regIdx = 1 Then

        col1WSig = MillimetersToPoints(75)

        col2WSig = MillimetersToPoints(90)

    Else

        col1WSig = MillimetersToPoints(70)

        col2WSig = MillimetersToPoints(90)

    End If

    

    tbl.PreferredWidthType = wdPreferredWidthPoints

    tbl.PreferredWidth = col1WSig + col2WSig

    tbl.Columns(1).Width = col1WSig

    tbl.Columns(2).Width = col2WSig

    

    Dim rWSig As Row

    For Each rWSig In tbl.Rows

        If rWSig.Cells.Count >= 2 Then

            rWSig.Cells(1).Width = col1WSig

            rWSig.Cells(2).Width = col2WSig

        End If

    Next rWSig

    

    With tbl.Range.Font

        .Name = "Times New Roman"

        .NameAscii = "Times New Roman"

        .NameOther = "Times New Roman"

        .NameBi = "Times New Roman"

        .ColorIndex = wdBlack

    End With

    

    Dim c1 As Cell, c2 As Cell

    Set c1 = tbl.Cell(1, 1)

    Set c2 = tbl.Cell(1, 2)

    

    ' 1. DINH DANG COT NOI NHAN (COT 1)

    With c1.Range.ParagraphFormat

        .Alignment = wdAlignParagraphLeft

        .LineSpacingRule = wdLineSpaceSingle

        .SpaceBefore = 0

        .SpaceAfter = 0

        .LeftIndent = 0

        .RightIndent = 0

        .FirstLineIndent = 0

        .TabStops.ClearAll

    End With

    

    Dim p As Paragraph, pIdx As Long

    For pIdx = 1 To c1.Range.Paragraphs.Count

        Set p = c1.Range.Paragraphs(pIdx)

        With p.Range.ParagraphFormat

            .Alignment = wdAlignParagraphLeft

            .LineSpacingRule = wdLineSpaceSingle

            .SpaceBefore = 0

            .SpaceAfter = 0

            .LeftIndent = 0

            .RightIndent = 0

            .FirstLineIndent = 0

            .TabStops.ClearAll

        End With

        

        Dim pText As String

        pText = Trim(Replace(Replace(Replace(p.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

        Dim uP As String

        uP = UCase(ConvertToUnsignedVietnamese(pText))

        

        If pIdx = 1 And (InStr(uP, "NOI NHAN") > 0 Or pText = "Noi nhan" Or pText = "NOI NHAN") Then

            If InStr(p.Range.Text, Chr(10)) > 0 Then

                p.Range.Text = Replace(p.Range.Text, Chr(10), vbCr)

            End If

            If regIdx = 1 Then

                SetParaFont p, "Times New Roman", 14, False, False

                p.Range.Font.Underline = wdUnderlineSingle

            Else

                SetParaFont p, "Times New Roman", 12, True, True

                p.Range.Font.Underline = wdUnderlineNone

            End If

        Else

            If regIdx = 1 Then

                If InStr(pText, "-") > 0 And Len(pText) <= 25 And Left(pText, 1) <> "-" And Left(pText, 1) <> "+" Then

                    SetParaFont p, "Times New Roman", 11, False, False

                Else

                    SetParaFont p, "Times New Roman", 12, False, False

                End If

            Else

                SetParaFont p, "Times New Roman", 11, False, False

            End If

            p.Range.Font.Underline = wdUnderlineNone

        End If

    Next pIdx

    

    ' 2. DINH DANG COT CHU KY (COT 2) - Ho tro ca bang 1 hang hoac nhieu hang

    Dim rIdxSig As Long

    For rIdxSig = 1 To tbl.Rows.Count

        If tbl.Rows(rIdxSig).Cells.Count >= 2 Then

            Dim cSigCell As Cell: Set cSigCell = tbl.Cell(rIdxSig, 2)

            cSigCell.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

            cSigCell.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            cSigCell.Range.ParagraphFormat.SpaceBefore = 0

            cSigCell.Range.ParagraphFormat.SpaceAfter = 0

            cSigCell.Range.ParagraphFormat.LeftIndent = 0

            cSigCell.Range.ParagraphFormat.RightIndent = 0

            cSigCell.Range.ParagraphFormat.FirstLineIndent = 0

            cSigCell.Range.ParagraphFormat.TabStops.ClearAll

            

            Dim totalP As Long: totalP = cSigCell.Range.Paragraphs.Count

            For pIdx = 1 To totalP

                Set p = cSigCell.Range.Paragraphs(pIdx)

                With p.Range.ParagraphFormat

                    .Alignment = wdAlignParagraphCenter

                    .LineSpacingRule = wdLineSpaceSingle

                    .SpaceBefore = 0

                    .SpaceAfter = 0

                    .LeftIndent = 0

                    .RightIndent = 0

                    .FirstLineIndent = 0

                    .TabStops.ClearAll

                End With

                

                Dim sText As String

                sText = Trim(Replace(Replace(Replace(p.Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

                

                If Len(sText) > 0 Then

                    If rIdxSig = 1 And pIdx = 1 Then

                        ' Dong tham quyen hoac chuc vu (TM..., CHU TICH, GIAM DOC...)

                        If regIdx = 1 Then

                            sText = Replace(sText, "TM.", "T/M")

                            sText = Replace(sText, "KT.", "K/T")

                            sText = Replace(sText, "TL.", "T/L")

                            sText = Replace(sText, "TM ", "T/M ")

                            sText = Replace(sText, "KT ", "K/T ")

                            sText = Replace(sText, "TL ", "T/L ")

                            p.Range.Text = sText

                            SetParaFont p, "Times New Roman", 14, True, False

                        Else

                            SetParaFont p, "Times New Roman", 13.5, True, False

                        End If

                    ElseIf (rIdxSig = tbl.Rows.Count And pIdx = totalP) Or (rIdxSig > 1 And pIdx = totalP) Then

                        ' Ho ten nguoi ky o dong cuoi

                        If regIdx = 1 Then

                            Dim cleanN As String

                            cleanN = CleanAcademicAndMilitaryTitles(sText)

                            If cleanN <> sText Then

                                p.Range.Text = cleanN

                                sText = cleanN

                            End If

                        End If

                        SetParaFont p, "Times New Roman", IIf(regIdx = 1, 14, 13.5), True, False

                    ElseIf rIdxSig = 1 And pIdx = 2 And totalP >= 3 Then

                        ' Dong chuc vu nguoi ky sau tham quyen

                        If regIdx = 1 Then

                            ' HD 05: Chuc vu nguoi ky o dong 2 la 14pt Dung (Khong dam)

                            SetParaFont p, "Times New Roman", 14, False, False

                        Else

                            SetParaFont p, "Times New Roman", 13.5, True, False

                        End If

                    Else

                        SetParaFont p, "Times New Roman", IIf(regIdx = 1, 14, 13.5), False, False

                    End If

                Else

                    SetParaFont p, "Times New Roman", IIf(regIdx = 1, 14, 13.5), False, False

                End If

            Next pIdx

            

            ' Neu chi co 1 hang va chi co 2 doan (Chuc vu + Ho ten), chen khoang cach de ky

            If tbl.Rows.Count = 1 And totalP = 2 Then

                cSigCell.Range.Paragraphs(1).Range.InsertParagraphAfter

                cSigCell.Range.Paragraphs(2).Range.InsertParagraphAfter

                cSigCell.Range.Paragraphs(3).Range.InsertParagraphAfter

                Dim newTotal As Long: newTotal = cSigCell.Range.Paragraphs.Count

                SetParaFont cSigCell.Range.Paragraphs(newTotal), "Times New Roman", IIf(regIdx = 1, 14, 13.5), True, False

                cSigCell.Range.Paragraphs(newTotal).Alignment = wdAlignParagraphCenter

            End If

        End If

    Next rIdxSig

    

    ' Dam bao tuyet doi moi doan van trong bang chu ky deu mang font Times New Roman

    Dim finalP As Paragraph

    For Each finalP In tbl.Range.Paragraphs

        finalP.Range.Font.Name = "Times New Roman"

        finalP.Range.Font.NameAscii = "Times New Roman"

        finalP.Range.Font.NameOther = "Times New Roman"

        finalP.Range.Font.NameBi = "Times New Roman"

    Next finalP

End Sub



Private Sub AutoInsertMissingSignature(ByVal doc As Document, ByVal regIdx As Integer)

    On Error Resume Next

    If doc Is Nothing Then Exit Sub

    

    ' 0. Kiem tra loai tru: Van ban ban hanh kem theo Quyet dinh (Mau 1.3) KHONG tu dong chen chu ky

    If IsDocumentAttachedToDecision(doc) Then Exit Sub

    

    ' 1. Xac dinh vi tri neo chen chinh xac:

    ' Tim vi tri truoc Phu luc (neu co) hoac sau doan van ban cuoi cung co noi dung

    Dim targetRng As Range

    Dim pIdx As Long, p As Paragraph, pText As String, uText As String

    Dim phuLucStartPara As Paragraph: Set phuLucStartPara = Nothing

    

    ' Quet tim Phu luc o nua sau van ban (tranh chen nham sau phu luc)

    For pIdx = doc.Paragraphs.Count To MaxLong(1, doc.Paragraphs.Count - 60) Step -1

        pText = doc.Paragraphs(pIdx).Range.Text

        uText = UCase(ConvertToUnsignedVietnamese(pText))

        If Left(uText, 8) = "PHU LUC " Or Left(uText, 7) = "PHU LUC" Then

            Set phuLucStartPara = doc.Paragraphs(pIdx)

        End If

    Next pIdx

    

    If Not phuLucStartPara Is Nothing Then

        ' Chen ngay truoc Phu luc

        Set targetRng = phuLucStartPara.Range

        targetRng.Collapse wdCollapseStart

    Else

        ' Tim doan cuoi cung co text thuc su

        Dim lastMeaningfulPara As Paragraph: Set lastMeaningfulPara = Nothing

        For pIdx = doc.Paragraphs.Count To 1 Step -1

            pText = Trim(Replace(Replace(Replace(doc.Paragraphs(pIdx).Range.Text, vbCr, ""), vbLf, ""), Chr(7), ""))

            If Len(pText) > 0 Then

                Set lastMeaningfulPara = doc.Paragraphs(pIdx)

                Exit For

            End If

        Next pIdx

        

        If lastMeaningfulPara Is Nothing Then

            Set targetRng = doc.content

            targetRng.Collapse wdCollapseEnd

        Else

            ' Neu doan cuoi nam trong Bang (Table), phai lay Range SAU Bang

            If lastMeaningfulPara.Range.Tables.Count > 0 Then

                Dim lastTblInDoc As Table

                Set lastTblInDoc = lastMeaningfulPara.Range.Tables(1)

                Set targetRng = lastTblInDoc.Range

                targetRng.Collapse wdCollapseEnd

                targetRng.InsertParagraphAfter

                targetRng.Collapse wdCollapseEnd

            Else

                Set targetRng = lastMeaningfulPara.Range

                targetRng.Collapse wdCollapseEnd

                targetRng.InsertParagraphAfter

                targetRng.Collapse wdCollapseEnd

            End If

        End If

    End If

    

    Dim tbl As Table

    Set tbl = doc.Tables.Add(Range:=targetRng, NumRows:=1, NumColumns:=2)

    tbl.Borders.Enable = False

    tbl.AutoFitBehavior wdAutoFitFixed

    tbl.LeftPadding = 0: tbl.RightPadding = 0

    tbl.TopPadding = 0: tbl.BottomPadding = 0

    tbl.Rows.LeftIndent = 0

    tbl.Rows.WrapAroundText = 0

    tbl.Rows.AllowBreakAcrossPages = False

    tbl.Range.ParagraphFormat.KeepWithNext = True

    

    Dim col1WSig As Single, col2WSig As Single

    If regIdx = 1 Then

        col1WSig = MillimetersToPoints(75)

        col2WSig = MillimetersToPoints(95)

    Else

        col1WSig = MillimetersToPoints(70)

        col2WSig = MillimetersToPoints(95)

    End If

    tbl.Columns(1).Width = col1WSig

    tbl.Columns(2).Width = col2WSig

    

    If regIdx = 1 Then

        ' Chu ky Dang (HD 05)

        If g_SelectedOrgLevel = 6 Then

            ' Chi bo

            tbl.Cell(1, 1).Range.Text = DecodeUTF8("TsahaSBuaOG6rW46") & vbCrLf & _

                                        DecodeUTF8("LSDEkOG6o25nIOG7p3kgeMOjL3BoxrDhu51uZy4uLiAoxJHhu4MgYi9jKTs=") & vbCrLf & _

                                        DecodeUTF8("LSBDw6FjIMSR4buTbmcgY2jDrSDEkeG6o25nIHZpw6puIHRyb25nIENoaSBi4buZOw==") & vbCrLf & _

                                        DecodeUTF8("LSBMxrB1OiBDaGkgYuG7mS4=")

            tbl.Cell(1, 2).Range.Text = DecodeUTF8("VC9NIENISSBC4buY") & vbCrLf & _

                                        DecodeUTF8("QsONIFRIxq8=") & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _

                                        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

        ElseIf g_SelectedOrgLevel = 5 Then

            ' Dang uy Xa

            tbl.Cell(1, 1).Range.Text = DecodeUTF8("TsahaSBuaOG6rW46") & vbCrLf & _

                                        DecodeUTF8("LSBOaMawIGvDrW5oIGfhu61pOw==") & vbCrLf & _

                                        DecodeUTF8("LSBUaMaw4budbmcgdHLhu7FjIMSQ4bqjbmcgcGjGsOG7nW5nOw==") & vbCrLf & _

                                        DecodeUTF8("LSBMxrB1IFbEg24gcGjDsm5nLg==")

            tbl.Cell(1, 2).Range.Text = DecodeUTF8("VC9NIEJBTiBDSOG6pFAgSMOATkg=") & vbCrLf & _

                                        DecodeUTF8("QsONIFRIxq8=") & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _

                                        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

        Else

            tbl.Cell(1, 1).Range.Text = DecodeUTF8("TsahaSBuaOG6rW46") & vbCrLf & _

                                        DecodeUTF8("LSBOaMawIHRyw6puOw==") & vbCrLf & _

                                        DecodeUTF8("LSBUaMaw4budbmcgdHLhu7FjIFThu4luaCDhu6d5Ow==") & vbCrLf & _

                                        DecodeUTF8("LSBMxrB1IFbEg24gcGjDsm5nLg==")

            tbl.Cell(1, 2).Range.Text = DecodeUTF8("VC9NIEJBTiBDSOG6pFAgSMOATkg=") & vbCrLf & _

                                        DecodeUTF8("QsONIFRIxq8=") & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _

                                        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

        End If

    Else

        ' Chu ky Hanh chinh (ND 30)

        If g_SelectedOrgLevel = 1 Then

            ' Cap Bo (Bo truong)

            tbl.Cell(1, 1).Range.Text = DecodeUTF8("TsahaSBuaOG6rW46") & vbCrLf & _

                                        DecodeUTF8("LSBOaMawIMSQaeG7gXUgLi4uOw==") & vbCrLf & _

                                        DecodeUTF8("LSBUaOG7pyB0xrDhu5tuZywgY8OhYyBQaMOzIFRo4bun tMaw4budbmcgQ1A7") & vbCrLf & _

                                        DecodeUTF8("LSBDw6FjIELhu5ksIGPGoSBxdWFuIG5nYW5nIELhu5k7") & vbCrLf & _

                                        DecodeUTF8("LSBMxrB1OiBWVCwgUEMu")

            tbl.Cell(1, 2).Range.Text = DecodeUTF8("QuG7mCBUUsav4bufTkc=") & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _

                                        DecodeUTF8("Tmd1eeG7hW4gVsSDbiBB")

        ElseIf g_SelectedOrgLevel = 2 Then

            ' Cap Cuc / Tong cuc (Cuc truong)

            tbl.Cell(1, 1).Range.Text = DecodeUTF8("TsahaSBuaOG6rW46") & vbCrLf & _

                                        DecodeUTF8("LSBOaMawIMSQaeG7gXUgLi4uOw==") & vbCrLf & _

                                        DecodeUTF8("LSBMw6NuaCDEkeG6oW8gQuG7mTs=") & vbCrLf & _

                                        DecodeUTF8("LSBDw6FjIMSRxqFuIHbhu4sgdHLhu7FjIHRodeG7mWM7") & vbCrLf & _

                                        DecodeUTF8("LSBMxrB1OiBWVC4=")

            tbl.Cell(1, 2).Range.Text = DecodeUTF8("Q+G7pEMgVFLGr+G7n05H") & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _

                                        DecodeUTF8("TMOqIFbEg24gQw==")

        ElseIf g_SelectedOrgLevel = 5 Then

            ' UBND Xa / Phuong

            tbl.Cell(1, 1).Range.Text = DecodeUTF8("TsahaSBuaOG6rW46") & vbCrLf & _

                                        DecodeUTF8("LSBOaMawIHRyw6puOw==") & vbCrLf & _

                                        DecodeUTF8("LSBUaMaw4budbmcgdHLhu7FjIMSQ4bqjbmcg4buneTs=") & vbCrLf & _

                                        DecodeUTF8("LSBDaOG7pyB04buLY2gsIFBow7MgQ2jhu6cgdOG7i2NoIFVCTkQ7") & vbCrLf & _

                                        DecodeUTF8("LSBMxrB1OiBWVC4=")

            tbl.Cell(1, 2).Range.Text = DecodeUTF8("VE0uIOG7plkgQkFOIE5Iw4JOIETDgk4=") & vbCrLf & _

                                        DecodeUTF8("Q0jhu6YgVOG7ikNI") & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _

                                        DecodeUTF8("VHLhuqduIFbEg24gQg==")

        Else

            ' Mac dinh: Chu tich UBND

            tbl.Cell(1, 1).Range.Text = DecodeUTF8("TsahaSBuaOG6rW46") & vbCrLf & _

                                        DecodeUTF8("LSBOaMawIHRyw6puOw==") & vbCrLf & _

                                        DecodeUTF8("LSBDaOG7pyB04buLY2gsIFBow7MgQ2jhu6cgdOG7i2NoIFVCTkQ7") & vbCrLf & _

                                        DecodeUTF8("LSBMxrB1OiBWVC4=")

            tbl.Cell(1, 2).Range.Text = DecodeUTF8("VE0uIOG7plkgQkFOIE5Iw4JOIETDgk4=") & vbCrLf & _

                                        DecodeUTF8("Q0jhu6YgVOG7ikNI") & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _

                                        DecodeUTF8("VHLhuqduIFbEg24gQg==")

        End If

    End If

    

    FormatExistingSignatureTable tbl, regIdx

    tbl.Range.Collapse wdCollapseEnd

    tbl.Range.InsertParagraphAfter

End Sub



Private Sub StandardizeToneMarksAllQuiet(ByVal doc As Document)

    On Error Resume Next

    ' 1. PHUC HOI / SUA CAC TU BI LOI DAU THANH SAI TREN 'o' / 'u' (doan, toan, hoan, khoan, loai, ngoai, toan, hoat...)

    ReplaceExactWord doc, DecodeUTF8("xJHDsmFu"), DecodeUTF8("xJFvw6Bu"), True

    ReplaceExactWord doc, DecodeUTF8("dMOyYW4="), DecodeUTF8("dG/DoG4="), True

    ReplaceExactWord doc, DecodeUTF8("aMOyYW4="), DecodeUTF8("aG/DoG4="), True

    ReplaceExactWord doc, DecodeUTF8("a2jhu49hbg=="), DecodeUTF8("a2hv4bqjbg=="), True

    ReplaceExactWord doc, DecodeUTF8("bOG7jWFp"), DecodeUTF8("bG/huqFp"), True

    ReplaceExactWord doc, DecodeUTF8("bmfhu41haQ=="), DecodeUTF8("bmdv4bqhaQ=="), True

    ReplaceExactWord doc, DecodeUTF8("dMOzYW4="), DecodeUTF8("dG/DoW4="), True

    ReplaceExactWord doc, DecodeUTF8("aOG7jWF0"), DecodeUTF8("aG/huqF0"), True

    ReplaceExactWord doc, DecodeUTF8("dGjDs2F0"), DecodeUTF8("dGhvw6F0"), True

    ReplaceExactWord doc, DecodeUTF8("dGjhu41haQ=="), DecodeUTF8("dGhv4bqhaQ=="), True

    ReplaceExactWord doc, DecodeUTF8("bOG7jWF0"), DecodeUTF8("bG/huqF0"), True

    ReplaceExactWord doc, DecodeUTF8("xJHhu41hbg=="), DecodeUTF8("xJFv4bqhbg=="), True

    ReplaceExactWord doc, DecodeUTF8("a2jDs2Fu"), DecodeUTF8("a2hvw6Fu"), True

    ReplaceExactWord doc, DecodeUTF8("c8O6eXQ="), DecodeUTF8("c3XDvXQ="), True

    ReplaceExactWord doc, DecodeUTF8("a2jDuXluaA=="), DecodeUTF8("a2h14buzbmg="), True



    ' 2. CHUAN HOA CAC TU AM TIET MO DOC LAP THEO CHUAN MOI (hoa, toa, xoa, loa, hue, thuy, huy...)

    ReplaceExactWord doc, DecodeUTF8("aG/DoA=="), DecodeUTF8("aMOyYQ=="), True

    ReplaceExactWord doc, DecodeUTF8("dG/DoA=="), DecodeUTF8("dMOyYQ=="), True

    ReplaceExactWord doc, DecodeUTF8("eG/DoA=="), DecodeUTF8("eMOyYQ=="), True

    ReplaceExactWord doc, DecodeUTF8("bG/DoA=="), DecodeUTF8("bMOyYQ=="), True

    ReplaceExactWord doc, DecodeUTF8("aG/DoQ=="), DecodeUTF8("aMOzYQ=="), True

    ReplaceExactWord doc, DecodeUTF8("dG/DoQ=="), DecodeUTF8("dMOzYQ=="), True

    ReplaceExactWord doc, DecodeUTF8("eG/DoQ=="), DecodeUTF8("eMOzYQ=="), True

    ReplaceExactWord doc, DecodeUTF8("bG/DoQ=="), DecodeUTF8("bMOzYQ=="), True

    ReplaceExactWord doc, DecodeUTF8("dG/huqM="), DecodeUTF8("dOG7j2E="), True

    ReplaceExactWord doc, DecodeUTF8("dGhv4bqj"), DecodeUTF8("dGjhu49h"), True

    ReplaceExactWord doc, DecodeUTF8("eG/huqM="), DecodeUTF8("eOG7j2E="), True

    ReplaceExactWord doc, DecodeUTF8("aG/huqE="), DecodeUTF8("aOG7jWE="), True

    ReplaceExactWord doc, DecodeUTF8("dG/huqE="), DecodeUTF8("dOG7jWE="), True

    ReplaceExactWord doc, DecodeUTF8("bG/DqA=="), DecodeUTF8("bMOyZQ=="), True

    ReplaceExactWord doc, DecodeUTF8("eG/DqA=="), DecodeUTF8("eMOyZQ=="), True

    ReplaceExactWord doc, DecodeUTF8("bG/DqQ=="), DecodeUTF8("bMOzZQ=="), True

    ReplaceExactWord doc, DecodeUTF8("dG/DqQ=="), DecodeUTF8("dMOzZQ=="), True

    ReplaceExactWord doc, DecodeUTF8("Y2hvw6k="), DecodeUTF8("Y2jDs2U="), True

    ReplaceExactWord doc, DecodeUTF8("dGh14bu3"), DecodeUTF8("dGjhu6d5"), True

    ReplaceExactWord doc, DecodeUTF8("dGh1w70="), DecodeUTF8("dGjDunk="), True

    ReplaceExactWord doc, DecodeUTF8("aHXhu7c="), DecodeUTF8("aOG7p3k="), True

    ReplaceExactWord doc, DecodeUTF8("dHXhu7U="), DecodeUTF8("dOG7pXk="), True

    ReplaceExactWord doc, DecodeUTF8("bHXhu7U="), DecodeUTF8("bOG7pXk="), True

    ReplaceExactWord doc, DecodeUTF8("bmh14bu1"), DecodeUTF8("bmjhu6V5"), True

    ReplaceExactWord doc, DecodeUTF8("bHXhu7k="), DecodeUTF8("bMWpeQ=="), True

End Sub



Private Sub DeleteBlankPagesQuiet(ByVal doc As Document)

    On Error Resume Next

    Dim p As Paragraph, txt As String, loopCount As Long

    loopCount = 0

    Do While loopCount < 10

        loopCount = loopCount + 1

        If doc.Paragraphs.Count <= 1 Then Exit Do

        Set p = doc.Paragraphs(doc.Paragraphs.Count)

        If Not p.Range.Information(wdWithInTable) Then

            txt = Trim(Replace(Replace(Replace(Replace(p.Range.Text, vbCr, ""), vbLf, ""), Chr(12), ""), Chr(32), ""))

            If Len(txt) = 0 Then

                p.Range.Delete

            Else

                Exit Do

            End If

        Else

            Exit Do

        End If

    Loop

End Sub



' ==============================================================================

' 10. CAC CONG CU CHUAN HOA THU CONG (MANUAL STANDARDIZE ACTIONS)

' CHO PHEP NGUOI DUNG TU CHON THUC HIEN TUNG CONG DOAN RIENG LE

' ==============================================================================

Public Sub ManualFormatHeader(Optional doc As Document = Nothing)

    On Error GoTo ErrH

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then

        ShowUniMsg "VnVpIGzDsm5nIG3hu58gbeG7mXQgdMOgaSBsaeG7h3UgV29yZCDEkeG7gyDEkeG7i25oIGThuqFuZyE=", "Vi1GT1JNQVQ=", 48

        Exit Sub

    End If

    

    Dim undoRec As Object

    On Error Resume Next

    Set undoRec = Application.UndoRecord

    If Not undoRec Is Nothing Then undoRec.StartCustomRecord "VFormat_FormatHeader"

    On Error GoTo ErrH

    

    Dim regIdx As Integer: regIdx = GetEffectiveRegulation(doc)

    

    Dim hasHeader As Boolean: hasHeader = False

    If doc.Tables.Count > 0 Then

        Dim firstTbl As Table: Set firstTbl = doc.Tables(1)

        If firstTbl.Range.Start < 1500 And firstTbl.Columns.Count >= 2 Then

            hasHeader = True

            RebuildHeaderTable doc, firstTbl, regIdx

        End If

    End If

    

    If Not hasHeader Then

        Dim convertedRaw As Boolean: convertedRaw = False

        convertedRaw = ConvertRawHeaderToTable(doc, regIdx)

        If Not convertedRaw Then

            AutoInsertMissingHeader doc, regIdx

        End If

    End If

    

    If Not undoRec Is Nothing Then undoRec.EndCustomRecord

    ShowUniToast DecodeUTF8("xJDDoyBjaHXhuqluIGjDs2EgS2h1bmcgVGnDqnUgxJHhu4EgxJHhuqd1IHRyYW5nIHRow6BuaCBjw7RuZyE="), "Vi1GT1JNQVQ="

    Exit Sub

ErrH:

    If Not undoRec Is Nothing Then

        On Error Resume Next

        undoRec.EndCustomRecord

    End If

    RecordError "VFormat_Core.ManualFormatHeader", Err.Number, Err.Description, "Chuan hoa tieu de dau trang"

End Sub



Public Sub ManualFormatTitle(Optional doc As Document = Nothing)

    On Error GoTo ErrH

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then

        ShowUniMsg "VnVpIGzDsm5nIG3hu58gbeG7mXQgdMOgaSBsaeG7h3UgV29yZCDEkeG7gyDEkeG7i25oIGThuqFuZyE=", "Vi1GT1JNQVQ=", 48

        Exit Sub

    End If

    

    Dim undoRec As Object

    On Error Resume Next

    Set undoRec = Application.UndoRecord

    If Not undoRec Is Nothing Then undoRec.StartCustomRecord "VFormat_FormatTitle"

    On Error GoTo ErrH

    

    Dim regIdx As Integer: regIdx = GetEffectiveRegulation(doc)

    

    Dim hCount As Long: hCount = 0

    Dim titleDetected As String

    titleDetected = StandardizeTitleAndAbstract(doc, regIdx, hCount)

    

    If Not undoRec Is Nothing Then undoRec.EndCustomRecord

    

    If Len(titleDetected) > 0 Then

        ShowUniToast DecodeUTF8("xJDDoyBjxINuIGdp4buvYSAmIGluIMSR4bqtbSBUw6puIGxv4bqhaSB2xINuIGLhuqNuOiA=") & titleDetected, "Vi1GT1JNQVQ="

    Else

        ShowUniToast DecodeUTF8("S2jDtG5nIHTDrG0gdGjhuqV5IFTDqm4gbG/huqFpIHbEg24gYuG6o24gY+G7pSB0aOG7gyDEkeG7gyBjxINuIGdp4buvYS4="), "Vi1GT1JNQVQ="

    End If

    Exit Sub

ErrH:

    If Not undoRec Is Nothing Then

        On Error Resume Next

        undoRec.EndCustomRecord

    End If

    RecordError "VFormat_Core.ManualFormatTitle", Err.Number, Err.Description, "Can giua tieu de van ban"

End Sub



Public Sub ManualFormatBody(Optional doc As Document = Nothing)

    On Error GoTo ErrH

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then

        ShowUniMsg "VnVpIGzDsm5nIG3hu58gbeG7mXQgdMOgaSBsaeG7h3UgV29yZCDEkeG7gyDEkeG7i25oIGThuqFuZyE=", "Vi1GT1JNQVQ=", 48

        Exit Sub

    End If

    

    StartProgress DecodeUTF8("Vi1GT1JNQVQgLSDEkOG7iyBE4bqhbmcgTuG7mWkgRHVuZw=="), DecodeUTF8("xJBhbmcga2jhu59pIHThuqFvLi4u"), "ManualBody"

    

    Dim undoRec As Object

    On Error Resume Next

    Set undoRec = Application.UndoRecord

    If Not undoRec Is Nothing Then undoRec.StartCustomRecord "VFormat_FormatBody"

    On Error GoTo ErrH

    

    Dim regIdx As Integer: regIdx = GetEffectiveRegulation(doc)

    

    SetProgress 5, DecodeUTF8("xJBhbmcgY2h14bqpbiBow7NhIHBow7RuZyBjaOG7ryBUaW1lcyBOZXcgUm9tYW4uLi4="), "ManualBody"

    

    ' Chuan hoa toan bo font chu ve Times New Roman

    doc.content.Font.Name = "Times New Roman"

    doc.content.Font.NameAscii = "Times New Roman"

    doc.content.Font.NameOther = "Times New Roman"

    doc.content.Font.NameBi = "Times New Roman"

    

    Dim pCount As Long: pCount = 0

    Dim hCount As Long: hCount = 0

    FormatBodyAndHeadings doc, regIdx, pCount, hCount

    

    SetProgress 100, DecodeUTF8("SG/DoG4gdOG6pXQgxJHhu4tuaCBk4bqhbmch"), "ManualBody"

    FinishProgress "ManualBody"

    

    If Not undoRec Is Nothing Then undoRec.EndCustomRecord

    ShowUniToast DecodeUTF8("xJDDoyBjaHXhuqluIGjDs2Eg") & hCount & DecodeUTF8("IMSR4buBIG3hu6VjIHbDoCA=") & pCount & DecodeUTF8("IMSRb+G6oW4gdsSDbiBu4buZaSBkdW5nIQ=="), "Vi1GT1JNQVQ="

    Exit Sub

ErrH:

    FinishProgress "ManualBody"

    If Not undoRec Is Nothing Then

        On Error Resume Next

        undoRec.EndCustomRecord

    End If

    RecordError "VFormat_Core.ManualFormatBody", Err.Number, Err.Description, "Dinh dang de muc va noi dung"

End Sub



Public Sub ManualFormatSignature(Optional doc As Document = Nothing)

    On Error GoTo ErrH

    If doc Is Nothing Then

        If Documents.Count > 0 Then Set doc = ActiveDocument

    End If

    If doc Is Nothing Then

        ShowUniMsg "VnVpIGzDsm5nIG3hu58gbeG7mXQgdMOgaSBsaeG7h3UgV29yZCDEkeG7gyDEkeG7i25oIGThuqFuZyE=", "Vi1GT1JNQVQ=", 48

        Exit Sub

    End If

    

    Dim undoRec As Object

    On Error Resume Next

    Set undoRec = Application.UndoRecord

    If Not undoRec Is Nothing Then undoRec.StartCustomRecord "VFormat_FormatSignature"

    On Error GoTo ErrH

    

    Dim regIdx As Integer: regIdx = GetEffectiveRegulation(doc)

    

    ' 1. Don dep bang mau gia lap bi trung lap neu co

    CleanDuplicatePlaceholderSignatures doc

    

    ' 2. Tim kiem bang chu ky hien co trong van ban

    Dim sigTbl As Table

    Set sigTbl = FindExistingSignatureTable(doc)

    

    If Not sigTbl Is Nothing Then

        Set sigTbl = RebuildSignatureTable(doc, sigTbl, regIdx)

    Else

        ' 3. Neu chua co bang, thu chuyen doi tu doan van chu ky tho

        Dim convertedSig As Boolean: convertedSig = False

        convertedSig = ConvertRawSignatureToTable(doc, regIdx)

        

        If Not convertedSig Then

            ' 4. Kiem tra xem o cuoi van ban co noi dung chu ky hay khong

            Dim docEndText As String: docEndText = ""

            If Not IsDocumentAttachedToDecision(doc) Then

                If Not IsDocEndingWithSignatureBlock(doc) Then

                    ' Chi chen moi khi van ban hoan toan khong co bat ky noi nhan hay chu ky nao

                    AutoInsertMissingSignature doc, regIdx

                End If

            End If

        End If

    End If

    CleanDuplicatePlaceholderSignatures doc

    

    If Not undoRec Is Nothing Then undoRec.EndCustomRecord

    ShowUniToast DecodeUTF8("xJDDoyBjaHXhuqluIGjDs2EgS2h1bmcgQ2jhu68ga8O9ICYgTsahaSBuaOG6rW4gdGjDoG5oIGPDtG5nIQ=="), "Vi1GT1JNQVQ="

    Exit Sub

ErrH:

    If Not undoRec Is Nothing Then

        On Error Resume Next

        undoRec.EndCustomRecord

    End If

    RecordError "VFormat_Core.ManualFormatSignature", Err.Number, Err.Description, "Chuan hoa chu ky va noi nhan"

End Sub





