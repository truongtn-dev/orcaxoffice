Attribute VB_Name = "OrcaX_Dynamic"
Option Explicit

' ==============================================================================
' MODULE: OrcaX_Dynamic
' MO TA:   Dynamic Ribbon cho OrcaX Workbench v2.0
'   - 4 dynamicMenu doc du lieu THAT tu user_presets.json:
'       * Co quan Hanh chinh (agencies_admin)  -> InsertQuocHieuAdminIndex
'       * Co quan Dang      (agencies_party)   -> InsertQuocHieuPartyIndex
'       * Chu ky Lanh dao   (signers_list)     -> InsertSigSavedIndex
'       * Mau Noi nhan      (recipients_list)  -> InsertRecSavedIndex
'   - Wrapper: Chuan hoa im lang, mo thu muc Bieu mau, mo Huong dan
'   - Cache IRibbonUI + Ribbon_InvalidateAll (dong bo menu khi presets doi)
' QUY UOC: 100% ASCII source; tieng Viet chi qua VFormat_Core.DecodeUTF8(base64)
' ==============================================================================

Private Const CUI_NS As String = "http://schemas.microsoft.com/office/2009/07/customui"

' (Chua co du lieu - mo Cai dat Ca nhan hoa de thiet lap)
Private Const B64_EMPTY_PRESET As String = "KENoxrBhIGPDsyBk4buvIGxp4buHdSAtIG3hu58gQ8OhIG5ow6JuIGjDs2EgxJHhu4MgdGhp4bq/dCBs4bqtcCk="

' Danh sach section JSON co the xuat hien trong user_presets.json
' (dung de xac dinh bien section khi dem so luong item)
Private Const KNOWN_SECTIONS As String = "agencies_admin|agencies_party|recipients_list|signers_list|doc_info|footer_info"

Private m_Ribbon As IRibbonUI

' ------------------------------------------------------------------------------
' RIBBON CACHE & INVALIDATE
' ------------------------------------------------------------------------------
Public Sub CacheRibbon(ribbon As IRibbonUI)
    On Error Resume Next
    Set m_Ribbon = ribbon
End Sub

Public Sub Ribbon_InvalidateAll()
    On Error Resume Next
    If Not m_Ribbon Is Nothing Then m_Ribbon.Invalidate
End Sub

' Goi sau khi presets luu thay doi de lam moi menu Soan nhanh
Public Sub RefreshDynamicMenus()
    Ribbon_InvalidateAll
End Sub

' ------------------------------------------------------------------------------
' DYNAMIC MENU CONTENT (getContent callbacks)
' ------------------------------------------------------------------------------
Public Sub Ribbon_GetAgenciesContent(control As IRibbonControl, ByRef returnedVal)
    returnedVal = BuildDynMenuXml("agencies_admin", "dynAgn", "Ribbon_InsertAgencyDyn")
End Sub

Public Sub Ribbon_GetPartyContent(control As IRibbonControl, ByRef returnedVal)
    returnedVal = BuildDynMenuXml("agencies_party", "dynPty", "Ribbon_InsertAgencyDyn")
End Sub

Public Sub Ribbon_GetSignersContent(control As IRibbonControl, ByRef returnedVal)
    returnedVal = BuildDynMenuXml("signers_list", "dynSig", "Ribbon_InsertSigDyn")
End Sub

Public Sub Ribbon_GetRecipientsContent(control As IRibbonControl, ByRef returnedVal)
    returnedVal = BuildDynMenuXml("recipients_list", "dynRec", "Ribbon_InsertRecDyn")
End Sub

Private Function BuildDynMenuXml(ByVal section As String, ByVal idPrefix As String, _
                                 ByVal actionName As String) As String
    On Error GoTo Fail
    Dim json As String
    json = VFormat_Core.GetUserPresetsJson()

    Dim n As Long
    n = DynSectionCount(json, section)

    Dim sb As String, i As Long, lbl As String
    sb = "<menu xmlns=""" & CUI_NS & """>"

    If n = 0 Then
        sb = sb & "<button id=""" & idPrefix & "_0"" label=""" & _
             XmlAttr(VFormat_Core.DecodeUTF8(B64_EMPTY_PRESET)) & _
             """ enabled=""false"" imageMso=""AccountSettings""/>"
    Else
        For i = 1 To n
            lbl = VFormat_Core.GetPresetDisplayLabel(json, section, i, idPrefix & " " & i)
            If Len(lbl) = 0 Then lbl = idPrefix & " " & i
            sb = sb & "<button id=""" & idPrefix & "_" & i & _
                 """ label=""" & XmlAttr(lbl) & _
                 """ onAction=""" & actionName & """ imageMso=""AccountSettings""/>"
        Next i
    End If

    sb = sb & "</menu>"
    BuildDynMenuXml = sb
    Exit Function
Fail:
    BuildDynMenuXml = "<menu xmlns=""" & CUI_NS & """><button id=""dynErr"" label=""(Data error)"" enabled=""false""/></menu>"
End Function

' Dem so item trong 1 section JSON, giu nghiem bien section
' (ExtractJsonArrayItemTag khong dung tai bien section nen khong dung truc tiep)
Private Function DynSectionCount(ByVal json As String, ByVal section As String) As Long
    On Error Resume Next
    Dim secPos As Long, marker As String
    marker = """" & section & """"
    secPos = InStr(1, json, marker)
    If secPos = 0 Then Exit Function

    ' Tim vi tri section ke tiep trong KNOWN_SECTIONS nam sau secPos
    Dim endPos As Long, probe As Long, s As Variant
    endPos = Len(json) + 1
    For Each s In Split(KNOWN_SECTIONS, "|")
        If CStr(s) <> section Then
            probe = InStr(secPos + Len(marker), json, """" & CStr(s) & """")
            If probe > 0 And probe < endPos Then endPos = probe
        End If
    Next s

    ' Dem so truong "id" trong doan [secPos, endPos)
    Dim slice As String, cnt As Long
    slice = Mid$(json, secPos, endPos - secPos)
    probe = 0
    Do
        probe = InStr(probe + 1, slice, """id""")
        If probe = 0 Then Exit Do
        cnt = cnt + 1
    Loop
    DynSectionCount = cnt
End Function

Private Function XmlAttr(ByVal s As String) As String
    On Error Resume Next
    Dim r As String
    r = s
    r = Replace(r, "&", "&amp;")
    r = Replace(r, "<", "&lt;")
    r = Replace(r, ">", "&gt;")
    r = Replace(r, """", "&quot;")
    XmlAttr = r
End Function

' ------------------------------------------------------------------------------
' DISPATCHERS - chen theo index doc tu control.Id ("dynAgn_3" -> idx 3)
' ------------------------------------------------------------------------------
Public Sub Ribbon_InsertAgencyDyn(control As IRibbonControl)
    On Error GoTo Fail
    Dim parts() As String
    parts = Split(control.Id, "_")
    Dim idx As Long
    idx = CLng(parts(UBound(parts)))
    If UCase$(Left$(parts(0), 6)) = "DYNPTY" Then
        VFormat_Core.InsertQuocHieuPartyIndex idx
    Else
        VFormat_Core.InsertQuocHieuAdminIndex idx
    End If
    Exit Sub
Fail:
    VFormat_Core.RecordError "OrcaX_Dynamic.Ribbon_InsertAgencyDyn", Err.Number, Err.Description
End Sub

Public Sub Ribbon_InsertSigDyn(control As IRibbonControl)
    On Error GoTo Fail
    Dim parts() As String
    parts = Split(control.Id, "_")
    VFormat_Core.InsertSigSavedIndex CLng(parts(UBound(parts)))
    Exit Sub
Fail:
    VFormat_Core.RecordError "OrcaX_Dynamic.Ribbon_InsertSigDyn", Err.Number, Err.Description
End Sub

Public Sub Ribbon_InsertRecDyn(control As IRibbonControl)
    On Error GoTo Fail
    Dim parts() As String
    parts = Split(control.Id, "_")
    VFormat_Core.InsertRecSavedIndex CLng(parts(UBound(parts)))
    Exit Sub
Fail:
    VFormat_Core.RecordError "OrcaX_Dynamic.Ribbon_InsertRecDyn", Err.Number, Err.Description
End Sub

' ------------------------------------------------------------------------------
' WRAPPERS MOI CHO RIBBON v2.0
' ------------------------------------------------------------------------------
' Chuan hoa im lang: khong hoi xac nhan, dung cho van ban don gian
Public Sub Ribbon_AutoStandardizeQuiet(control As IRibbonControl)
    On Error GoTo Fail
    If Documents.Count = 0 Then Exit Sub
    Dim doc As Document
    Set doc = ActiveDocument
    VFormat_Core.ExecuteAutoStandardize doc, True
    Exit Sub
Fail:
    VFormat_Core.RecordError "OrcaX_Dynamic.Ribbon_AutoStandardizeQuiet", Err.Number, Err.Description
End Sub

' Mo thu muc chua cac bieu mau da dong bo tren may nguoi dung
Public Sub Ribbon_OpenTemplatesFolder(control As IRibbonControl)
    On Error Resume Next
    VFormat_Core.EnsureVFormatAppDir
    VFormat_Core.SafeOpenUrlOrFile VFormat_Core.GetVFormatAppDir() & "\templates"
End Sub

' Mo huong dan su dung (uu tien file local, fallback website)
Public Sub Ribbon_OpenHelpGuide(Optional control As IRibbonControl = Nothing)
    On Error Resume Next
    Dim guidePath As String
    guidePath = VFormat_Core.GetVFormatAppDir() & "\HUONG_DAN_SU_DUNG.txt"
    If Len(Dir$(guidePath)) > 0 Then
        VFormat_Core.SafeOpenUrlOrFile guidePath
    Else
        VFormat_Core.SafeOpenUrlOrFile "https://orcax.vn/huong-dan"
    End If
End Sub
