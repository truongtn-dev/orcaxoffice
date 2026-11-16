Attribute VB_Name = "frmLicense"

Attribute VB_Base = "0{3F021434-6177-4F78-A1F8-138A67DAFD10}{25B827C3-2A45-432B-8A54-768AACE76C2B}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = False

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = False

Option Explicit



#If VBA7 Then

    Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr

    Private Declare PtrSafe Function DefWindowProcW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

#Else

    Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long

    Private Declare Function DefWindowProcW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

#End If



Private Const WM_SETTEXT As Long = &HC



Private Sub SetUnicodeTitle()

    On Error Resume Next

    #If VBA7 Then

        Dim h As LongPtr

        h = FindWindowW(CLngPtr(0), StrPtr(Me.Caption))

        If h = 0 Then h = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(Me.Caption))

        If h = 0 Then h = FindWindowW(StrPtr("ThunderXFrame"), StrPtr(Me.Caption))

        If h <> 0 Then

            DefWindowProcW h, WM_SETTEXT, CLngPtr(0), StrPtr(VFormat_Core.DecodeUTF8("T1JDQVggT0ZGSUNFIFNVSVRFIC0gVEjDlE5HIFRJTiAmIELhuqJOIFFVWeG7gE4="))

        End If

    #Else

        Dim h As Long

        h = FindWindowW(0&, StrPtr(Me.Caption))

        If h = 0 Then h = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(Me.Caption))

        If h = 0 Then h = FindWindowW(StrPtr("ThunderXFrame"), StrPtr(Me.Caption))

        If h <> 0 Then

            DefWindowProcW h, WM_SETTEXT, 0&, StrPtr(VFormat_Core.DecodeUTF8("T1JDQVggT0ZGSUNFIFNVSVRFIC0gVEjDlE5HIFRJTiAmIELhuqJOIFFVWeG7gE4="))

        End If

    #End If

End Sub



Private Sub UserForm_Initialize()

    On Error Resume Next

    SetUnicodeTitle

    txtKey.Text = VFormat_License.GetSavedLicenseKey()

    If Len(VFormat_License.g_LastDiagnosticReason) > 0 Then

        lblDiagInfo.Caption = "Trang thai: " & VFormat_License.g_LastDiagnosticReason

    End If

End Sub



Private Sub UserForm_Activate()

    On Error Resume Next

    SetUnicodeTitle

End Sub



Private Sub btnClose_Click()

    Me.Hide

End Sub



Private Sub btnBuy_Click()

    VFormat_License.OpenBuyWebsite

End Sub



Private Sub btnTestConn_Click()

    On Error Resume Next

    lblDiagInfo.Caption = VFormat_Core.DecodeUTF8("xJBhbmcga2nhu4NtIHRyYSBr4bq/dCBu4buRaS4uLg==")

    DoEvents

    Dim res As String

    res = VFormat_License.TestServerConnectivity()

    lblDiagInfo.Caption = VFormat_License.g_LastDiagnosticReason

    VFormat_Core.ShowUniMsg res, VFormat_Core.DecodeUTF8("Vi1GT1JNQVQgLSBLaeG7g20gVHJhIEvhur90IE7hu5FpIE3DoXkgQ2jhu6c="), 64

End Sub



Private Sub btnActivate_Click()

    On Error Resume Next

    Dim k As String: k = Trim(txtKey.Text)

    k = Replace(k, vbCr, "")

    k = Replace(k, vbLf, "")

    k = Replace(k, vbTab, "")

    k = Replace(k, Chr(160), "")

    k = Trim(k)

    If Len(k) < 8 Then

        VFormat_Core.ShowUniMsg "Vui long nhap ma License Key hop le! (Vi du: VFYR-XXXX-XXXX-XXXX)", "T1JDQVggT0ZGSUNF", 48

        Exit Sub

    End If

    btnActivate.Caption = "Dang xu ly..."

    lblStatusBadge.Caption = VFormat_Core.DecodeUTF8("xJBhbmcga+G6v3QgbuG7kWkgbcOheSBjaOG7pyB4w6FjIHRo4buxYy4uLg==")

    lblStatusBadge.BackColor = &HE0F2FE

    lblStatusBadge.ForeColor = &H284C7

    DoEvents

    Dim actRes As String, actStatus As String

    actRes = VFormat_License.ActivateLicenseOnline(k)

    btnActivate.Caption = "KICH HOAT"

    If Len(actRes) = 0 Then

        lblDiagInfo.Caption = "Loi: " & VFormat_License.g_LastDiagnosticReason

        VFormat_Core.ShowUniMsg "Khong the ket noi may chu xac thuc ban quyen (" & VFormat_License.g_LastDiagnosticReason & "). Vui long kiem tra ket noi mang hoac bam 'Kiem tra ket noi'!", "T1JDQVggT0ZGSUNF", 48

        Exit Sub

    End If

    actStatus = UCase(VFormat_License.ExtractJsonValue(actRes, "status"))

    If actStatus = "ACTIVE" Then

        VFormat_Core.ShowUniMsg "xJDDoyBrw61jaCBob+G6oXQgYuG6o24gcXV54buBbiBWLUZPUk1BVCB0aMOgbmggY8O0bmchIEPhuqNtIMahbiBi4bqhbiDEkcOjIHRpbiBkw7luZyE=", "T1JDQVggT0ZGSUNF", 64

        Me.Hide

    ElseIf InStr(UCase(actRes), "MAX_DEVICES") > 0 Then

        VFormat_Core.ShowUniMsg "TcOjIGvDrWNoIGhv4bqhdCDEkcOjIMSR4bqhdCBnaeG7m2kgaOG6oW4gdGhp4bq/dCBi4buLIHThu5FpIMSRYSAoMSBtw6F5KS4gVnVpIGzDsm5nIGxpw6puIGjhu4cgdMOhYyBnaeG6oyDEkeG7gyDEkcaw4bujYyBo4buXIHRy4bujIQ==", "T1JDQVggT0ZGSUNF", 48

    ElseIf InStr(UCase(actRes), "EXPIRED") > 0 Then

        VFormat_Core.ShowUniMsg "TcOjIGLhuqNuIHF1eeG7gW4gbsOgeSDEkcOjIGjhur90IGjhuqFuIHPhu60gZOG7pW5nISBWdWkgbMOybmcgbXVhIGhv4bq3YyBnaWEgaOG6oW4ga2V5IG3hu5tpIMSR4buDIHRp4bq/cCB04bulYy4=", "T1JDQVggT0ZGSUNF", 48

    ElseIf InStr(UCase(actRes), "LOCK") > 0 Or InStr(UCase(actRes), "BLOCK") > 0 Or InStr(UCase(actRes), "REVOK") > 0 Then

        VFormat_Core.ShowUniMsg "QuG6o24gcXV54buBbiBuw6B5IMSRw6MgYuG7iyBraMOzYSBob+G6t2MgdOG6oW0gZOG7q25nIGLhu59pIFF14bqjbiB0cuG7iyB2acOqbiE=", "T1JDQVggT0ZGSUNF", 48

    ElseIf actStatus = "KEY_INVALID" Or InStr(UCase(actRes), "INVALID") > 0 Then

        VFormat_Core.ShowUniMsg "TcOjIGvDrWNoIGhv4bqhdCBraMO0bmcgY2jDrW5oIHjDoWMgaG/hurdjIGtow7RuZyB04buTbiB04bqhaS4gVnVpIGzDsm5nIGtp4buDbSB0cmEgbOG6oWkgdOG7q25nIGvDvSB04buxIQ==", "T1JDQVggT0ZGSUNF", 48

    Else

        VFormat_Core.ShowUniMsg "TcOjIGvDrWNoIGhv4bqhdCBraMO0bmcgaOG7o3AgbOG7hyEgVnVpIGzDsm5nIGtp4buDbSB0cmEgbOG6oWkgdOG7q25nIGvDvSB04buxLg==", "T1JDQVggT0ZGSUNF", 48

    End If

End Sub



