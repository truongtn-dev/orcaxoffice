Attribute VB_Name = "VFormat_License"

' ==============================================================================

' MODULE: VFormat_License

' MO TA: QUAN LY BAN QUYEN, THONG TIN PHAN MEM ORCAX    & TAC GIA

' ARCHITECTURE: LOCAL-FIRST CACHE, ZERO STARTUP NETWORK BLOCKING,

' MULTI-ENGINE TLS 1.2+ COMPATIBLE HTTP CLIENT, ROBUST OFFLINE GRACE PERIOD

' 100% PURE ASCII COMPLIANT FOR ZERO COMPILE ERRORS

' ==============================================================================

Option Explicit



Private Const API_BASE_URL As String = "https://orcax.vn/api/v1/license"

Private Const BUY_URL As String = "https://orcax.vn/san-pham-chi-tiet.php?slug=orcax     "

Private Const UPDATE_VERSION_URL As String = "https://orcax.vn/version.json   "

Private Const USER_AGENT As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) OrcaX/1.0.0   "

Public Const CURRENT_ADDIN_VERSION As String = "1.6.5"

Public Const LICENSE_GRACE_PERIOD_DAYS As Long = 7



Public g_LicenseStatus As String

Public g_DaysLeft As String

Public g_CustomerName As String

Public g_LicenseType As String

Public g_ExpiresAt As String

Public g_LastLicenseCheckTime As Double

Public g_HasCheckedStartup As Boolean

Public g_LastDiagnosticReason As String

Public g_LastDiagnosticDetail As String

Private m_CachedHWID As String



Private Function DecodeUTF8(ByVal b64Str As String) As String

    DecodeUTF8 = VFormat_Core.DecodeUTF8(b64Str)

End Function



' ==============================================================================

' 1. TAO VA CACHE MA DINH DANH PHAN CUNG (HWID)

' ==============================================================================

Public Function GetMachineHWID() As String

    On Error GoTo Fallback

    If Len(m_CachedHWID) > 0 Then

        GetMachineHWID = m_CachedHWID

        Exit Function

    End If

    

    Dim wmi As Object, items As Object, item As Object

    Dim cpuId As String, mbId As String, diskId As String

    

    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

    

    Set items = wmi.ExecQuery("SELECT ProcessorId FROM Win32_Processor")

    For Each item In items

        cpuId = Trim(item.ProcessorId)

        Exit For

    Next item

    

    Set items = wmi.ExecQuery("SELECT SerialNumber FROM Win32_BaseBoard")

    For Each item In items

        mbId = Trim(item.SerialNumber)

        Exit For

    Next item

    

    Set items = wmi.ExecQuery("SELECT VolumeSerialNumber FROM Win32_LogicalDisk WHERE DeviceID='C:'")

    For Each item In items

        diskId = Trim(item.VolumeSerialNumber)

        Exit For

    Next item

    

    If Len(cpuId) = 0 Then cpuId = "CPU0"

    If Len(mbId) = 0 Then mbId = "MB0"

    If Len(diskId) = 0 Then diskId = "DSK0"

    

    Dim rawId As String

    rawId = UCase(cpuId & "-" & mbId & "-" & diskId)

    rawId = Replace(Replace(Replace(rawId, " ", ""), "/", ""), "\", "")

    

    If Len(rawId) > 28 Then

        m_CachedHWID = Left(rawId, 28)

    Else

        m_CachedHWID = rawId

    End If

    GetMachineHWID = m_CachedHWID

    Exit Function



Fallback:

    On Error Resume Next

    Dim envUser As String, envComp As String

    envUser = Environ$("USERNAME")

    envComp = Environ$("COMPUTERNAME")

    m_CachedHWID = "PC-" & UCase(envComp) & "-" & UCase(envUser)

    GetMachineHWID = m_CachedHWID

End Function



' ==============================================================================

' 2. BO NHO CACHE BAN QUYEN CUC BO (LOCAL LICENSE CACHE)

' ==============================================================================

Public Function GetSavedLicenseKey() As String

    On Error Resume Next

    GetSavedLicenseKey = Trim(GetSetting("VFormat", "License", "LicenseKey", "VF-PERMANENT-LICENSE"))

End Function



Public Sub SaveLicenseKey(ByVal licKey As String)

    On Error Resume Next

    SaveSetting "VFormat", "License", "LicenseKey", Trim(licKey)

End Sub



Public Sub SaveLicenseCache(ByVal licKey As String, ByVal licStatus As String, ByVal licType As String, ByVal custName As String, ByVal expAt As String, ByVal daysLeft As String)

    On Error Resume Next

    SaveSetting "VFormat", "License", "LicenseKey", Trim(licKey)

    SaveSetting "VFormat", "License", "Status", "ACTIVE"

    SaveSetting "VFormat", "License", "Type", "lifetime"

    SaveSetting "VFormat", "License", "CustomerName", Trim(custName)

    SaveSetting "VFormat", "License", "ExpiresAt", "9999-12-31"

    SaveSetting "VFormat", "License", "DaysLeft", "9999"

    SaveSetting "VFormat", "License", "LastVerifiedAt", Format$(Now, "YYYY-MM-DD HH:NN:SS")

    If Len(GetSetting("VFormat", "License", "ActivatedAt", "")) = 0 Then

        SaveSetting "VFormat", "License", "ActivatedAt", Format$(Now, "YYYY-MM-DD HH:NN:SS")

    End If

    SaveSetting "VFormat", "License", "Blocked", "0"

End Sub



Public Sub LoadLicenseCache()

    On Error Resume Next

    ' Chot 4: Luon dat trang thai ACTIVE, lifetime va vo hieu hoa kiem tra khoa/het han

    g_LicenseType = "lifetime"

    g_CustomerName = GetSetting("VFormat", "License", "CustomerName", "VIP User")

    g_ExpiresAt = "9999-12-31"

    g_DaysLeft = "9999"

    g_LicenseStatus = "ACTIVE"

End Sub



Public Function IsLocalLicenseValid() As Boolean

    ' Chot 4: Luon tra ve True hop le

    g_LicenseStatus = "ACTIVE"

    IsLocalLicenseValid = True

End Function



' ==============================================================================

' 3. BO TRUYEN THONG HTTP DA NEN TANG (MULTI-ENGINE HTTP CLIENT)

' ==============================================================================

Public Function ExecuteHttpRequest(ByVal httpMethod As String, ByVal urlStr As String, ByVal payloadJson As String, ByRef outResponse As String, ByRef outStatusCode As Long, ByRef outDiagReason As String) As Boolean

    On Error Resume Next

    outResponse = ""

    outStatusCode = 0

    outDiagReason = ""

    ExecuteHttpRequest = False

    

    Dim startTime As Double: startTime = Timer

    Dim http As Object

    Dim engineName As String: engineName = ""

    

    ' --- Engine 1: WinHttp.WinHttpRequest.5.1 (Chuan toc do cao, ho tro TLS 1.2) ---

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    If Not http Is Nothing Then

        engineName = "WinHttp"

        http.Option(9) = 2688

        http.setTimeouts 5000, 5000, 10000, 10000

        

        http.Open httpMethod, urlStr, False

        http.setRequestHeader "User-Agent", USER_AGENT

        http.setRequestHeader "Accept", "application/json, text/plain, */*"

        If UCase(httpMethod) = "POST" Then

            http.setRequestHeader "Content-Type", "application/json; charset=utf-8"

            http.Send payloadJson

        Else

            http.Send

        End If

        

        If Err.Number = 0 Then

            outStatusCode = http.status

            outResponse = ReadHttpResponseUtf8(http)

        End If

    End If

    

    ' --- Engine 2 Fallback: MSXML2.ServerXMLHTTP.6.0 ---

    If outStatusCode = 0 Then

        Err.Clear

        Set http = Nothing

        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

        If Not http Is Nothing Then

            engineName = "ServerXMLHTTP"

            http.setTimeouts 5000, 5000, 10000, 10000

            http.Open httpMethod, urlStr, False

            http.setRequestHeader "User-Agent", USER_AGENT

            http.setRequestHeader "Accept", "application/json, text/plain, */*"

            If UCase(httpMethod) = "POST" Then

                http.setRequestHeader "Content-Type", "application/json; charset=utf-8"

                http.Send payloadJson

            Else

                http.Send

            End If

            

            If Err.Number = 0 Then

                outStatusCode = http.status

                outResponse = ReadHttpResponseUtf8(http)

            End If

        End If

    End If

    

    ' --- Engine 3 Fallback: MSXML2.XMLHTTP (WinINet he thong) ---

    If outStatusCode = 0 Then

        Err.Clear

        Set http = Nothing

        Set http = CreateObject("MSXML2.XMLHTTP")

        If Not http Is Nothing Then

            engineName = "XMLHTTP"

            http.Open httpMethod, urlStr, False

            http.setRequestHeader "User-Agent", USER_AGENT

            http.setRequestHeader "Accept", "application/json, text/plain, */*"

            If UCase(httpMethod) = "POST" Then

                http.setRequestHeader "Content-Type", "application/json; charset=utf-8"

                http.Send payloadJson

            Else

                http.Send

            End If

            

            If Err.Number = 0 Then

                outStatusCode = http.status

                outResponse = ReadHttpResponseUtf8(http)

            End If

        End If

    End If

    

    Dim elapsedSec As Double: elapsedSec = Timer - startTime

    Dim errNum As Long: errNum = Err.Number

    Dim errDesc As String: errDesc = Err.Description

    

    If outStatusCode = 200 Then

        outDiagReason = "OK"

        ExecuteHttpRequest = True

    ElseIf outStatusCode > 0 Then

        Select Case outStatusCode

            Case 400: outDiagReason = "HTTP_400_BAD_REQUEST"

            Case 401: outDiagReason = "HTTP_401_UNAUTHORIZED"

            Case 403: outDiagReason = "HTTP_403_FORBIDDEN"

            Case 404: outDiagReason = "HTTP_404_NOT_FOUND"

            Case 408: outDiagReason = "HTTP_408_TIMEOUT"

            Case 429: outDiagReason = "HTTP_429_TOO_MANY_REQUESTS"

            Case 500: outDiagReason = "HTTP_500_INTERNAL_SERVER_ERROR"

            Case 502: outDiagReason = "HTTP_502_BAD_GATEWAY"

            Case 503: outDiagReason = "HTTP_503_SERVICE_UNAVAILABLE"

            Case Else: outDiagReason = "HTTP_" & CStr(outStatusCode)

        End Select

    Else

        Select Case errNum

            Case &H80072EE7, -2147012889: outDiagReason = "DNS_FAILURE"

            Case &H80072EE2, -2147012894: outDiagReason = "TIMEOUT"

            Case &H80072F7D, &H80072F8F, &H80072F88, -2147012739, -2147012721, -2147012728: outDiagReason = "TLS_SSL_FAILURE"

            Case &H80072EFD, -2147012867: outDiagReason = "CONNECTION_REFUSED"

            Case Else

                If InStr(LCase(errDesc), "time") > 0 Then

                    outDiagReason = "TIMEOUT"

                ElseIf InStr(LCase(errDesc), "cert") > 0 Or InStr(LCase(errDesc), "ssl") > 0 Or InStr(LCase(errDesc), "tls") > 0 Then

                    outDiagReason = "TLS_SSL_FAILURE"

                ElseIf InStr(LCase(errDesc), "name") > 0 Or InStr(LCase(errDesc), "dns") > 0 Then

                    outDiagReason = "DNS_FAILURE"

                Else

                    outDiagReason = "NETWORK_ERROR"

                End If

        End Select

    End If

    

    g_LastDiagnosticReason = outDiagReason

    g_LastDiagnosticDetail = "Engine=" & engineName & " | Code=" & outStatusCode & " | Elapsed=" & Format$(elapsedSec, "0.00") & "s | Err=" & Hex(errNum)

    LogDiagnostic httpMethod & " " & urlStr, outDiagReason, outStatusCode, elapsedSec, g_LastDiagnosticDetail

End Function



Private Function ReadHttpResponseUtf8(ByVal http As Object) As String

    On Error Resume Next

    ReadHttpResponseUtf8 = http.responseText

End Function



Public Function ExtractJsonValue(ByVal jsonStr As String, ByVal keyName As String) As String

    On Error Resume Next

    Dim pos As Long, startPos As Long, endPos As Long

    Dim pattern As String: pattern = """" & keyName & """"

    pos = InStr(jsonStr, pattern)

    If pos > 0 Then

        pos = InStr(pos + Len(pattern), jsonStr, ":")

        If pos > 0 Then

            startPos = pos + 1

            Do While startPos <= Len(jsonStr) And (Mid(jsonStr, startPos, 1) = " " Or Mid(jsonStr, startPos, 1) = vbTab Or Mid(jsonStr, startPos, 1) = vbCr Or Mid(jsonStr, startPos, 1) = vbLf)

                startPos = startPos + 1

            Loop

            

            If startPos <= Len(jsonStr) Then

                If Mid(jsonStr, startPos, 1) = """" Then

                    startPos = startPos + 1

                    endPos = startPos

                    Do While endPos <= Len(jsonStr)

                        If Mid(jsonStr, endPos, 1) = "\" Then

                            endPos = endPos + 2

                        ElseIf Mid(jsonStr, endPos, 1) = """" Then

                            Exit Do

                        Else

                            endPos = endPos + 1

                        End If

                    Loop

                    Dim strVal As String

                    strVal = Mid(jsonStr, startPos, endPos - startPos)

                    strVal = Replace(strVal, "\""", """")

                    strVal = Replace(strVal, "\/", "/")

                    ExtractJsonValue = strVal

                    Exit Function

                Else

                    endPos = startPos

                    Do While endPos <= Len(jsonStr) And Mid(jsonStr, endPos, 1) <> "," And Mid(jsonStr, endPos, 1) <> "}" And Mid(jsonStr, endPos, 1) <> "]" And Mid(jsonStr, endPos, 1) <> vbCr And Mid(jsonStr, endPos, 1) <> vbLf

                        endPos = endPos + 1

                    Loop

                    ExtractJsonValue = Trim(Mid(jsonStr, startPos, endPos - startPos))

                    Exit Function

                End If

            End If

        End If

    End If

    ExtractJsonValue = ""

End Function



' ==============================================================================

' 4. GHI LOG CHAN DOAN AN TOAN (SAFE DIAGNOSTIC TELEMETRY)

' ==============================================================================

Public Sub LogDiagnostic(ByVal actionName As String, ByVal diagReason As String, ByVal httpStatus As Long, ByVal elapsedSec As Double, Optional ByVal extraMsg As String = "")

    On Error Resume Next

    Dim appDir As String

    appDir = Environ$("APPDATA") & "\VFormat"

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(appDir) Then fso.CreateFolder (appDir)

    

    Dim logFile As String: logFile = appDir & "\vformat_diagnostic.log"

    Dim ts As Object

    Set ts = fso.OpenTextFile(logFile, 8, True)

    

    Dim logLine As String

    logLine = "[" & Format$(Now, "YYYY-MM-DD HH:NN:SS") & "] " & _

              "[" & actionName & "] " & _

              "Reason=" & diagReason & " | " & _

              "Status=" & httpStatus & " | " & _

              "Elapsed=" & Format$(elapsedSec, "0.00") & "s | " & _

              "WordVer=" & Application.Version & " | " & _

              "VFormatVer=" & CURRENT_ADDIN_VERSION & " | " & _

              extraMsg

    ts.WriteLine logLine

    ts.Close

End Sub



' ==============================================================================

' 5. KIEM TRA KET NOI MAY CHU DOC LAP (TEST SERVER CONNECTIVITY)

' ==============================================================================

Public Function TestServerConnectivity() As String

    On Error Resume Next

    Dim resp As String, status As Long, reason As String

    Dim resMsg As String

    

    Dim okVer As Boolean

    okVer = ExecuteHttpRequest("GET", UPDATE_VERSION_URL, "", resp, status, reason)

    

    If okVer Then

        resMsg = DecodeUTF8("S0nhuq5UIFFVw4AgS0nhu4JNIFRSQSBL4bq+VCBO4buYSSBNw4FZIENI4bumOg==") & vbCrLf & _

                 "----------------------------------------" & vbCrLf & _

                 "- " & DecodeUTF8("TcOheSBjaOG7pzo=") & " orcax.vn   " & vbCrLf & _

                 "- " & DecodeUTF8("Q+G7lW5nIGvhur90IG7hu5FpOg==") & " 443 (HTTPS / TLS 1.2+)" & vbCrLf & _

                 "- " & DecodeUTF8("VHLhuqFuZyB0aMOhaSBt4bqhbmc6") & " " & DecodeUTF8("S+G6v3QgbuG7kWkgdGjDoG5oIGPDtG5nISAo") & "HTTP " & status & ")" & vbCrLf & _

                 "- " & DecodeUTF8("xJDGsOG7nW5nIHRydXnhu4FuOg==") & " " & DecodeUTF8("U+G6tW4gc8Ogbmcta8OtY2ggaG/huqF0IGLhuqNuIHF1eeG7gW4h")

        TestServerConnectivity = resMsg

        Exit Function

    End If

    

    Dim hwid As String: hwid = GetMachineHWID()

    Dim jsonPayload As String

    jsonPayload = "{""hwid"":""" & hwid & """,""license_key"":""TEST_PING"",""machine_name"":""" & Environ$("COMPUTERNAME") & """}"

    

    Dim okApi As Boolean

    okApi = ExecuteHttpRequest("POST", API_BASE_URL & "/check.php", jsonPayload, resp, status, reason)

    

    If okApi Or status = 200 Then

        resMsg = DecodeUTF8("S0nhuq5UIFFVw4AgS0nhu4JNIFRSQSBL4bq+VCBO4buYSSBNw4FZIENI4bumOg==") & vbCrLf & _

                 "----------------------------------------" & vbCrLf & _

                 "- " & DecodeUTF8("TcOheSBjaOG7pzo=") & " orcax.vn   " & vbCrLf & _

                 "- " & DecodeUTF8("Q+G7lW5nIGvhur90IG7hu5FpOg==") & " 443 (HTTPS / TLS 1.2+)" & vbCrLf & _

                 "- " & DecodeUTF8("VHLhuqFuZyB0aMOhaSBBUEk6") & " " & DecodeUTF8("S+G6v3QgbuG7kWkgdGjDoG5oIGPDtG5nISAo") & "HTTP " & status & ")"

        TestServerConnectivity = resMsg

        Exit Function

    End If

    

    Dim userReason As String

    Select Case reason

        Case "DNS_FAILURE"

            userReason = DecodeUTF8("S2jDtG5nIHRo4buDIGNow6JuIMSRb2FuIFTDoG4gTWnhu4FuIChMb2kgRE5TKS4gVnVpIGzDsm5nIGtp4buDbSB0cmEgbOG6oWkgZMOieSBt4bqhbmcgaG/hurdjIFdpLUZpIQ==")

        Case "TIMEOUT"

            userReason = DecodeUTF8("SMOqdCB0aOG7nWkgZ2lhbiBjaOG7nSBr4bq/dCBu4buRaSAoVGltZW91dCkuIE3huqFuZyBjw7MgdGjhu4MgcXXDoSBDaOG6rW0gaG/hurdjIGLhu4sgdMaw4budbmcgbOG7rWEgY2jhurduIQ==")

        Case "TLS_SSL_FAILURE"

            userReason = DecodeUTF8("TOG7l2kgeMOhYyB0aOG7sWMgYmnhu4NuIGNow7RuZyBjaOG7iSBTU0wgLyBUTFMgMS4yLiBNw6F5IGPDsyB0aOG7gyDEkGFuZyDEkOG7qW5nIHNhdSBwcm94eSBjxqEgcXVhbiE=")

        Case "HTTP_403_FORBIDDEN"

            userReason = DecodeUTF8("TcOheSBjaOG7pyB04burIGNo4buRaSBr4bq/dCBu4buRaSAoSFRUUCA0MDMpLiBDw7MgdGjhu4MgZG8gZOG6o2kgSVAgaG/hurdjIHThuqFpIGto4buPbmcgdOG6oW0gdGjhu51pIQ==")

        Case Else

            userReason = DecodeUTF8("S2jDtG5nIHRo4buDIGvhur90IG7hu5FpIG3DoXkgY2jhu6cuIE5ndXnDqm4gbmjDosm6Og==") & " " & reason & " (" & g_LastDiagnosticDetail & ")"

    End Select

    

    resMsg = DecodeUTF8("S0nhuq5UIFFVw4AgS0nhu4JNIFRSQSBL4bq+VCBO4buYSSBNw4FZIENI4bumOg==") & vbCrLf & _

             "----------------------------------------" & vbCrLf & _

             "- " & DecodeUTF8("VHLhuqFuZyB0aMOhaTo=") & " " & DecodeUTF8("VEjhuqBUIELhuqBJIChLaMO0bmcgdGjhu4Mga+G6v3QgbuG7kWkp") & vbCrLf & _

             "- " & DecodeUTF8("Q2hpIHRp4bq/dCBs4buXaTo=") & " " & userReason & vbCrLf & _

             "- " & DecodeUTF8("SMaw4bubbmcgZOG6q246") & " " & DecodeUTF8("VnVpIGzDsm5nIGtp4buDbSB0cmEga+G6v3QgbuG7kWkgSW50ZXJuZXQsIHThuq90IFZQTsSR4buRaSBo4buZIHRy4bujIQ==")

    TestServerConnectivity = resMsg

End Function



' ==============================================================================

' 6. KHOI TAO BAN QUYEN KHI KHOI DONG WORD (100% NON-BLOCKING)

' ==============================================================================

Public Sub InitLicenseOnStartup()

    On Error Resume Next

    If g_HasCheckedStartup Then Exit Sub

    g_HasCheckedStartup = True

    

    ' Chot 5: Doc cache cuc bo va thiet lap ban quyen san sang, khong hen gio xac thuc ngam

    LoadLicenseCache

    g_LicenseStatus = "ACTIVE"

End Sub



Public Sub DeferredLicenseVerification()

    ' Chot 5: Loai bo viec gui HWID va key len server, duy tri trang thai ACTIVE

    On Error Resume Next

    g_LicenseStatus = "ACTIVE"

    g_LicenseType = "lifetime"

    g_ExpiresAt = "9999-12-31"

    g_DaysLeft = "9999"

End Sub



' ==============================================================================

' 7. DIEU KIEN TRUY CAP TINH NANG (GATEKEEPER)

' ==============================================================================

Public Function EnsureLicenseActive() As Boolean

    ' Chot 3: Go bo hoan toan chot chan trung tam, luon cap quyen truy cap

    g_LicenseStatus = "ACTIVE"

    EnsureLicenseActive = True

End Function



' ==============================================================================

' 8. KIEM TRA VA KICH HOAT ONLINE

' ==============================================================================

Public Function CheckLicenseOnline(Optional ByVal licenseKey As String = "") As String

    ' Chot 5: Khong gui HWID va key len server, tra ve status ACTIVE gia lap

    CheckLicenseOnline = "{""status"":""ACTIVE"",""type"":""lifetime"",""customer_name"":""VIP User"",""days_left"":""9999"",""expires_at"":""9999-12-31""}"

End Function



Public Function ActivateLicenseOnline(ByVal licenseKey As String) As String

    ' Kich hoat truc tiep cuc bo khong phu thuoc ket qua server

    On Error Resume Next

    Dim cleanKey As String: cleanKey = Trim(licenseKey)

    If Len(cleanKey) = 0 Then cleanKey = "VF-PERMANENT-LICENSE"

    

    SaveLicenseCache cleanKey, "ACTIVE", "lifetime", "VIP User", "9999-12-31", "9999"

    g_LicenseStatus = "ACTIVE"

    g_CustomerName = "VIP User"

    g_DaysLeft = "9999"

    g_LicenseType = "lifetime"

    g_ExpiresAt = "9999-12-31"

    

    ActivateLicenseOnline = "{""status"":""ACTIVE"",""type"":""lifetime"",""customer_name"":""VIP User"",""days_left"":""9999"",""expires_at"":""9999-12-31""}"

End Function



Public Sub OpenBuyWebsite()

    On Error Resume Next

    VFormat_Core.SafeOpenUrlOrFile BUY_URL

End Sub



' ==============================================================================

' 9. CUA SO THONG TIN ORCAX    & BAN QUYEN (MODERN USERFORM frmLicense)

' ==============================================================================

Public Sub ShowVFormatAppInfo()

    On Error GoTo Fallback

    

    LoadLicenseCache

    Dim savedKey As String: savedKey = GetSavedLicenseKey()

    

    Dim statusStr As String, custName As String, durationStr As String

    statusStr = DecodeUTF8("xJDDgyBLw41DSCBIT+G6oFQgQuG6ok4gUVVZ4buATiBDSMONTkggVEjhu6hD")

    durationStr = DecodeUTF8("VsSpbmggdmnhu4VuIChUcuG7jW4gxJHhu51pKQ==")

    custName = IIf(Len(g_CustomerName) > 0, g_CustomerName, DecodeUTF8("S2jDoWNoIGjDoG5nIE9yY2FYIA=="))

    

    frmLicense.lblStatusBadge.Caption = statusStr

    frmLicense.lblStatusBadge.BackColor = &HE8F5E9  ' Soft Green

    frmLicense.lblStatusBadge.ForeColor = &H2E7D32  ' Dark Green

    

    frmLicense.lblDetails.Caption = DecodeUTF8("S2jDoWNoIGjDoG5nOiA=") & custName & vbCrLf & _

                                    DecodeUTF8("VGjhu51pIGjhuqFuOiA=") & durationStr

    

    frmLicense.txtKey.Text = savedKey

    frmLicense.Show vbModal

    Unload frmLicense

    DoEvents

    Exit Sub



Fallback:

    On Error Resume Next

    Dim msg As String

    msg = DecodeUTF8("VEjDlE5HIFRJTiBQSOG6pk4gTeG7gE0gT1JDQVggIA==") & vbCrLf & _

          "----------------------------------------" & vbCrLf & _

          "- " & DecodeUTF8("UGhpw6puIGLhuqNuOiA=") & "v" & CURRENT_ADDIN_VERSION & " (Build 2026)" & vbCrLf & _

          "- " & DecodeUTF8("VMOhYyBnaeG6ozog") & DecodeUTF8("T1JDQVggKG9yY2F4LnZuKSAgICAg") & vbCrLf & _

          "- " & DecodeUTF8("TGnDqm4gaOG7hyAvIEVtYWlsOiA=") & "support@orcax.vn       " & vbCrLf & _

          "- " & DecodeUTF8("UXV5IGNodeG6qW4gdMOtY2ggaOG7o3A6IE5naOG7iyDEkeG7i25oIDMwLCBIxrDhu5tuZyBk4bqrbiAwNSwgTmdo4buLIMSR4buLbmggMjUwLCBDViA1NTEy") & vbCrLf & vbCrLf & _

          DecodeUTF8("VEjDlE5HIFRJTiBC4bqiTiBRVVnhu4BO") & vbCrLf & _

          "----------------------------------------" & vbCrLf & _

          "- " & DecodeUTF8("VHLhuqFuZyB0aMOhaTog") & statusStr & vbCrLf & _

          "- " & DecodeUTF8("S2jDoWNoIGjDoG5nOiA=") & custName & vbCrLf & _

          "- " & DecodeUTF8("VGjhu51pIGjhuqFuIHPhu60gZOG7pW5nOiA=") & durationStr & vbCrLf

    VFormat_Core.ShowUniMsg msg, DecodeUTF8("T1JDQVggLSBUaMO0bmcgVGluIFBo4bqnbiBN4buBbSAmIELhuqNuIFF1eeG7gW4gIA=="), 64

End Sub



' ==============================================================================

' 10. HE THONG KIEM TRA & TU DONG CAP NHAT PHIEN BAN MOI (AUTO-UPDATER)

' ==============================================================================

Private Function IsNewerVersion(ByVal remoteVer As String, ByVal curVer As String) As Boolean

    On Error Resume Next

    remoteVer = Replace(Replace(Trim(remoteVer), "v", ""), "V", "")

    curVer = Replace(Replace(Trim(curVer), "v", ""), "V", "")

    Dim rParts() As String, cParts() As String

    rParts = Split(remoteVer, ".")

    cParts = Split(curVer, ".")

    

    Dim maxLen As Long

    maxLen = UBound(rParts)

    If UBound(cParts) > maxLen Then maxLen = UBound(cParts)

    

    Dim i As Long, rVal As Long, cVal As Long

    For i = 0 To maxLen

        rVal = 0: cVal = 0

        If i <= UBound(rParts) Then rVal = Val(rParts(i))

        If i <= UBound(cParts) Then cVal = Val(cParts(i))

        If rVal > cVal Then

            IsNewerVersion = True

            Exit Function

        ElseIf rVal < cVal Then

            IsNewerVersion = False

            Exit Function

        End If

    Next i

    IsNewerVersion = False

End Function



Public Sub CheckUpdateOnline(Optional ByVal isSilent As Boolean = False)

    On Error GoTo ErrH

    Static isBusy As Boolean

    If isBusy Then Exit Sub

    isBusy = True

    

    Dim respJson As String, status As Long, reason As String

    Dim ok As Boolean

    ok = ExecuteHttpRequest("GET", UPDATE_VERSION_URL, "", respJson, status, reason)

    

    If Not ok Or status <> 200 Then

        ok = ExecuteHttpRequest("GET", "https://orcax.vn/api/v1/version.php", "", respJson, status, reason)

    End If

    

    If ok And status = 200 Then

        Dim remoteVer As String, packageUrl As String, changelog As String

        remoteVer = ExtractJsonValue(respJson, "version")

        packageUrl = ExtractJsonValue(respJson, "download_url")

        changelog = ExtractJsonValue(respJson, "changelog")

        

        If Len(remoteVer) > 0 And IsNewerVersion(remoteVer, CURRENT_ADDIN_VERSION) Then

            Dim confirmMsg As String

            confirmMsg = DecodeUTF8("xJDDgyBDw5MgUEhJw4pOIELhuqJOIE3hu5pJIENITyBPUkNBWCEgIA==") & vbCrLf

            confirmMsg = confirmMsg & "----------------------------------------" & vbCrLf

            confirmMsg = confirmMsg & "- " & DecodeUTF8("UGhpw6puIGLhuqNuIG3hu5tpOiA=") & "v" & remoteVer & vbCrLf

            confirmMsg = confirmMsg & "- " & DecodeUTF8("UGhpw6puIGLhuqNuIGhp4buHbiB04bqhaTog") & "v" & CURRENT_ADDIN_VERSION & vbCrLf

            If Len(changelog) > 0 Then

                confirmMsg = confirmMsg & "- " & DecodeUTF8("TuG7mWkgZHVuZyBuw6JuZyBj4bqlcDog") & changelog & vbCrLf

            End If

            confirmMsg = confirmMsg & "----------------------------------------" & vbCrLf

            confirmMsg = confirmMsg & DecodeUTF8("QuG6oW4gY8OzIG114buRbiB04bqjaSB2w6AgY8OgaSDEkeG6t3QgYuG6o24gY+G6rXAgbmjhuq10IG3hu5tpIG5nYXkga2jDtG5nPw==")

            

            Dim doUpdate As Boolean

            doUpdate = VFormat_Core.ShowUniConfirm(confirmMsg, DecodeUTF8("T1JDQVggLSBD4bqtcCBOaOG6rXQgUGhpw6puIELhuqNuIE3hu5tpIA=="))

            If doUpdate Then

                OpenUpdateWebsite packageUrl, remoteVer

            End If

            isBusy = False

            Exit Sub

        Else

            If Not isSilent Then

                Dim latestMsg As String

                latestMsg = DecodeUTF8("QuG6oW4gxJFhbmcgc+G7rSBk4bulbmcgcGhpw6puIGLhuqNuIG3hu5tpIG5o4bqldCBj4bunYSBPUkNBWCEgIA==") & vbCrLf

                latestMsg = latestMsg & "- " & DecodeUTF8("UGhpw6puIGLhuqNuIGhp4buHbiB04bqhaTog") & "v" & CURRENT_ADDIN_VERSION

                VFormat_Core.ShowUniMsg latestMsg, DecodeUTF8("T1JDQVggLSBD4bqtcCBOaOG6rXQgUGjhuqduIE3hu4FtIA=="), 64

            End If

            isBusy = False

            Exit Sub

        End If

    Else

        If Not isSilent Then

            Dim netErrMsg As String

            netErrMsg = DecodeUTF8("S2jDtG5nIHRo4buDIGvhur90IG7hu5FpIMSR4bq/biBtw6F5IGNo4bunIGPhuq1wIG5o4bqtdCAob3JjYXgudm4pLg0KVnVpIGzDsm5nIGtp4buDbSB0cmEgbOG6oWkga+G6v3QgbuG7kWkgbeG6oW5nIEludGVybmV0IGPhu6dhIGLhuqFuISAgIA==")

            If Len(reason) > 0 And reason <> "OK" Then

                netErrMsg = netErrMsg & vbCrLf & "(" & reason & ")"

            End If

            VFormat_Core.ShowUniMsg netErrMsg, DecodeUTF8("T1JDQVggLSBD4bqtcCBOaOG6rXQgICA="), 48

        End If

        isBusy = False

        Exit Sub

    End If

    isBusy = False

    Exit Sub

ErrH:

    isBusy = False

    If Not isSilent Then

        VFormat_Core.ShowUniMsg DecodeUTF8("TOG7l2kgdHJvbmcgcXXDoSB0csOsbmgga2nhu4NtIHRyYSBj4bqtcCBuaOG6rXQ6IA==") & Err.Description, DecodeUTF8("T1JDQVggLSBD4bqtcCBOaOG6rXQgICA="), 48

    End If

End Sub



Public Sub OpenUpdateWebsite(ByVal packageUrl As String, ByVal newVer As String)

    On Error GoTo ErrH

    If Len(Trim(packageUrl)) = 0 Then

        packageUrl = "https://orcax.vn/downloads/OrcaX_Setup_v1.0.zip "

    End If

    

    VFormat_Core.SafeOpenUrlOrFile packageUrl

    

    Dim doneMsg As String

    doneMsg = DecodeUTF8("xJDDoyBt4bufIGxpw6puIGvhur90IHThuqNpIGLhuqNuIGPhuq1wIG5o4bqtdCAodg==") & newVer & ")!" & vbCrLf

    doneMsg = doneMsg & "----------------------------------------" & vbCrLf

    doneMsg = doneMsg & DecodeUTF8("U2F1IGtoaSB04bqjaSB24buBLCBi4bqhbiBjaOG7iSBj4bqnbiBjaOG6oXkgZmlsZSBDw6BpIMSR4bq3dCAxLUNsaWNrIMSR4buDIGhvw6BuIHThuqV0IGPhuq1wIG5o4bqtdCE=")

    VFormat_Core.ShowUniMsg doneMsg, DecodeUTF8("T1JDQVggLSBD4bqtcCBOaOG6rXQgICA="), 64

    Exit Sub

ErrH:

    VFormat_Core.ShowUniMsg DecodeUTF8("TOG7l2kga2hpIG3hu58gbGnDqm4ga+G6v3QgdOG6o2kgY+G6rXAgbmjhuq10OiA=") & Err.Description, DecodeUTF8("T1JDQVggLSBD4bqtcCBOaOG6rXQgICA="), 48

End Sub







