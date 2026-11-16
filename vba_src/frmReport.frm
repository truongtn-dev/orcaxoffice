Attribute VB_Name = "frmReport"
Attribute VB_Base = "0{959ED058-FBEC-484E-8350-AB244869A297}{90B471A9-D4C8-4D01-A3AC-1CE3BB9AB6AD}"
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
Public ResultChoice As Integer

Private Sub SetUnicodeTitle()
    On Error Resume Next
    #If VBA7 Then
        Dim h As LongPtr
        h = FindWindowW(CLngPtr(0), StrPtr(Me.Caption))
        If h = 0 Then h = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(Me.Caption))
        If h = 0 Then h = FindWindowW(StrPtr("ThunderXFrame"), StrPtr(Me.Caption))
        If h <> 0 Then
            DefWindowProcW h, WM_SETTEXT, CLngPtr(0), StrPtr(VFormat_Core.DecodeUTF8("Vi1GT1JNQVQgLSBL4bq+VCBRVeG6oiBDSFXhuqhOIEjDk0E="))
        End If
    #Else
        Dim h As Long
        h = FindWindowW(0&, StrPtr(Me.Caption))
        If h = 0 Then h = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(Me.Caption))
        If h = 0 Then h = FindWindowW(StrPtr("ThunderXFrame"), StrPtr(Me.Caption))
        If h <> 0 Then
            DefWindowProcW h, WM_SETTEXT, 0&, StrPtr(VFormat_Core.DecodeUTF8("Vi1GT1JNQVQgLSBL4bq+VCBRVeG6oiBDSFXhuqhOIEjDk0E="))
        End If
    #End If
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    SetUnicodeTitle
    lblTitle.Font.Name = "Arial"
    lblTitle.Caption = "K" & ChrW(7870) & "T QU" & ChrW(7842) & " CHU" & ChrW(7848) & "N H" & ChrW(211) & "A TH" & ChrW(7874) & " TH" & ChrW(7912) & "C (1-CLICK)"
    txtReport.Font.Name = "Arial"
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    SetUnicodeTitle
    lblTitle.Font.Name = "Arial"
    lblTitle.Caption = "K" & ChrW(7870) & "T QU" & ChrW(7842) & " CHU" & ChrW(7848) & "N H" & ChrW(211) & "A TH" & ChrW(7874) & " TH" & ChrW(7912) & "C (1-CLICK)"
End Sub

Private Sub btnKeep_Click()
    ResultChoice = 1
    Me.Hide
End Sub

Private Sub btnUndo_Click()
    ResultChoice = 2
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        ResultChoice = 1
    End If
End Sub

