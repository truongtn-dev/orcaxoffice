Attribute VB_Name = "frmPageNumber"
Attribute VB_Base = "0{7782B5DE-780A-4646-BEA3-937D59291BC5}{8CF8C674-75EB-470F-B115-0DE9B5E927F7}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error Resume Next
    optFromStart.GroupName = "grpPageMode"
    optCustom.GroupName = "grpPageMode"
    optHeader.GroupName = "grpPagePosition"
    optFooter.GroupName = "grpPagePosition"
    optFromStart.Value = True
    chkHideFirst.Value = True
    optCustom.Value = False
    txtStartPage.Enabled = False
    txtStartNumber.Enabled = False
    txtStartPage.Text = "4"
    txtStartNumber.Text = "1"
    optHeader.Value = True
    optFooter.Value = False
    lblStartPage.ForeColor = &H94A3B8
    lblStartNumber.ForeColor = &H94A3B8
    txtStartPage.BackColor = &HF8FAFC
    txtStartNumber.BackColor = &HF8FAFC
End Sub

Private Sub optFromStart_Click()
    If optFromStart.Value Then
        chkHideFirst.Enabled = True
        txtStartPage.Enabled = False
        txtStartNumber.Enabled = False
        lblStartPage.ForeColor = &H94A3B8
        lblStartNumber.ForeColor = &H94A3B8
        txtStartPage.BackColor = &HF8FAFC
        txtStartNumber.BackColor = &HF8FAFC
    End If
End Sub

Private Sub optCustom_Click()
    If optCustom.Value Then
        chkHideFirst.Enabled = False
        txtStartPage.Enabled = True
        txtStartNumber.Enabled = True
        lblStartPage.ForeColor = &H0
        lblStartNumber.ForeColor = &H0
        txtStartPage.BackColor = &HFFFFFF
        txtStartNumber.BackColor = &HFFFFFF
        On Error Resume Next
        txtStartPage.SetFocus
    End If
End Sub

Private Sub btnApply_Click()
    On Error GoTo ErrH
    Dim fromPg As Integer, startNum As Integer
    Dim hideFirst As Boolean, isHdr As Boolean
    isHdr = optHeader.Value
    If optCustom.Value = True Then
        fromPg = Val(txtStartPage.Text)
        startNum = Val(txtStartNumber.Text)
        If fromPg < 1 Then fromPg = 1
        If startNum < 1 Then startNum = 1
        hideFirst = False
    Else
        fromPg = 1
        startNum = 1
        hideFirst = chkHideFirst.Value
    End If
    Me.Hide
    VFormat_Core.ApplyPageNumbering fromPg, startNum, hideFirst, isHdr
    Unload Me
    Exit Sub
ErrH:
    MsgBox "Loi khi danh so trang: " & Err.Description, vbCritical, "ORCAX-VB"
End Sub

Private Sub btnRemove_Click()
    On Error Resume Next
    Me.Hide
    VFormat_Core.RemoveAllPageNumbers
    Unload Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

