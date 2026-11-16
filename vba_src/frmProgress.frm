Attribute VB_Name = "frmProgress"
Attribute VB_Base = "0{6169FB20-31A0-45BF-B005-8EEC5C88891C}{23BCD223-E8BE-4A67-84CB-B595B6225E0F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True
End Sub

Public Sub SetProgress(ByVal pct As Integer, ByVal msg As String)
    On Error Resume Next
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100
    lblStatus.Caption = msg
    lblPercent.Caption = "Ti" & ChrW(7871) & "n tr" & ChrW(236) & "nh: " & pct & "%"
    lblBarFill.Width = (338 * pct) / 100
    DoEvents
End Sub

