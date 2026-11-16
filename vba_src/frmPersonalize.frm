Attribute VB_Name = "frmPersonalize"
Attribute VB_Base = "0{1F844FC2-192E-4B65-8B5A-4448328F01A3}{BA89A5A0-6E7F-4EA3-8C4B-F7F93216F579}"
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

Private m_Json As String
Private m_CurTab As Long
Private m_CurAdminIdx As Long
Private m_CurPartyIdx As Long
Private m_CurSignerIdx As Long
Private m_CurRecIdx As Long
Private m_IsLoading As Boolean

Private Sub SetUnicodeTitle()
    On Error Resume Next
    #If VBA7 Then
        Dim h As LongPtr
    #Else
        Dim h As Long
    #End If
    h = FindWindowW(0, StrPtr(Me.Caption))
    If h = 0 Then h = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(Me.Caption))
    If h = 0 Then h = FindWindowW(StrPtr("ThunderXFrame"), StrPtr(Me.Caption))
    If h <> 0 Then
        DefWindowProcW h, WM_SETTEXT, 0, StrPtr(VFormat_Core.DecodeUTF8("Vi1GT1JNQVQgLSBDw4BJIMSQ4bq2VCBDw4EgTkjDgk4gSMOTQSBExqogTGnhu4ZVIA=="))
    End If
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    m_IsLoading = True
    SetUnicodeTitle

    m_Json = VFormat_Core.GetUserPresetsJson()

    cboAdminSlot.Clear
    cboAdminSlot.AddItem "Mau 1 (Cap Bo / Mac dinh)"
    cboAdminSlot.AddItem "Mau 2 (Cuc truc thuoc Bo)"
    cboAdminSlot.AddItem "Mau 3 (Cap So / Nganh)"
    cboAdminSlot.AddItem "Mau 4 (Cap Xa / Phuong)"
    cboAdminSlot.AddItem "Mau 5 (Cap Dac khu)"
    cboAdminSlot.AddItem "Mau 6 (Van phong UBND)"

    cboPartySlot.Clear
    cboPartySlot.AddItem "Mau 1 (Mac dinh)"
    cboPartySlot.AddItem "Mau 2"
    cboPartySlot.AddItem "Mau 3"
    cboPartySlot.AddItem "Mau 4"

    cboSignerSlot.Clear
    cboSignerSlot.AddItem "Mau 1 (Bo truong)"
    cboSignerSlot.AddItem "Mau 2 (KT. Bo truong / Thu truong)"
    cboSignerSlot.AddItem "Mau 3 (Cuc truong)"
    cboSignerSlot.AddItem "Mau 4 (Giam doc So)"
    cboSignerSlot.AddItem "Mau 5 (Chu tich UBND)"
    cboSignerSlot.AddItem "Mau 6 (Bi thu Tinh uy)"
    cboSignerSlot.AddItem "Mau 7 (Bi thu Chi bo)"
    cboSignerSlot.AddItem "Mau 8 (Chanh Van phong)"

    cboRecSlot.Clear
    cboRecSlot.AddItem "Mau 1 (Cap Bo / Ngang Bo)"
    cboRecSlot.AddItem "Mau 2 (Cong van thong dung)"
    cboRecSlot.AddItem "Mau 3 (Quyet dinh ban hanh)"
    cboRecSlot.AddItem "Mau 4 (Chi bo Dang)"

    m_CurAdminIdx = 1
    cboAdminSlot.ListIndex = 0

    m_CurPartyIdx = 1
    cboPartySlot.ListIndex = 0

    m_CurSignerIdx = 1
    cboSignerSlot.ListIndex = 0

    m_CurRecIdx = 1
    cboRecSlot.ListIndex = 0

    m_IsLoading = False

    LoadAdminSlot 1
    LoadPartySlot 1
    LoadSignerSlot 1
    LoadRecSlot 1

    SwitchTab 1
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    SetUnicodeTitle
End Sub

Public Sub SwitchTab(ByVal tabIdx As Long)
    On Error Resume Next
    m_CurTab = tabIdx

    btnTabAdmin.BackColor = IIf(tabIdx = 1, &H822D1E, &HE2E8F0)
    btnTabAdmin.ForeColor = IIf(tabIdx = 1, &HFFFFFF, &H334155)
    btnTabAdmin.Font.Bold = (tabIdx = 1)

    btnTabParty.BackColor = IIf(tabIdx = 2, &H822D1E, &HE2E8F0)
    btnTabParty.ForeColor = IIf(tabIdx = 2, &HFFFFFF, &H334155)
    btnTabParty.Font.Bold = (tabIdx = 2)

    btnTabSigner.BackColor = IIf(tabIdx = 3, &H822D1E, &HE2E8F0)
    btnTabSigner.ForeColor = IIf(tabIdx = 3, &HFFFFFF, &H334155)
    btnTabSigner.Font.Bold = (tabIdx = 3)

    btnTabRec.BackColor = IIf(tabIdx = 4, &H822D1E, &HE2E8F0)
    btnTabRec.ForeColor = IIf(tabIdx = 4, &HFFFFFF, &H334155)
    btnTabRec.Font.Bold = (tabIdx = 4)

    fraAdmin.Visible = (tabIdx = 1)
    fraParty.Visible = (tabIdx = 2)
    fraSigner.Visible = (tabIdx = 3)
    fraRec.Visible = (tabIdx = 4)
End Sub

Private Sub btnTabAdmin_Click()
    SwitchTab 1
End Sub

Private Sub btnTabParty_Click()
    SwitchTab 2
End Sub

Private Sub btnTabSigner_Click()
    SwitchTab 3
End Sub

Private Sub btnTabRec_Click()
    SwitchTab 4
End Sub

Private Sub LoadAdminSlot(ByVal idx As Long)
    On Error Resume Next
    m_IsLoading = True
    txtAdminAlias.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "agencies_admin", "alias", idx)
    txtAdminAgency.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "agencies_admin", "agency_name", idx)
    txtAdminParent.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "agencies_admin", "parent_agency", idx)
    txtAdminLoc.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "agencies_admin", "location", idx)
    txtAdminPrefix.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "agencies_admin", "code_prefix", idx)
    m_IsLoading = False
End Sub

Private Sub SaveCurrentAdminSlot()
    On Error Resume Next
    If m_CurAdminIdx < 1 Or m_CurAdminIdx > 6 Then Exit Sub
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "agencies_admin", "alias", m_CurAdminIdx, txtAdminAlias.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "agencies_admin", "agency_name", m_CurAdminIdx, txtAdminAgency.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "agencies_admin", "parent_agency", m_CurAdminIdx, txtAdminParent.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "agencies_admin", "location", m_CurAdminIdx, txtAdminLoc.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "agencies_admin", "code_prefix", m_CurAdminIdx, txtAdminPrefix.Text)
End Sub

Private Sub cboAdminSlot_Change()
    If m_IsLoading Then Exit Sub
    SaveCurrentAdminSlot
    m_CurAdminIdx = cboAdminSlot.ListIndex + 1
    LoadAdminSlot m_CurAdminIdx
End Sub

Private Sub LoadPartySlot(ByVal idx As Long)
    On Error Resume Next
    m_IsLoading = True
    txtPartyAlias.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "agencies_party", "alias", idx)
    txtPartyName.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "agencies_party", "party_name", idx)
    txtPartyParent.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "agencies_party", "parent_party", idx)
    txtPartyLoc.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "agencies_party", "location", idx)
    txtPartyPrefix.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "agencies_party", "code_prefix", idx)
    m_IsLoading = False
End Sub

Private Sub SaveCurrentPartySlot()
    On Error Resume Next
    If m_CurPartyIdx < 1 Or m_CurPartyIdx > 4 Then Exit Sub
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "agencies_party", "alias", m_CurPartyIdx, txtPartyAlias.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "agencies_party", "party_name", m_CurPartyIdx, txtPartyName.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "agencies_party", "parent_party", m_CurPartyIdx, txtPartyParent.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "agencies_party", "location", m_CurPartyIdx, txtPartyLoc.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "agencies_party", "code_prefix", m_CurPartyIdx, txtPartyPrefix.Text)
End Sub

Private Sub cboPartySlot_Change()
    If m_IsLoading Then Exit Sub
    SaveCurrentPartySlot
    m_CurPartyIdx = cboPartySlot.ListIndex + 1
    LoadPartySlot m_CurPartyIdx
End Sub

Private Sub LoadSignerSlot(ByVal idx As Long)
    On Error Resume Next
    m_IsLoading = True
    txtSignerAlias.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "signers_list", "alias", idx)
    txtSignerTitle.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "signers_list", "title", idx)
    txtSignerPos.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "signers_list", "position", idx)
    txtSignerName.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "signers_list", "fullname", idx)
    txtSignerType.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "signers_list", "signer_type", idx)
    m_IsLoading = False
End Sub

Private Sub SaveCurrentSignerSlot()
    On Error Resume Next
    If m_CurSignerIdx < 1 Or m_CurSignerIdx > 8 Then Exit Sub
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "signers_list", "alias", m_CurSignerIdx, txtSignerAlias.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "signers_list", "title", m_CurSignerIdx, txtSignerTitle.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "signers_list", "position", m_CurSignerIdx, txtSignerPos.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "signers_list", "fullname", m_CurSignerIdx, txtSignerName.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "signers_list", "signer_type", m_CurSignerIdx, txtSignerType.Text)
End Sub

Private Sub cboSignerSlot_Change()
    If m_IsLoading Then Exit Sub
    SaveCurrentSignerSlot
    m_CurSignerIdx = cboSignerSlot.ListIndex + 1
    LoadSignerSlot m_CurSignerIdx
End Sub

Private Sub LoadRecSlot(ByVal idx As Long)
    On Error Resume Next
    m_IsLoading = True
    txtRecTitle.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "recipients_list", "title", idx)
    txtRecContent.Text = VFormat_Core.ExtractJsonArrayItemTag(m_Json, "recipients_list", "content", idx)
    m_IsLoading = False
End Sub

Private Sub SaveCurrentRecSlot()
    On Error Resume Next
    If m_CurRecIdx < 1 Or m_CurRecIdx > 4 Then Exit Sub
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "recipients_list", "title", m_CurRecIdx, txtRecTitle.Text)
    m_Json = VFormat_Core.UpdateJsonArrayItemTag(m_Json, "recipients_list", "content", m_CurRecIdx, txtRecContent.Text)
End Sub

Private Sub cboRecSlot_Change()
    If m_IsLoading Then Exit Sub
    SaveCurrentRecSlot
    m_CurRecIdx = cboRecSlot.ListIndex + 1
    LoadRecSlot m_CurRecIdx
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrH
    SaveCurrentAdminSlot
    SaveCurrentPartySlot
    SaveCurrentSignerSlot
    SaveCurrentRecSlot

    VFormat_Core.SaveUserPresetsJson m_Json
    VFormat_Core.ShowUniMsg "xJDDoyBsxrB1IGPDoGkgxJHhurd0IGPDoSBuaMOibiBow7NhIGThu68gTGnhu4d1IHRow6BuaCBjw7RuZyE=", "Vi1GT1JNQVQ=", 64
    Me.Hide
    Exit Sub
ErrH:
    VFormat_Core.ShowUniMsg "Loi khi luu cai dat: " & Err.Description, "Vi1GT1JNQVQ=", 48
End Sub

Private Sub btnReset_Click()
    On Error GoTo ErrH
    Dim cf As Boolean
    cf = VFormat_Core.ShowUniConfirm("QuG6oW4gY8OzIGNo4bqvYyBjaOG6r24gbeG7kW4ga2jDtGkgcGjhu6VjIGPhuqV1IGjDrG5oIG3hurdjIMSR4buLbmggYmFuIMSR4bqndSBraMO0bmc/", "Vi1GT1JNQVQ=")
    If cf Then
        m_Json = VFormat_Core.GetDefaultPresetsJson()
        VFormat_Core.SaveUserPresetsJson m_Json
        LoadAdminSlot m_CurAdminIdx
        LoadPartySlot m_CurPartyIdx
        LoadSignerSlot m_CurSignerIdx
        LoadRecSlot m_CurRecIdx
        VFormat_Core.ShowUniMsg "xJDDoyBow7RpIHBo4bulYyBj4bqldSBow6xuaCBt4bq3YyDEkeG7i25oIGJhbiDEkeG6p3UgdGjDoG5oIGPDtG5nIQ==", "Vi1GT1JNQVQ=", 64
    End If
    Exit Sub
ErrH:
    VFormat_Core.ShowUniMsg "Loi khi khoi phuc mac dinh: " & Err.Description, "Vi1GT1JNQVQ=", 48
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub

