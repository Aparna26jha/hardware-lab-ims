VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUserAdmin
   Caption         =   "Hardware Lab IMS - Accounts"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9800
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUserAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Controls are created in UserForm_Initialize (see modUi).

Private mBlocked As Boolean

Private lblMessage As MSForms.Label
Private lblSelected As MSForms.Label
Private WithEvents lstUsers As MSForms.ListBox
Private cboRole As MSForms.ComboBox
Private cboLead As MSForms.ComboBox
Private chkActive As MSForms.CheckBox
Private txtAlias As MSForms.TextBox
Private txtEmail As MSForms.TextBox
Private txtName As MSForms.TextBox
Private WithEvents btnApply As MSForms.CommandButton
Private WithEvents btnReject As MSForms.CommandButton
Private WithEvents btnAllow As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim widths As Variant
    Dim heading As Object
    If Not IMS_IsSuperRole(gRole) Then
        mBlocked = True
        Exit Sub
    End If
    UiSizeForm Me, "Hardware Lab IMS - Accounts", 640, 480
    UiTitle Me, "lblTitle", "Accounts and access requests", 18, 12, 600
    UiHint Me, "lblHint", "Pick a person, choose the real role (and a lead for a User), tick Approve, then Save. Reject closes a request.", 18, 38, 600, 26

    widths = Array(80, 170, 90, 70, 70, 100)
    UiColumnHeads Me, "lblCol", Array("Alias", "Name", "Role", "Status", "Lead", "Requested"), widths, 18, 70
    Set lstUsers = UiList(Me, "lstUsers", 18, 84, 600, 170, UiWidthsText(widths))
    Set lblSelected = UiLabel(Me, "lblSelected", "Select an account in the list.", 18, 260, 600, 14)
    lblSelected.Font.Bold = True

    UiLabel Me, "lblRole", "Real role", 18, 278, 150, 12
    Set cboRole = UiCombo(Me, "cboRole", 18, 292, 150)
    IMS_FillCombo cboRole, IMS_RoleList(), False
    UiLabel Me, "lblLead", "Lead Admin (Users only)", 176, 278, 150, 12
    Set cboLead = UiCombo(Me, "cboLead", 176, 292, 150)
    Set chkActive = UiCheck(Me, "chkActive", "Approve / Active", 334, 292, 120)
    Set btnApply = UiButton(Me, "btnApply", "Save", 460, 290, 76)
    Set btnReject = UiButton(Me, "btnReject", "Reject", 542, 290, 76)

    Set heading = UiLabel(Me, "lblDirHead", "Lab directory", 18, 326, 600, 14)
    heading.Font.Bold = True
    UiHint Me, "lblDirHint", "A person must be on this list before they can request access. Their alias, mail ID and full name must match.", 18, 340, 600, 12
    UiLabel Me, "lblAlias", "Alias ID", 18, 356, 110, 12
    Set txtAlias = UiText(Me, "txtAlias", 18, 370, 110, 32)
    UiLabel Me, "lblEmail", "Mail ID", 134, 356, 190, 12
    Set txtEmail = UiText(Me, "txtEmail", 134, 370, 190, 120)
    UiLabel Me, "lblName", "Full name", 330, 356, 170, 12
    Set txtName = UiText(Me, "txtName", 330, 370, 170, 120)
    Set btnAllow = UiButton(Me, "btnAllow", "Add to directory", 506, 368, 112)

    Set lblMessage = UiLabel(Me, "lblMessage", "", 18, 402, 520, 36)
    Set btnClose = UiButton(Me, "btnClose", "Close", 542, 404, 76)
    btnClose.Cancel = True
    LoadUsers
End Sub

Private Sub UserForm_Activate()
    If mBlocked Then
        MsgBox "Only a Super Admin can manage accounts.", vbExclamation, IMS_APP_NAME
        Unload Me
    End If
End Sub

Private Sub lstUsers_Click()
    Dim statusText As String
    If lstUsers.ListIndex < 0 Then Exit Sub
    statusText = CStr(lstUsers.List(lstUsers.ListIndex, 3))
    lblSelected.Caption = CStr(lstUsers.List(lstUsers.ListIndex, 0)) & "  " & CStr(lstUsers.List(lstUsers.ListIndex, 1)) & _
        "  -  " & statusText & ", asked for " & CStr(lstUsers.List(lstUsers.ListIndex, 5))
    UiPick cboRole, CStr(lstUsers.List(lstUsers.ListIndex, 2))
    chkActive.Value = (statusText = "Active" Or statusText = "Pending")
    UiPick cboLead, CStr(lstUsers.List(lstUsers.ListIndex, 4))
    If cboLead.ListIndex < 0 Then cboLead.ListIndex = 0
End Sub

Private Sub btnApply_Click()
    Dim empId As String
    If lstUsers.ListIndex < 0 Then
        IMS_Flash lblMessage, "Select an account first.", True
        Exit Sub
    End If
    empId = CStr(lstUsers.List(lstUsers.ListIndex, 0))
    If IMS_SetUserRole(empId, cboRole.Text, CBool(chkActive.Value), cboLead.Text) Then
        LoadUsers
        IMS_Flash lblMessage, "Saved " & empId & ".", False
    Else
        IMS_Flash lblMessage, gLastError, True
    End If
End Sub

Private Sub btnReject_Click()
    Dim empId As String
    If lstUsers.ListIndex < 0 Then
        IMS_Flash lblMessage, "Select an account first.", True
        Exit Sub
    End If
    empId = CStr(lstUsers.List(lstUsers.ListIndex, 0))
    If IMS_RejectUser(empId) Then
        LoadUsers
        IMS_Flash lblMessage, "Rejected " & empId & ".", False
    Else
        IMS_Flash lblMessage, gLastError, True
    End If
End Sub

Private Sub btnAllow_Click()
    If IMS_AddAllowedPerson(txtAlias.Text, txtEmail.Text, txtName.Text) Then
        IMS_Flash lblMessage, "Directory updated for " & UCase$(Trim$(txtAlias.Text)) & ". That person can now request access.", False
        txtAlias.Text = ""
        txtEmail.Text = ""
        txtName.Text = ""
    Else
        IMS_Flash lblMessage, gLastError, True
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub LoadUsers()
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    data = IMS_ListUsers()
    lstUsers.Clear
    cboLead.Clear
    cboLead.AddItem ""
    lblSelected.Caption = "Select an account in the list."
    If Not IsArray(data) Then
        If gLastError <> "" Then IMS_Flash lblMessage, gLastError, True
        Exit Sub
    End If
    For r = 1 To UBound(data, 1)
        lstUsers.AddItem CStr(data(r, 1))
        For c = 2 To 6
            lstUsers.List(lstUsers.ListCount - 1, c - 1) = CStr(data(r, c))
        Next c
        If CStr(data(r, 3)) = "Lead Admin" And CStr(data(r, 4)) = "Active" Then
            cboLead.AddItem CStr(data(r, 1))
        End If
    Next r
End Sub
