VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAccount
   Caption         =   "Hardware Lab IMS - My account"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Controls are created in UserForm_Initialize (see modUi).

Private lblMessage As MSForms.Label
Private txtCurrent As MSForms.TextBox
Private txtNew As MSForms.TextBox
Private txtConfirm As MSForms.TextBox
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim info As Object
    Dim heading As Object
    UiSizeForm Me, "Hardware Lab IMS - My account", 380, 384
    UiTitle Me, "lblTitle", "My account", 18, 12, 340
    UiHint Me, "lblHint", "Check your details. To change your access code, type the current code and the new code twice, then Save.", 18, 38, 340, 26

    Set info = IMS_GetMyAccount()
    If info Is Nothing Then
        DetailRow "Alias ID", gEmpID, 70
        DetailRow "Name", gEmpName, 88
        DetailRow "Role", gRole, 106
        DetailRow "Mail ID", "", 124
        DetailRow "Lead Admin", "", 142
    Else
        DetailRow "Alias ID", CStr(info("EmpID")), 70
        DetailRow "Name", CStr(info("EmpName")), 88
        DetailRow "Role", CStr(info("Role")), 106
        DetailRow "Mail ID", CStr(info("Email")), 124
        DetailRow "Lead Admin", CStr(info("Lead")), 142
    End If
    UiHint Me, "lblDetailsHint", "A Super Admin changes your role, lead, or mail ID.", 18, 160, 340, 12

    Set heading = UiLabel(Me, "lblCodeHead", "Change my access code", 18, 184, 340, 14)
    heading.Font.Bold = True
    UiLabel Me, "lblCurrent", "Current access code", 18, 202, 340, 12
    Set txtCurrent = UiText(Me, "txtCurrent", 18, 216, 340, 32, True)
    UiLabel Me, "lblNew", "New access code (4 to 32 characters)", 18, 240, 340, 12
    Set txtNew = UiText(Me, "txtNew", 18, 254, 340, 32, True)
    UiLabel Me, "lblConfirm", "Type the new access code again", 18, 278, 340, 12
    Set txtConfirm = UiText(Me, "txtConfirm", 18, 292, 340, 32, True)
    Set btnSave = UiButton(Me, "btnSave", "Save new code", 18, 320, 110)
    btnSave.Default = True
    Set btnClose = UiButton(Me, "btnClose", "Close", 136, 320, 80)
    btnClose.Cancel = True
    Set lblMessage = UiLabel(Me, "lblMessage", "", 18, 346, 340, 26)
    If info Is Nothing Then IMS_Flash lblMessage, gLastError, True
End Sub

Private Sub btnSave_Click()
    If IMS_ChangeAccessCode(txtCurrent.Text, txtNew.Text, txtConfirm.Text) Then
        txtCurrent.Text = ""
        txtNew.Text = ""
        txtConfirm.Text = ""
        IMS_Flash lblMessage, "Your access code was changed. Use the new code next time you sign in.", False
    Else
        IMS_Flash lblMessage, gLastError, True
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub DetailRow(ByVal fieldName As String, ByVal value As String, ByVal topPt As Single)
    Dim c As Object
    Set c = UiLabel(Me, "lblK" & CStr(topPt), fieldName, 18, topPt, 90, 14)
    c.ForeColor = UI_MUTED
    If value = "" Then value = "-"
    Set c = UiLabel(Me, "lblV" & CStr(topPt), value, 110, topPt, 248, 14)
    c.Font.Bold = True
End Sub
