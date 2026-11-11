VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegister
   Caption         =   "Hardware Lab IMS - Request access"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Controls are created in UserForm_Initialize (see modUi).

Private lblMessage As MSForms.Label
Private txtEmpID As MSForms.TextBox
Private txtEmpName As MSForms.TextBox
Private txtEmail As MSForms.TextBox
Private cboRequest As MSForms.ComboBox
Private txtAccessCode As MSForms.TextBox
Private txtConfirm As MSForms.TextBox
Private WithEvents btnCreate As MSForms.CommandButton
Private WithEvents btnBack As MSForms.CommandButton

Private Sub UserForm_Initialize()
    UiSizeForm Me, "Hardware Lab IMS - Request access", 360, 420
    UiTitle Me, "lblTitle", "Request access", 18, 12, 320
    UiHint Me, "lblHint", "Your alias, mail ID and full name must match the lab directory. You can sign in after a Super Admin approves the request.", 18, 38, 320, 38
    UiLabel Me, "lblEmpID", "Alias ID", 18, 80, 320, 12
    Set txtEmpID = UiText(Me, "txtEmpID", 18, 94, 320, 32)
    UiLabel Me, "lblEmpName", "Full name", 18, 118, 320, 12
    Set txtEmpName = UiText(Me, "txtEmpName", 18, 132, 320, 120)
    UiLabel Me, "lblEmail", "Mail ID", 18, 156, 320, 12
    Set txtEmail = UiText(Me, "txtEmail", 18, 170, 320, 120)
    UiLabel Me, "lblRequest", "Requested role (a Super Admin decides the real role)", 18, 194, 320, 12
    Set cboRequest = UiCombo(Me, "cboRequest", 18, 208, 320)
    IMS_FillCombo cboRequest, IMS_RoleList(), False
    UiPick cboRequest, "User"
    UiLabel Me, "lblCode", "Choose a lab access code (4 to 32 characters)", 18, 232, 320, 12
    Set txtAccessCode = UiText(Me, "txtAccessCode", 18, 246, 320, 32, True)
    UiLabel Me, "lblConfirm", "Type the access code again", 18, 270, 320, 12
    Set txtConfirm = UiText(Me, "txtConfirm", 18, 284, 320, 32, True)
    Set btnCreate = UiButton(Me, "btnCreate", "Request access", 18, 314, 110)
    btnCreate.Default = True
    Set btnBack = UiButton(Me, "btnBack", "Back to sign in", 136, 314, 110)
    btnBack.Cancel = True
    Set lblMessage = UiLabel(Me, "lblMessage", "", 18, 344, 320, 40)
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    txtEmpID.SetFocus
End Sub

Private Sub btnCreate_Click()
    If IMS_RegisterUser(txtEmpID.Text, txtEmpName.Text, txtEmail.Text, cboRequest.Text, txtAccessCode.Text, txtConfirm.Text) Then
        MsgBox "Request sent for " & UCase$(Trim$(txtEmpID.Text)) & "." & vbCrLf & "A Super Admin must approve it before sign-in.", vbInformation, IMS_APP_NAME
        Unload Me
    Else
        IMS_Flash lblMessage, gLastError, True
    End If
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub
