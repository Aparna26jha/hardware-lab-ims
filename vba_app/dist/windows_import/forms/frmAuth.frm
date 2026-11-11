VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAuth
   Caption         =   "Hardware Lab IMS - Sign in"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Controls are created in UserForm_Initialize (see modUi).

Private lblMessage As MSForms.Label
Private txtEmpID As MSForms.TextBox
Private txtAccessCode As MSForms.TextBox
Private WithEvents btnLogin As MSForms.CommandButton
Private WithEvents btnRegister As MSForms.CommandButton
Private WithEvents btnFillSuper As MSForms.CommandButton
Private WithEvents btnFillAdmin As MSForms.CommandButton
Private WithEvents btnFillLead As MSForms.CommandButton
Private WithEvents btnFillUser As MSForms.CommandButton

Private Sub UserForm_Initialize()
    UiSizeForm Me, "Hardware Lab IMS - Sign in", 360, 300
    UiTitle Me, "lblTitle", "Hardware Lab IMS", 18, 12, 320
    UiHint Me, "lblHint", "Type your alias ID and lab access code, then press Enter. New here? Choose Create account.", 18, 38, 320, 26
    UiLabel Me, "lblEmpID", "Alias ID", 18, 72, 320, 12
    Set txtEmpID = UiText(Me, "txtEmpID", 18, 86, 320, 32)
    UiLabel Me, "lblCode", "Lab access code", 18, 112, 320, 12
    Set txtAccessCode = UiText(Me, "txtAccessCode", 18, 126, 320, 32, True)
    Set btnLogin = UiButton(Me, "btnLogin", "Sign in", 18, 156, 100)
    btnLogin.Default = True
    Set btnRegister = UiButton(Me, "btnRegister", "Create account", 126, 156, 110)

    UiHint Me, "lblDemo", "Sample accounts for the walkthrough:", 18, 190, 320, 12
    Set btnFillSuper = UiButton(Me, "btnFillSuper", "Super Admin", 18, 204, 76)
    Set btnFillAdmin = UiButton(Me, "btnFillAdmin", "Custodian", 99, 204, 76)
    Set btnFillLead = UiButton(Me, "btnFillLead", "Lead", 180, 204, 76)
    Set btnFillUser = UiButton(Me, "btnFillUser", "Engineer", 261, 204, 77)
    Me.Controls("lblDemo").Visible = IMS_SHOW_DEMO_FILL
    btnFillSuper.Visible = IMS_SHOW_DEMO_FILL
    btnFillAdmin.Visible = IMS_SHOW_DEMO_FILL
    btnFillLead.Visible = IMS_SHOW_DEMO_FILL
    btnFillUser.Visible = IMS_SHOW_DEMO_FILL

    Set lblMessage = UiLabel(Me, "lblMessage", "", 18, 232, 320, 36)
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    txtEmpID.SetFocus
End Sub

Private Sub btnLogin_Click()
    If IMS_TryLogin(txtEmpID.Text, txtAccessCode.Text) Then
        Unload Me
    Else
        IMS_Flash lblMessage, gLastError, True
    End If
End Sub

Private Sub btnRegister_Click()
    frmRegister.Show
End Sub

Private Sub btnFillSuper_Click()
    FillSample "SA001", "1001"
End Sub

Private Sub btnFillAdmin_Click()
    FillSample "AD010", "2001"
End Sub

Private Sub btnFillLead_Click()
    FillSample "LD200", "4001"
End Sub

Private Sub btnFillUser_Click()
    FillSample "EN1042", "3001"
End Sub

Private Sub FillSample(ByVal empId As String, ByVal code As String)
    txtEmpID.Text = empId
    txtAccessCode.Text = code
    lblMessage.Caption = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then IMS_SignOut
End Sub
