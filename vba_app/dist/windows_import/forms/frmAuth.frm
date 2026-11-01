VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAuth
   Caption         =   "Hardware Lab IMS - Sign in"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblTitle
      Caption         =   "Hardware Lab IMS"
      Height          =   360
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   6240
   End
   Begin MSForms.Label lblHint
      Caption         =   "Use your employee ID and lab access code. Do not reuse a Windows password."
      Height          =   480
      Left            =   300
      TabIndex        =   1
      Top             =   660
      Width           =   6240
   End
   Begin MSForms.Label lblEmpID
      Caption         =   "Employee ID"
      Height          =   240
      Left            =   300
      TabIndex        =   2
      Top             =   1260
      Width           =   6240
   End
   Begin MSForms.TextBox txtEmpID
      Height          =   360
      Left            =   300
      TabIndex        =   3
      Top             =   1500
      Width           =   6240
   End
   Begin MSForms.Label lblCode
      Caption         =   "Lab access code"
      Height          =   240
      Left            =   300
      TabIndex        =   4
      Top             =   2040
      Width           =   6240
   End
   Begin MSForms.TextBox txtAccessCode
      Height          =   360
      Left            =   300
      TabIndex        =   5
      Top             =   2280
      Width           =   6240
   End
   Begin MSForms.CommandButton btnLogin
      Caption         =   "Sign in"
      Default         =   -1
      Height          =   420
      Left            =   300
      TabIndex        =   6
      Top             =   2820
      Width           =   1800
   End
   Begin MSForms.CommandButton btnRegister
      Caption         =   "Create account"
      Height          =   420
      Left            =   2280
      TabIndex        =   7
      Top             =   2820
      Width           =   2100
   End
   Begin MSForms.CommandButton btnFillSuper
      Caption         =   "Sample Super Admin"
      Height          =   360
      Left            =   300
      TabIndex        =   8
      Top             =   3420
      Width           =   1980
   End
   Begin MSForms.CommandButton btnFillAdmin
      Caption         =   "Sample Custodian"
      Height          =   360
      Left            =   2400
      TabIndex        =   9
      Top             =   3420
      Width           =   1980
   End
   Begin MSForms.CommandButton btnFillUser
      Caption         =   "Sample Engineer"
      Height          =   360
      Left            =   4500
      TabIndex        =   10
      Top             =   3420
      Width           =   1980
   End
   Begin MSForms.Label lblMessage
      Caption         =   ""
      Height          =   900
      Left            =   300
      TabIndex        =   11
      Top             =   4020
      Width           =   6240
   End
End
Attribute VB_Name = "frmAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Hardware Lab IMS - Sign in"
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 16
    txtAccessCode.PasswordChar = "*"
    txtEmpID.MaxLength = 32
    txtAccessCode.MaxLength = 32
    btnFillSuper.Visible = IMS_SHOW_DEMO_FILL
    btnFillAdmin.Visible = IMS_SHOW_DEMO_FILL
    btnFillUser.Visible = IMS_SHOW_DEMO_FILL
    lblMessage.Caption = ""
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    txtEmpID.SetFocus
End Sub

Private Sub btnLogin_Click()
    If IMS_TryLogin(txtEmpID.Text, txtAccessCode.Text) Then
        Me.Hide
    Else
        IMS_Flash lblMessage, gLastError, True
    End If
End Sub

Private Sub btnRegister_Click()
    frmRegister.Show
End Sub

Private Sub btnFillSuper_Click()
    txtEmpID.Text = "SA001"
    txtAccessCode.Text = "1001"
    lblMessage.Caption = ""
End Sub

Private Sub btnFillAdmin_Click()
    txtEmpID.Text = "AD010"
    txtAccessCode.Text = "2001"
    lblMessage.Caption = ""
End Sub

Private Sub btnFillUser_Click()
    txtEmpID.Text = "EN1042"
    txtAccessCode.Text = "3001"
    lblMessage.Caption = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        IMS_SignOut
        Me.Hide
        Cancel = True
    End If
End Sub
