VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegister
   Caption         =   "Hardware Lab IMS - Create account"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblTitle
      Caption         =   "Create a lab account"
      Height          =   360
      Left            =   300
      TabIndex        =   0
      Top             =   180
      Width           =   6240
   End
   Begin MSForms.Label lblRoleHint
      Caption         =   "The alias and mail ID must already be on the lab directory. Sign-in stays closed until a Super Admin approves the role."
      Height          =   720
      Left            =   300
      TabIndex        =   1
      Top             =   560
      Width           =   6240
   End
   Begin MSForms.Label lblEmpID
      Caption         =   "Alias ID"
      Height          =   240
      Left            =   300
      TabIndex        =   2
      Top             =   1360
      Width           =   6240
   End
   Begin MSForms.TextBox txtEmpID
      Height          =   360
      Left            =   300
      TabIndex        =   3
      Top             =   1600
      Width           =   6240
   End
   Begin MSForms.Label lblEmpName
      Caption         =   "Full name"
      Height          =   240
      Left            =   300
      TabIndex        =   4
      Top             =   2100
      Width           =   6240
   End
   Begin MSForms.TextBox txtEmpName
      Height          =   360
      Left            =   300
      TabIndex        =   5
      Top             =   2340
      Width           =   6240
   End
   Begin MSForms.Label lblEmail
      Caption         =   "Mail ID"
      Height          =   240
      Left            =   300
      TabIndex        =   6
      Top             =   2820
      Width           =   6240
   End
   Begin MSForms.TextBox txtEmail
      Height          =   360
      Left            =   300
      TabIndex        =   7
      Top             =   3060
      Width           =   6240
   End
   Begin MSForms.Label lblRequest
      Caption         =   "Requested role (a Super Admin decides the real role)"
      Height          =   240
      Left            =   300
      TabIndex        =   8
      Top             =   3540
      Width           =   6240
   End
   Begin MSForms.ComboBox cboRequest
      Height          =   360
      Left            =   300
      Style           =   2
      TabIndex        =   9
      Top             =   3780
      Width           =   6240
   End
   Begin MSForms.Label lblCode
      Caption         =   "Lab access code"
      Height          =   240
      Left            =   300
      TabIndex        =   10
      Top             =   4260
      Width           =   6240
   End
   Begin MSForms.TextBox txtAccessCode
      Height          =   360
      Left            =   300
      TabIndex        =   11
      Top             =   4500
      Width           =   6240
   End
   Begin MSForms.Label lblConfirm
      Caption         =   "Confirm lab access code"
      Height          =   240
      Left            =   300
      TabIndex        =   12
      Top             =   4980
      Width           =   6240
   End
   Begin MSForms.TextBox txtConfirm
      Height          =   360
      Left            =   300
      TabIndex        =   13
      Top             =   5220
      Width           =   6240
   End
   Begin MSForms.CommandButton btnCreate
      Caption         =   "Request access"
      Default         =   -1
      Height          =   420
      Left            =   300
      TabIndex        =   14
      Top             =   5820
      Width           =   2100
   End
   Begin MSForms.CommandButton btnBack
      Caption         =   "Back to sign in"
      Height          =   420
      Left            =   2580
      TabIndex        =   15
      Top             =   5820
      Width           =   2100
   End
   Begin MSForms.Label lblMessage
      Caption         =   ""
      Height          =   1200
      Left            =   300
      TabIndex        =   16
      Top             =   6420
      Width           =   6240
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 16
    txtEmpID.MaxLength = 32
    txtEmpName.MaxLength = 120
    txtAccessCode.MaxLength = 32
    txtConfirm.MaxLength = 32
    txtAccessCode.PasswordChar = "*"
    txtConfirm.PasswordChar = "*"
    txtEmail.MaxLength = 120
    IMS_FillCombo cboRequest, IMS_RoleList(), False
    cboRequest.Value = "User"
End Sub

Private Sub btnCreate_Click()
    If IMS_RegisterUser(txtEmpID.Text, txtEmpName.Text, txtEmail.Text, cboRequest.Value, txtAccessCode.Text, txtConfirm.Text) Then
        MsgBox "Request sent for " & UCase$(Trim$(txtEmpID.Text)) & "." & vbCrLf & "A Super Admin must approve it before sign-in.", vbInformation, IMS_APP_NAME
        Unload Me
    Else
        IMS_Flash lblMessage, gLastError, True
    End If
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub
