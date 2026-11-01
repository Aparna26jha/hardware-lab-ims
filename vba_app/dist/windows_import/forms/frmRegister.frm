VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegister
   Caption         =   "Hardware Lab IMS - Create account"
   ClientHeight    =   6240
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
      Caption         =   "New accounts are engineers (User). A Super Admin can later mark a custodian or another Super Admin."
      Height          =   540
      Left            =   300
      TabIndex        =   1
      Top             =   560
      Width           =   6240
   End
   Begin MSForms.Label lblEmpID
      Caption         =   "Employee ID"
      Height          =   240
      Left            =   300
      TabIndex        =   2
      Top             =   1200
      Width           =   6240
   End
   Begin MSForms.TextBox txtEmpID
      Height          =   360
      Left            =   300
      TabIndex        =   3
      Top             =   1440
      Width           =   6240
   End
   Begin MSForms.Label lblEmpName
      Caption         =   "Full name"
      Height          =   240
      Left            =   300
      TabIndex        =   4
      Top             =   1920
      Width           =   6240
   End
   Begin MSForms.TextBox txtEmpName
      Height          =   360
      Left            =   300
      TabIndex        =   5
      Top             =   2160
      Width           =   6240
   End
   Begin MSForms.Label lblCode
      Caption         =   "Lab access code"
      Height          =   240
      Left            =   300
      TabIndex        =   6
      Top             =   2640
      Width           =   6240
   End
   Begin MSForms.TextBox txtAccessCode
      Height          =   360
      Left            =   300
      TabIndex        =   7
      Top             =   2880
      Width           =   6240
   End
   Begin MSForms.Label lblConfirm
      Caption         =   "Confirm lab access code"
      Height          =   240
      Left            =   300
      TabIndex        =   8
      Top             =   3360
      Width           =   6240
   End
   Begin MSForms.TextBox txtConfirm
      Height          =   360
      Left            =   300
      TabIndex        =   9
      Top             =   3600
      Width           =   6240
   End
   Begin MSForms.CommandButton btnCreate
      Caption         =   "Create account"
      Default         =   -1
      Height          =   420
      Left            =   300
      TabIndex        =   10
      Top             =   4200
      Width           =   2100
   End
   Begin MSForms.CommandButton btnBack
      Caption         =   "Back to sign in"
      Height          =   420
      Left            =   2580
      TabIndex        =   11
      Top             =   4200
      Width           =   2100
   End
   Begin MSForms.Label lblMessage
      Caption         =   ""
      Height          =   900
      Left            =   300
      TabIndex        =   12
      Top             =   4800
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
End Sub

Private Sub btnCreate_Click()
    If IMS_RegisterUser(txtEmpID.Text, txtEmpName.Text, txtAccessCode.Text, txtConfirm.Text) Then
        MsgBox "Account created. Sign in with employee ID " & UCase$(Trim$(txtEmpID.Text)) & ".", vbInformation, IMS_APP_NAME
        Unload Me
    Else
        IMS_Flash lblMessage, gLastError, True
    End If
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub
