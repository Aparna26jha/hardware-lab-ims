VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUserAdmin
   Caption         =   "Hardware Lab IMS - Accounts"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9800
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblTitle
      Caption         =   "User directory"
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8400
   End
   Begin MSForms.Label lblHeaders
      Caption         =   "Alias | Name | Role | Status | Requested"
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   560
      Width           =   8400
   End
   Begin MSForms.ListBox lstUsers
      ColumnCount     =   6
      Height          =   2800
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   9300
   End
   Begin MSForms.Label lblSelected
      Caption         =   "Select an account"
      Height          =   280
      Left            =   240
      TabIndex        =   3
      Top             =   3760
      Width           =   9300
   End
   Begin MSForms.ComboBox cboRole
      Height          =   360
      Left            =   240
      Style           =   2
      TabIndex        =   4
      Top             =   4180
      Width           =   2800
   End
   Begin MSForms.ComboBox cboLead
      Height          =   360
      Left            =   3160
      Style           =   2
      TabIndex        =   5
      Top             =   4180
      Width           =   2400
   End
   Begin MSForms.CheckBox chkActive
      Caption         =   "Approve / Active"
      Height          =   300
      Left            =   5680
      TabIndex        =   6
      Top             =   4200
      Width           =   1800
   End
   Begin MSForms.CommandButton btnApply
      Caption         =   "Save"
      Height          =   400
      Left            =   7600
      TabIndex        =   7
      Top             =   4160
      Width           =   900
   End
   Begin MSForms.CommandButton btnReject
      Caption         =   "Reject"
      Height          =   400
      Left            =   8600
      TabIndex        =   8
      Top             =   4160
      Width           =   900
   End
   Begin MSForms.Label lblAllow
      Caption         =   "Add a person to the lab directory before they can request access"
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   4800
      Width           =   9300
   End
   Begin MSForms.TextBox txtAlias
      Height          =   360
      Left            =   240
      TabIndex        =   10
      Top             =   5120
      Width           =   1800
   End
   Begin MSForms.TextBox txtEmail
      Height          =   360
      Left            =   2160
      TabIndex        =   11
      Top             =   5120
      Width           =   2800
   End
   Begin MSForms.TextBox txtName
      Height          =   360
      Left            =   5080
      TabIndex        =   12
      Top             =   5120
      Width           =   2800
   End
   Begin MSForms.CommandButton btnAllow
      Caption         =   "Add to directory"
      Height          =   400
      Left            =   8000
      TabIndex        =   13
      Top             =   5100
      Width           =   1500
   End
   Begin MSForms.CommandButton btnClose
      Caption         =   "Close"
      Cancel          =   -1
      Height          =   400
      Left            =   8000
      TabIndex        =   14
      Top             =   5680
      Width           =   1500
   End
   Begin MSForms.Label lblMessage
      Caption         =   "Role on save is the real role. The requested role is only a request. A User can be linked to one Lead Admin."
      Height          =   900
      Left            =   240
      TabIndex        =   15
      Top             =   6200
      Width           =   7600
   End
End
Attribute VB_Name = "frmUserAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    If Not IMS_IsSuperRole(gRole) Then
        MsgBox "Only a Super Admin can manage accounts.", vbExclamation, IMS_APP_NAME
        Unload Me
        Exit Sub
    End If
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 16
    IMS_FillCombo cboRole, IMS_RoleList(), False
    lstUsers.ColumnCount = 6
    lstUsers.ColumnWidths = "90 pt;180 pt;110 pt;70 pt;0 pt;110 pt"
    txtAlias.MaxLength = 32
    txtEmail.MaxLength = 120
    txtName.MaxLength = 120
    LoadUsers
End Sub

Private Sub lstUsers_Click()
    If lstUsers.ListIndex < 0 Then Exit Sub
    lblSelected.Caption = CStr(lstUsers.List(lstUsers.ListIndex, 0)) & "  " & CStr(lstUsers.List(lstUsers.ListIndex, 1)) & "  requested " & CStr(lstUsers.List(lstUsers.ListIndex, 5))
    cboRole.Value = CStr(lstUsers.List(lstUsers.ListIndex, 2))
    chkActive.Value = (CStr(lstUsers.List(lstUsers.ListIndex, 3)) = "Yes" Or CStr(lstUsers.List(lstUsers.ListIndex, 3)) = "Pending")
    cboLead.Value = CStr(lstUsers.List(lstUsers.ListIndex, 4))
End Sub

Private Sub btnApply_Click()
    Dim empId As String
    If lstUsers.ListIndex < 0 Then
        IMS_Flash lblMessage, "Select an account first.", True
        Exit Sub
    End If
    empId = CStr(lstUsers.List(lstUsers.ListIndex, 0))
    If IMS_SetUserRole(empId, cboRole.Value, CBool(chkActive.Value), cboLead.Value) Then
        IMS_Flash lblMessage, "Saved " & empId & ".", False
        LoadUsers
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
        IMS_Flash lblMessage, "Rejected " & empId & ".", False
        LoadUsers
    Else
        IMS_Flash lblMessage, gLastError, True
    End If
End Sub

Private Sub btnAllow_Click()
    If IMS_AddAllowedPerson(txtAlias.Text, txtEmail.Text, txtName.Text) Then
        IMS_Flash lblMessage, "Directory updated for " & UCase$(Trim$(txtAlias.Text)) & ".", False
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
    lstUsers.ColumnCount = 6
    cboLead.Clear
    cboLead.AddItem ""
    If Not IsArray(data) Then Exit Sub
    For r = 1 To UBound(data, 1)
        lstUsers.AddItem CStr(data(r, 1))
        For c = 2 To 6
            lstUsers.List(lstUsers.ListCount - 1, c - 1) = CStr(data(r, c))
        Next c
        If CStr(data(r, 3)) = "Lead Admin" And CStr(data(r, 4)) = "Yes" Then
            cboLead.AddItem CStr(data(r, 1))
        End If
    Next r
End Sub
