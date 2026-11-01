VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUserAdmin
   Caption         =   "Hardware Lab IMS - Accounts"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
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
      Caption         =   "Employee ID | Name | Role | Active"
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   560
      Width           =   8400
   End
   Begin MSForms.ListBox lstUsers
      ColumnCount     =   4
      Height          =   3600
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   8400
   End
   Begin MSForms.Label lblSelected
      Caption         =   "Select an account"
      Height          =   280
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   8400
   End
   Begin MSForms.ComboBox cboRole
      Height          =   360
      Left            =   240
      Style           =   2
      TabIndex        =   4
      Top             =   4980
      Width           =   3200
   End
   Begin MSForms.CheckBox chkActive
      Caption         =   "Active"
      Height          =   300
      Left            =   3600
      TabIndex        =   5
      Top             =   5000
      Width           =   1400
   End
   Begin MSForms.CommandButton btnApply
      Caption         =   "Save role"
      Height          =   400
      Left            =   5200
      TabIndex        =   6
      Top             =   4960
      Width           =   1600
   End
   Begin MSForms.CommandButton btnClose
      Caption         =   "Close"
      Cancel          =   -1
      Height          =   400
      Left            =   6960
      TabIndex        =   7
      Top             =   4960
      Width           =   1400
   End
   Begin MSForms.Label lblMessage
      Caption         =   "Self-registration always creates a User. Promote custodians here."
      Height          =   900
      Left            =   240
      TabIndex        =   8
      Top             =   5600
      Width           =   8400
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
    lstUsers.ColumnCount = 4
    lstUsers.ColumnWidths = "90 pt;220 pt;120 pt;60 pt"
    LoadUsers
End Sub

Private Sub lstUsers_Click()
    If lstUsers.ListIndex < 0 Then Exit Sub
    lblSelected.Caption = CStr(lstUsers.List(lstUsers.ListIndex, 0)) & "  " & CStr(lstUsers.List(lstUsers.ListIndex, 1))
    cboRole.Value = CStr(lstUsers.List(lstUsers.ListIndex, 2))
    chkActive.Value = (CStr(lstUsers.List(lstUsers.ListIndex, 3)) = "Yes")
End Sub

Private Sub btnApply_Click()
    Dim empId As String
    If lstUsers.ListIndex < 0 Then
        IMS_Flash lblMessage, "Select an account first.", True
        Exit Sub
    End If
    empId = CStr(lstUsers.List(lstUsers.ListIndex, 0))
    If IMS_SetUserRole(empId, cboRole.Value, CBool(chkActive.Value)) Then
        IMS_Flash lblMessage, "Saved " & empId & ".", False
        LoadUsers
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
    lstUsers.ColumnCount = 4
    If Not IsArray(data) Then Exit Sub
    For r = 1 To UBound(data, 1)
        lstUsers.AddItem CStr(data(r, 1))
        For c = 2 To 4
            lstUsers.List(lstUsers.ListCount - 1, c - 1) = CStr(data(r, c))
        Next c
    Next r
End Sub
