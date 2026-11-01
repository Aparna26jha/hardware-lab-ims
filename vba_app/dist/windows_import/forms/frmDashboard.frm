VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard
   Caption         =   "Hardware Lab IMS"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15600
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblWho
      Caption         =   "Signed in"
      Height          =   300
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   9000
   End
   Begin MSForms.CommandButton btnInward
      Caption         =   "Inward"
      Height          =   360
      Left            =   9360
      TabIndex        =   1
      Top             =   90
      Width           =   1400
   End
   Begin MSForms.CommandButton btnTransfer
      Caption         =   "Transfer"
      Height          =   360
      Left            =   10860
      TabIndex        =   2
      Top             =   90
      Width           =   1500
   End
   Begin MSForms.CommandButton btnUsers
      Caption         =   "Accounts"
      Height          =   360
      Left            =   12480
      TabIndex        =   3
      Top             =   90
      Width           =   1400
   End
   Begin MSForms.CommandButton btnSignOut
      Caption         =   "Sign out"
      Height          =   360
      Left            =   14000
      TabIndex        =   4
      Top             =   90
      Width           =   1400
   End
   Begin MSForms.Label lblOwner
      Caption         =   "Ownership"
      Height          =   220
      Left            =   180
      TabIndex        =   5
      Top             =   560
      Width           =   2200
   End
   Begin MSForms.ComboBox cboOwnership
      Height          =   320
      Left            =   180
      Style           =   2
      TabIndex        =   6
      Top             =   780
      Width           =   2400
   End
   Begin MSForms.Label lblStatusFilter
      Caption         =   "Status"
      Height          =   220
      Left            =   2700
      TabIndex        =   7
      Top             =   560
      Width           =   2200
   End
   Begin MSForms.ComboBox cboStatus
      Height          =   320
      Left            =   2700
      Style           =   2
      TabIndex        =   8
      Top             =   780
      Width           =   2800
   End
   Begin MSForms.Label lblTypeFilter
      Caption         =   "Asset type"
      Height          =   220
      Left            =   5640
      TabIndex        =   9
      Top             =   560
      Width           =   2200
   End
   Begin MSForms.ComboBox cboType
      Height          =   320
      Left            =   5640
      Style           =   2
      TabIndex        =   10
      Top             =   780
      Width           =   2600
   End
   Begin MSForms.Label lblSearch
      Caption         =   "Search"
      Height          =   220
      Left            =   8400
      TabIndex        =   11
      Top             =   560
      Width           =   3000
   End
   Begin MSForms.TextBox txtSearch
      Height          =   320
      Left            =   8400
      TabIndex        =   12
      Top             =   780
      Width           =   3600
   End
   Begin MSForms.CommandButton btnRefresh
      Caption         =   "Refresh"
      Height          =   320
      Left            =   12120
      TabIndex        =   13
      Top             =   780
      Width           =   1400
   End
   Begin MSForms.Label lblHeaders
      Caption         =   "Barcode | Ownership | Type | Model | PO | Status | Assignee ID | Assignee name"
      Height          =   240
      Left            =   180
      TabIndex        =   14
      Top             =   1200
      Width           =   15000
   End
   Begin MSForms.ListBox lstAssets
      ColumnCount     =   8
      Height          =   5600
      Left            =   180
      TabIndex        =   15
      Top             =   1460
      Width           =   15200
   End
   Begin MSForms.Label lblCount
      Caption         =   "0 assets"
      Height          =   240
      Left            =   180
      TabIndex        =   16
      Top             =   7160
      Width           =   4000
   End
   Begin MSForms.CommandButton btnExport
      Caption         =   "Export inventory CSV"
      Height          =   400
      Left            =   180
      TabIndex        =   17
      Top             =   7560
      Width           =   2600
   End
   Begin MSForms.CommandButton btnExportAudit
      Caption         =   "Export audit CSV"
      Height          =   400
      Left            =   2920
      TabIndex        =   18
      Top             =   7560
      Width           =   2300
   End
   Begin MSForms.CommandButton btnReset
      Caption         =   "Reset sample data"
      Height          =   400
      Left            =   5360
      TabIndex        =   19
      Top             =   7560
      Width           =   2300
   End
   Begin MSForms.Label lblMessage
      Caption         =   ""
      Height          =   700
      Left            =   180
      TabIndex        =   20
      Top             =   8100
      Width           =   15200
   End
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    lblWho.Font.Bold = True
    IMS_FillCombo cboOwnership, IMS_OwnershipList(), True
    IMS_FillCombo cboStatus, IMS_StatusList(), True
    IMS_FillCombo cboType, IMS_AssetTypeList(), True
    lstAssets.ColumnCount = 8
    lstAssets.ColumnWidths = "110 pt;115 pt;100 pt;90 pt;100 pt;130 pt;80 pt;140 pt"
    btnInward.Visible = IMS_IsCustodianRole(gRole)
    btnTransfer.Visible = IMS_IsCustodianRole(gRole)
    btnUsers.Visible = IMS_IsSuperRole(gRole)
    btnReset.Visible = IMS_IsSuperRole(gRole)
    btnExportAudit.Visible = IMS_IsCustodianRole(gRole)
    If gRole = "User" Then
        lblWho.Caption = gEmpName & "  (" & gEmpID & ")  Engineer view - your assigned assets only"
    Else
        lblWho.Caption = gEmpName & "  (" & gEmpID & ")  " & gRole
    End If
End Sub

Private Sub UserForm_Activate()
    LoadAssets
End Sub

Private Sub btnRefresh_Click()
    LoadAssets
End Sub

Private Sub cboOwnership_Change()
    LoadAssets
End Sub

Private Sub cboStatus_Change()
    LoadAssets
End Sub

Private Sub cboType_Change()
    LoadAssets
End Sub

Private Sub txtSearch_AfterUpdate()
    LoadAssets
End Sub

Private Sub btnInward_Click()
    frmInward.Show
End Sub

Private Sub btnTransfer_Click()
    frmTransfer.Show
End Sub

Private Sub btnUsers_Click()
    frmUserAdmin.Show
End Sub

Private Sub btnSignOut_Click()
    IMS_SignOut
    Me.Hide
End Sub

Private Sub btnReset_Click()
    IMS_ResetBackend
    LoadAssets
End Sub

Private Sub btnExport_Click()
    Dim data As Variant
    Dim path As Variant
    data = CurrentRows()
    path = IMS_SuggestExportPath("LabInventory")
    If VarType(path) = vbBoolean Then Exit Sub
    On Error GoTo Fail
    IMS_WriteCsv CStr(path), IMS_AssetHeaders(), data
    IMS_Flash lblMessage, "Inventory snapshot written to " & CStr(path), False
    Exit Sub
Fail:
    IMS_Flash lblMessage, Err.Description, True
End Sub

Private Sub btnExportAudit_Click()
    Dim assets As Variant
    Dim auditRows As Variant
    Dim filtered As Variant
    Dim path As Variant
    Dim wanted As Object
    Dim r As Long
    Dim kept As Long
    Dim i As Long
    assets = CurrentRows()
    auditRows = IMS_QueryAudit()
    If gLastError <> "" And IsEmpty(auditRows) Then
        IMS_Flash lblMessage, gLastError, True
        Exit Sub
    End If
    Set wanted = CreateObject("Scripting.Dictionary")
    wanted.CompareMode = 1
    If IsArray(assets) Then
        For r = 1 To UBound(assets, 1)
            wanted(CStr(assets(r, 1))) = True
        Next r
    End If
    If Not IsArray(auditRows) Then
        filtered = Empty
    Else
        kept = 0
        For r = 1 To UBound(auditRows, 1)
            If wanted.Exists(CStr(auditRows(r, 3))) Then kept = kept + 1
        Next r
        If kept = 0 Then
            filtered = Empty
        Else
            ReDim filtered(1 To kept, 1 To 7)
            i = 0
            For r = 1 To UBound(auditRows, 1)
                If wanted.Exists(CStr(auditRows(r, 3))) Then
                    i = i + 1
                    Dim c As Long
                    For c = 1 To 7
                        filtered(i, c) = auditRows(r, c)
                    Next c
                End If
            Next r
        End If
    End If
    path = IMS_SuggestExportPath("LabAudit")
    If VarType(path) = vbBoolean Then Exit Sub
    On Error GoTo Fail
    IMS_WriteCsv CStr(path), IMS_AuditHeaders(), filtered
    IMS_Flash lblMessage, "Audit snapshot for the assets on screen written to " & CStr(path), False
    Exit Sub
Fail:
    IMS_Flash lblMessage, Err.Description, True
End Sub

Private Sub LoadAssets()
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    On Error GoTo Fail
    data = CurrentRows()
    lstAssets.Clear
    lstAssets.ColumnCount = 8
    If Not IsArray(data) Then
        lblCount.Caption = "0 assets"
        Exit Sub
    End If
    For r = 1 To UBound(data, 1)
        lstAssets.AddItem CStr(data(r, 1))
        For c = 2 To 8
            lstAssets.List(lstAssets.ListCount - 1, c - 1) = CStr(data(r, c))
        Next c
    Next r
    lblCount.Caption = CStr(UBound(data, 1)) & " assets"
    Exit Sub
Fail:
    lblCount.Caption = "0 assets"
    IMS_Flash lblMessage, Err.Description, True
End Sub

Private Function CurrentRows() As Variant
    CurrentRows = IMS_QueryAssets(cboOwnership.Value, cboStatus.Value, cboType.Value, txtSearch.Text, "")
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        IMS_SignOut
        Me.Hide
        Cancel = True
    End If
End Sub
