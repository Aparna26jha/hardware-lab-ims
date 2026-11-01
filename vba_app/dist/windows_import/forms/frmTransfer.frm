VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTransfer
   Caption         =   "Hardware Lab IMS - Transfer"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblTitle
      Caption         =   "Assign and transfer"
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7200
   End
   Begin MSForms.Label lblBarcode
      Caption         =   "Asset barcode"
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7200
   End
   Begin MSForms.TextBox txtBarcode
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   7200
   End
   Begin MSForms.Label lblOwnership
      Caption         =   "Ownership:"
      Height          =   280
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   7200
   End
   Begin MSForms.Label lblType
      Caption         =   "Type / model:"
      Height          =   280
      Left            =   240
      TabIndex        =   4
      Top             =   1620
      Width           =   7200
   End
   Begin MSForms.Label lblCurrent
      Caption         =   "Current assignee:"
      Height          =   280
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   7200
   End
   Begin MSForms.Label lblState
      Caption         =   "Status:"
      Height          =   280
      Left            =   240
      TabIndex        =   6
      Top             =   2220
      Width           =   7200
   End
   Begin MSForms.Label lblStatus
      Caption         =   "New status"
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   2700
      Width           =   7200
   End
   Begin MSForms.ComboBox cboStatus
      Height          =   360
      Left            =   240
      Style           =   2
      TabIndex        =   8
      Top             =   2940
      Width           =   7200
   End
   Begin MSForms.Label lblNewEmp
      Caption         =   "Target employee ID"
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   7200
   End
   Begin MSForms.TextBox txtNewEmpID
      Height          =   360
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   7200
   End
   Begin MSForms.Label lblRemarks
      Caption         =   "Remarks"
      Height          =   240
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   7200
   End
   Begin MSForms.TextBox txtRemarks
      Height          =   720
      Left            =   240
      MultiLine       =   -1
      TabIndex        =   12
      Top             =   4440
      Width           =   7200
   End
   Begin MSForms.CommandButton btnTransfer
      Caption         =   "Record transfer"
      Height          =   420
      Left            =   240
      TabIndex        =   13
      Top             =   5340
      Width           =   2100
   End
   Begin MSForms.CommandButton btnClose
      Caption         =   "Close"
      Cancel          =   -1
      Height          =   420
      Left            =   2520
      TabIndex        =   14
      Top             =   5340
      Width           =   1500
   End
   Begin MSForms.Label lblMessage
      Caption         =   ""
      Height          =   900
      Left            =   240
      TabIndex        =   15
      Top             =   5880
      Width           =   7200
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoaded As Boolean

Private Sub UserForm_Initialize()
    If Not IMS_IsCustodianRole(gRole) Then
        MsgBox "Only inventory custodians can transfer hardware.", vbExclamation, IMS_APP_NAME
        Unload Me
        Exit Sub
    End If
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 16
    IMS_FillCombo cboStatus, IMS_StatusList(), False
    cboStatus.Value = "Assigned to Engineer"
    txtBarcode.MaxLength = 64
    txtNewEmpID.MaxLength = 32
    ClearView
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    txtBarcode.SetFocus
End Sub

Private Sub txtBarcode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        LoadAsset
    End If
End Sub

Private Sub btnTransfer_Click()
    If Not mLoaded Then
        IMS_Flash lblMessage, "Scan a barcode first.", True
        Exit Sub
    End If
    If IMS_TransferAsset(txtBarcode.Text, txtNewEmpID.Text, txtRemarks.Text, cboStatus.Value) Then
        IMS_Flash lblMessage, "Transfer recorded for " & IMS_NormalizeBarcode(txtBarcode.Text) & ".", False
        LoadAsset
        txtNewEmpID.Text = ""
        txtRemarks.Text = ""
    Else
        IMS_Flash lblMessage, gLastError, True
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub LoadAsset()
    Dim asset As Object
    Dim code As String
    code = IMS_NormalizeBarcode(txtBarcode.Text)
    txtBarcode.Text = code
    Set asset = IMS_GetAsset(code)
    If asset Is Nothing Then
        mLoaded = False
        lblOwnership.Caption = "Ownership:"
        lblType.Caption = "Type / model:"
        lblCurrent.Caption = "Current assignee:"
        lblState.Caption = "Status:"
        IMS_Flash lblMessage, "Barcode " & code & " is not in the lab database.", True
        Exit Sub
    End If
    mLoaded = True
    lblOwnership.Caption = "Ownership: " & DictStr(asset, "Ownership")
    lblType.Caption = "Type / model: " & DictStr(asset, "AssetType") & "  " & DictStr(asset, "Model")
    If DictStr(asset, "CurrentAssigneeID") = "" Then
        lblCurrent.Caption = "Current assignee: (in lab cage)"
    Else
        lblCurrent.Caption = "Current assignee: " & DictStr(asset, "CurrentAssigneeID") & "  " & DictStr(asset, "CurrentAssigneeName")
    End If
    lblState.Caption = "Status: " & DictStr(asset, "Status")
    IMS_Flash lblMessage, "Asset loaded. Choose the new status and record the transfer.", False
End Sub

Private Sub ClearView()
    mLoaded = False
    lblOwnership.Caption = "Ownership:"
    lblType.Caption = "Type / model:"
    lblCurrent.Caption = "Current assignee:"
    lblState.Caption = "Status:"
    IMS_Flash lblMessage, "Scan a barcode. Enter loads the current assignee.", False
End Sub

Private Function DictStr(ByVal d As Object, ByVal key As String) As String
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsNull(d(key)) Then Exit Function
    DictStr = Trim$(CStr(d(key)))
End Function
