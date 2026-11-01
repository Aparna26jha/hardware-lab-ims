VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInward
   Caption         =   "Hardware Lab IMS - Inward"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14400
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblTitle
      Caption         =   "Inward material"
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6800
   End
   Begin MSForms.Label lblBarcode
      Caption         =   "Barcode / asset tag"
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   6800
   End
   Begin MSForms.TextBox txtBarcode
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   6800
   End
   Begin MSForms.OptionButton optAMD
      Caption         =   "AMD-Specific"
      GroupName       =   "Ownership"
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2400
   End
   Begin MSForms.OptionButton optWipro
      Caption         =   "Wipro-Specific"
      GroupName       =   "Ownership"
      Height          =   300
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   2800
   End
   Begin MSForms.Label lblAssetType
      Caption         =   "Asset type"
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   1740
      Width           =   6800
   End
   Begin MSForms.ComboBox cboAssetType
      Height          =   360
      Left            =   240
      Style           =   2
      TabIndex        =   6
      Top             =   1980
      Width           =   6800
   End
   Begin MSForms.Label lblPO
      Caption         =   "Purchase order"
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   2460
      Width           =   6800
   End
   Begin MSForms.TextBox txtPONumber
      Height          =   360
      Left            =   240
      TabIndex        =   8
      Top             =   2700
      Width           =   6800
   End
   Begin MSForms.Label lblManufacturer
      Caption         =   "Manufacturer"
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   3180
      Width           =   3200
   End
   Begin MSForms.TextBox txtManufacturer
      Height          =   360
      Left            =   240
      TabIndex        =   10
      Top             =   3420
      Width           =   3200
   End
   Begin MSForms.Label lblMfs
      Caption         =   "Manufacturer serial"
      Height          =   240
      Left            =   3600
      TabIndex        =   11
      Top             =   3180
      Width           =   3440
   End
   Begin MSForms.TextBox txtMfsSerialNo
      Height          =   360
      Left            =   3600
      TabIndex        =   12
      Top             =   3420
      Width           =   3440
   End
   Begin MSForms.Label lblMake
      Caption         =   "Make"
      Height          =   240
      Left            =   240
      TabIndex        =   13
      Top             =   3900
      Width           =   3200
   End
   Begin MSForms.TextBox txtMake
      Height          =   360
      Left            =   240
      TabIndex        =   14
      Top             =   4140
      Width           =   3200
   End
   Begin MSForms.Label lblModel
      Caption         =   "Model"
      Height          =   240
      Left            =   3600
      TabIndex        =   15
      Top             =   3900
      Width           =   3440
   End
   Begin MSForms.TextBox txtModel
      Height          =   360
      Left            =   3600
      TabIndex        =   16
      Top             =   4140
      Width           =   3440
   End
   Begin MSForms.Label lblCapacity
      Caption         =   "Capacity / revision"
      Height          =   240
      Left            =   240
      TabIndex        =   17
      Top             =   4620
      Width           =   6800
   End
   Begin MSForms.TextBox txtCapacityRev
      Height          =   360
      Left            =   240
      TabIndex        =   18
      Top             =   4860
      Width           =   6800
   End
   Begin MSForms.Label lblDescription
      Caption         =   "Description"
      Height          =   240
      Left            =   7320
      TabIndex        =   19
      Top             =   600
      Width           =   6800
   End
   Begin MSForms.TextBox txtDescription
      Height          =   960
      Left            =   7320
      MultiLine       =   -1
      TabIndex        =   20
      Top             =   840
      Width           =   6800
   End
   Begin MSForms.Label lblArrival
      Caption         =   "Arrival date (YYYY-MM-DD)"
      Height          =   240
      Left            =   7320
      TabIndex        =   21
      Top             =   1980
      Width           =   3200
   End
   Begin MSForms.TextBox txtArrivalDate
      Height          =   360
      Left            =   7320
      TabIndex        =   22
      Top             =   2220
      Width           =   3200
   End
   Begin MSForms.Label lblExpiry
      Caption         =   "Expiry date (YYYY-MM-DD)"
      Height          =   240
      Left            =   10680
      TabIndex        =   23
      Top             =   1980
      Width           =   3440
   End
   Begin MSForms.TextBox txtExpiryDate
      Height          =   360
      Left            =   10680
      TabIndex        =   24
      Top             =   2220
      Width           =   3440
   End
   Begin MSForms.Label lblAssignee
      Caption         =   "Initial assignee employee ID (optional)"
      Height          =   240
      Left            =   7320
      TabIndex        =   25
      Top             =   2760
      Width           =   6800
   End
   Begin MSForms.TextBox txtInitialAssignee
      Height          =   360
      Left            =   7320
      TabIndex        =   26
      Top             =   3000
      Width           =   6800
   End
   Begin MSForms.CommandButton btnSave
      Caption         =   "Save"
      Height          =   420
      Left            =   7320
      TabIndex        =   27
      Top             =   3600
      Width           =   1800
   End
   Begin MSForms.CommandButton btnClear
      Caption         =   "Clear"
      Height          =   420
      Left            =   9240
      TabIndex        =   28
      Top             =   3600
      Width           =   1500
   End
   Begin MSForms.CommandButton btnClose
      Caption         =   "Close"
      Cancel          =   -1
      Height          =   420
      Left            =   10860
      TabIndex        =   29
      Top             =   3600
      Width           =   1500
   End
   Begin MSForms.Label lblMessage
      Caption         =   "Scan a barcode. The scanner's Enter key looks the tag up."
      Height          =   1200
      Left            =   7320
      TabIndex        =   30
      Top             =   4200
      Width           =   6800
   End
End
Attribute VB_Name = "frmInward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mExisting As Boolean

Private Sub UserForm_Initialize()
    If Not IMS_IsCustodianRole(gRole) Then
        MsgBox "Only inventory custodians can inward hardware.", vbExclamation, IMS_APP_NAME
        Unload Me
        Exit Sub
    End If
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 16
    IMS_FillCombo cboAssetType, IMS_AssetTypeList(), False
    txtBarcode.MaxLength = 64
    txtInitialAssignee.MaxLength = 32
    ClearForm
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    txtBarcode.SetFocus
End Sub

Private Sub txtBarcode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        ApplyBarcode
    End If
End Sub

Private Sub txtModel_AfterUpdate()
    Dim hint As Object
    If mExisting Then Exit Sub
    If Trim$(txtModel.Text) = "" Then Exit Sub
    Set hint = IMS_SuggestFromModel(txtModel.Text)
    FillEmpty hint
End Sub

Private Sub btnSave_Click()
    Dim fields As Object
    Set fields = CollectFields()
    If mExisting Then
        If IMS_UpdateAssetMetadata(fields) Then
            IMS_Flash lblMessage, "Metadata updated for " & fields("Barcode") & ". Assignee and status were left unchanged.", False
        Else
            IMS_Flash lblMessage, gLastError, True
        End If
    Else
        If IMS_SaveNewAsset(fields) Then
            IMS_Flash lblMessage, "Inward saved for " & fields("Barcode") & ".", False
            ClearForm
        Else
            IMS_Flash lblMessage, gLastError, True
        End If
    End If
End Sub

Private Sub btnClear_Click()
    ClearForm
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub ApplyBarcode()
    Dim code As String
    Dim asset As Object
    Dim hint As Object
    code = IMS_NormalizeBarcode(txtBarcode.Text)
    txtBarcode.Text = code
    If code = "" Then Exit Sub
    If Not IMS_ValidBarcode(code) Then
        IMS_Flash lblMessage, "Barcode must be 4 to 64 letters, digits, hyphens, or underscores.", True
        Exit Sub
    End If
    Set asset = IMS_GetAsset(code)
    If Not asset Is Nothing Then
        mExisting = True
        FillFromAsset asset
        IMS_Flash lblMessage, "This barcode is already on file. Saving updates metadata only. Use Transfer to change the assignee.", False
        Exit Sub
    End If
    mExisting = False
    txtInitialAssignee.Enabled = True
    Set hint = IMS_SuggestFromBarcode(code)
    FillEmpty hint
    If hint.Exists("Model") Then
        IMS_Flash lblMessage, "New barcode. Empty fields were filled from the latest " & IMS_FamilyPrefix(code) & " asset.", False
    Else
        IMS_Flash lblMessage, "New barcode. Fill the purchase order before saving.", False
    End If
End Sub

Private Sub FillFromAsset(ByVal asset As Object)
    SetOwnership DictStr(asset, "Ownership")
    cboAssetType.Value = DictStr(asset, "AssetType")
    txtPONumber.Text = DictStr(asset, "PONumber")
    txtManufacturer.Text = DictStr(asset, "Manufacturer")
    txtMfsSerialNo.Text = DictStr(asset, "MfsSerialNo")
    txtMake.Text = DictStr(asset, "Make")
    txtModel.Text = DictStr(asset, "Model")
    txtCapacityRev.Text = DictStr(asset, "Capacity_Rev")
    txtDescription.Text = DictStr(asset, "Description")
    txtArrivalDate.Text = DateOnly(DictStr(asset, "ArrivalDate"))
    txtExpiryDate.Text = DateOnly(DictStr(asset, "ExpiryDate"))
    txtInitialAssignee.Text = DictStr(asset, "CurrentAssigneeID")
    txtInitialAssignee.Enabled = False
End Sub

Private Sub FillEmpty(ByVal hint As Object)
    If hint Is Nothing Then Exit Sub
    If Len(Trim$(cboAssetType.Value & "")) = 0 And hint.Exists("AssetType") Then cboAssetType.Value = hint("AssetType")
    If txtManufacturer.Text = "" And hint.Exists("Manufacturer") Then txtManufacturer.Text = hint("Manufacturer")
    If txtMake.Text = "" And hint.Exists("Make") Then txtMake.Text = hint("Make")
    If txtModel.Text = "" And hint.Exists("Model") Then txtModel.Text = hint("Model")
    If txtCapacityRev.Text = "" And hint.Exists("Capacity_Rev") Then txtCapacityRev.Text = hint("Capacity_Rev")
End Sub

Private Function CollectFields() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("Barcode") = txtBarcode.Text
    If optAMD.Value Then
        d("Ownership") = "AMD-Specific"
    ElseIf optWipro.Value Then
        d("Ownership") = "Wipro-Specific"
    Else
        d("Ownership") = ""
    End If
    d("AssetType") = cboAssetType.Value
    d("PONumber") = txtPONumber.Text
    d("Manufacturer") = txtManufacturer.Text
    d("MfsSerialNo") = txtMfsSerialNo.Text
    d("Make") = txtMake.Text
    d("Model") = txtModel.Text
    d("Capacity_Rev") = txtCapacityRev.Text
    d("Description") = txtDescription.Text
    d("ArrivalDate") = txtArrivalDate.Text
    d("ExpiryDate") = txtExpiryDate.Text
    d("AssigneeID") = txtInitialAssignee.Text
    Set CollectFields = d
End Function

Private Sub SetOwnership(ByVal value As String)
    optAMD.Value = (value = "AMD-Specific")
    optWipro.Value = (value = "Wipro-Specific")
End Sub

Private Sub ClearForm()
    mExisting = False
    txtBarcode.Text = ""
    optAMD.Value = False
    optWipro.Value = False
    cboAssetType.ListIndex = -1
    txtPONumber.Text = ""
    txtManufacturer.Text = ""
    txtMfsSerialNo.Text = ""
    txtMake.Text = ""
    txtModel.Text = ""
    txtCapacityRev.Text = ""
    txtDescription.Text = ""
    txtArrivalDate.Text = ""
    txtExpiryDate.Text = ""
    txtInitialAssignee.Text = ""
    txtInitialAssignee.Enabled = True
    IMS_Flash lblMessage, "Scan a barcode. The scanner Enter key looks the tag up.", False
End Sub

Private Function DictStr(ByVal d As Object, ByVal key As String) As String
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsNull(d(key)) Then Exit Function
    DictStr = Trim$(CStr(d(key)))
End Function

Private Function DateOnly(ByVal value As String) As String
    If Len(value) >= 10 Then DateOnly = Left$(value, 10) Else DateOnly = value
End Function
