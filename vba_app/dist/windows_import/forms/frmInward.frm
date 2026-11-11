VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInward
   Caption         =   "Hardware Lab IMS - Inward"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Controls are created in UserForm_Initialize (see modUi).
' No button is the Default button: the scanner sends Enter after each barcode.

Private mExisting As Boolean
Private mBlocked As Boolean

Private lblMessage As MSForms.Label
Private WithEvents txtBarcode As MSForms.TextBox
Private optAMD As MSForms.OptionButton
Private optWipro As MSForms.OptionButton
Private cboAssetType As MSForms.ComboBox
Private txtPONumber As MSForms.TextBox
Private txtManufacturer As MSForms.TextBox
Private txtMfsSerialNo As MSForms.TextBox
Private txtMake As MSForms.TextBox
Private WithEvents txtModel As MSForms.TextBox
Private txtCapacityRev As MSForms.TextBox
Private txtDescription As MSForms.TextBox
Private txtArrivalDate As MSForms.TextBox
Private txtExpiryDate As MSForms.TextBox
Private txtInitialAssignee As MSForms.TextBox
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnClear As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    If Not IMS_IsCustodianRole(gRole) Then
        mBlocked = True
        Exit Sub
    End If
    UiSizeForm Me, "Hardware Lab IMS - Inward", 700, 384
    UiTitle Me, "lblTitle", "Inward material", 18, 12, 660
    UiHint Me, "lblHint", "Scan the barcode and let the scanner press Enter. A known board fills in its details. Fields marked * are required.", 18, 38, 660, 26

    UiLabel Me, "lblBarcode", "Barcode / asset tag *", 18, 70, 320, 12
    Set txtBarcode = UiText(Me, "txtBarcode", 18, 84, 320, 64)
    UiLabel Me, "lblOwner", "Ownership *", 18, 108, 320, 12
    Set optAMD = UiOption(Me, "optAMD", "AMD-Specific", "Ownership", 18, 122, 130)
    Set optWipro = UiOption(Me, "optWipro", "Wipro-Specific", "Ownership", 160, 122, 140)
    UiLabel Me, "lblAssetType", "Asset type *", 18, 146, 320, 12
    Set cboAssetType = UiCombo(Me, "cboAssetType", 18, 160, 320)
    IMS_FillCombo cboAssetType, IMS_AssetTypeList(), False
    UiLabel Me, "lblPO", "Purchase order *", 18, 184, 320, 12
    Set txtPONumber = UiText(Me, "txtPONumber", 18, 198, 320, 64)
    UiLabel Me, "lblManufacturer", "Manufacturer", 18, 222, 155, 12
    Set txtManufacturer = UiText(Me, "txtManufacturer", 18, 236, 155, 120)
    UiLabel Me, "lblMfs", "Manufacturer serial", 183, 222, 155, 12
    Set txtMfsSerialNo = UiText(Me, "txtMfsSerialNo", 183, 236, 155, 64)
    UiLabel Me, "lblMake", "Make", 18, 260, 155, 12
    Set txtMake = UiText(Me, "txtMake", 18, 274, 155, 80)
    UiLabel Me, "lblModel", "Model", 183, 260, 155, 12
    Set txtModel = UiText(Me, "txtModel", 183, 274, 155, 80)
    UiLabel Me, "lblCapacity", "Capacity / revision", 18, 298, 320, 12
    Set txtCapacityRev = UiText(Me, "txtCapacityRev", 18, 312, 320, 80)

    UiLabel Me, "lblDescription", "Description", 360, 70, 320, 12
    Set txtDescription = UiMemo(Me, "txtDescription", 360, 84, 320, 56)
    UiLabel Me, "lblArrival", "Arrival date (YYYY-MM-DD)", 360, 146, 155, 12
    Set txtArrivalDate = UiText(Me, "txtArrivalDate", 360, 160, 155, 10)
    UiLabel Me, "lblExpiry", "Expiry date (YYYY-MM-DD)", 525, 146, 155, 12
    Set txtExpiryDate = UiText(Me, "txtExpiryDate", 525, 160, 155, 10)
    UiLabel Me, "lblAssignee", "Give it to this alias ID now (optional)", 360, 184, 320, 12
    Set txtInitialAssignee = UiText(Me, "txtInitialAssignee", 360, 198, 320, 32)
    Set btnSave = UiButton(Me, "btnSave", "Save", 360, 230, 90)
    Set btnClear = UiButton(Me, "btnClear", "Clear", 458, 230, 80)
    Set btnClose = UiButton(Me, "btnClose", "Close", 546, 230, 80)
    btnClose.Cancel = True
    Set lblMessage = UiLabel(Me, "lblMessage", "", 360, 262, 320, 70)
    ClearForm
End Sub

Private Sub UserForm_Activate()
    If mBlocked Then
        MsgBox "Only inventory custodians can inward hardware.", vbExclamation, IMS_APP_NAME
        Unload Me
        Exit Sub
    End If
    On Error Resume Next
    txtBarcode.SetFocus
End Sub

Private Sub txtBarcode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        ApplyBarcode
    End If
End Sub

Private Sub txtModel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then ApplyModel
End Sub

Private Sub btnSave_Click()
    Dim fields As Object
    Set fields = CollectFields()
    If mExisting Then
        If IMS_UpdateAssetMetadata(fields) Then
            IMS_Flash lblMessage, "Details updated for " & fields("Barcode") & ". The holder and status did not change.", False
        Else
            IMS_Flash lblMessage, gLastError, True
        End If
    Else
        If IMS_SaveNewAsset(fields) Then
            ClearForm
            IMS_Flash lblMessage, "Inward saved for " & fields("Barcode") & ". Scan the next barcode.", False
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
    On Error GoTo Fail
    Set asset = IMS_GetAsset(code)
    If Not asset Is Nothing Then
        mExisting = True
        FillFromAsset asset
        IMS_Flash lblMessage, "This barcode is already on file. Save updates its details only. Use Transfer to change who holds it.", False
        Exit Sub
    End If
    mExisting = False
    txtInitialAssignee.Enabled = True
    Set hint = IMS_SuggestFromBarcode(code)
    FillEmpty hint
    If hint.Exists("Model") Then
        IMS_Flash lblMessage, "New barcode. Empty fields were filled from the latest " & IMS_FamilyPrefix(code) & " asset. Check them, add the PO, and Save.", False
    Else
        IMS_Flash lblMessage, "New barcode. Choose AMD or Wipro, the type, and the purchase order, then Save.", False
    End If
    Exit Sub
Fail:
    IMS_Flash lblMessage, "Could not read the lab database. " & Err.Description, True
End Sub

Private Sub ApplyModel()
    Dim hint As Object
    If mExisting Then Exit Sub
    If Trim$(txtModel.Text) = "" Then Exit Sub
    On Error Resume Next
    Set hint = IMS_SuggestFromModel(txtModel.Text)
    FillEmpty hint
End Sub

Private Sub FillFromAsset(ByVal asset As Object)
    SetOwnership DictStr(asset, "Ownership")
    UiPick cboAssetType, DictStr(asset, "AssetType")
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
    If cboAssetType.ListIndex < 0 And hint.Exists("AssetType") Then UiPick cboAssetType, CStr(hint("AssetType"))
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
    d("AssetType") = cboAssetType.Text
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
    IMS_Flash lblMessage, "Scan a barcode to begin.", False
    On Error Resume Next
    txtBarcode.SetFocus
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
