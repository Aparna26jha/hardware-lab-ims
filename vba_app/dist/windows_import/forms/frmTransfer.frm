VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTransfer
   Caption         =   "Hardware Lab IMS - Transfer"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Controls are created in UserForm_Initialize (see modUi).
' No button is the Default button: the scanner sends Enter after each barcode.

Private mLoaded As Boolean
Private mBlocked As Boolean

Private lblMessage As MSForms.Label
Private lblOwnership As MSForms.Label
Private lblType As MSForms.Label
Private lblCurrent As MSForms.Label
Private lblState As MSForms.Label
Private lblHistoryNote As MSForms.Label
Private lstHistory As MSForms.ListBox
Private WithEvents txtBarcode As MSForms.TextBox
Private cboStatus As MSForms.ComboBox
Private txtNewEmpID As MSForms.TextBox
Private txtRemarks As MSForms.TextBox
Private WithEvents btnTransfer As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim widths As Variant
    If Not IMS_IsCustodianRole(gRole) Then
        mBlocked = True
        Exit Sub
    End If
    UiSizeForm Me, "Hardware Lab IMS - Transfer", 790, 400
    UiTitle Me, "lblTitle", "Assign and transfer", 18, 12, 740
    UiHint Me, "lblHint", "Scan the barcode and press Enter. Check who holds it and its history, choose the new status and person, then Record transfer.", 18, 38, 740, 26

    UiLabel Me, "lblBarcode", "Asset barcode *", 18, 70, 320, 12
    Set txtBarcode = UiText(Me, "txtBarcode", 18, 84, 320, 64)
    Set lblOwnership = UiLabel(Me, "lblOwnership", "Ownership:", 18, 110, 320, 14)
    Set lblType = UiLabel(Me, "lblType", "Type / model:", 18, 126, 320, 14)
    Set lblCurrent = UiLabel(Me, "lblCurrent", "Current holder:", 18, 142, 320, 14)
    Set lblState = UiLabel(Me, "lblState", "Status:", 18, 158, 320, 14)
    lblCurrent.Font.Bold = True

    UiLabel Me, "lblStatus", "New status *", 18, 182, 320, 12
    Set cboStatus = UiCombo(Me, "cboStatus", 18, 196, 320)
    IMS_FillCombo cboStatus, IMS_StatusList(), False
    UiPick cboStatus, "Assigned to Engineer"
    UiLabel Me, "lblNewEmp", "Give it to this alias ID (needed for Assigned to Engineer)", 18, 220, 320, 12
    Set txtNewEmpID = UiText(Me, "txtNewEmpID", 18, 234, 320, 32)
    UiLabel Me, "lblRemarks", "Remarks (needed when scrapping or returning)", 18, 258, 320, 12
    Set txtRemarks = UiMemo(Me, "txtRemarks", 18, 272, 320, 36)
    Set btnTransfer = UiButton(Me, "btnTransfer", "Record transfer", 18, 314, 110)
    Set btnClose = UiButton(Me, "btnClose", "Close", 136, 314, 80)
    btnClose.Cancel = True
    Set lblMessage = UiLabel(Me, "lblMessage", "", 18, 342, 320, 30)

    Dim heading As Object
    Set heading = UiLabel(Me, "lblHistory", "History of this asset", 360, 70, 400, 12)
    heading.Font.Bold = True
    Set lblHistoryNote = UiHint(Me, "lblHistoryNote", "Every earlier holder appears here after a scan, newest first.", 360, 84, 400, 12)
    widths = Array(74, 0, 100, 100, 44, 82)
    UiColumnHeads Me, "lblHistCol", Array("When", "", "From", "To", "By", "Remarks"), widths, 360, 100
    Set lstHistory = UiList(Me, "lstHistory", 360, 114, 400, 228, UiWidthsText(widths))
    ClearView
End Sub

Private Sub UserForm_Activate()
    If mBlocked Then
        MsgBox "Only inventory custodians can transfer hardware.", vbExclamation, IMS_APP_NAME
        Unload Me
        Exit Sub
    End If
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
    If IMS_TransferAsset(txtBarcode.Text, txtNewEmpID.Text, txtRemarks.Text, cboStatus.Text) Then
        LoadAsset
        txtNewEmpID.Text = ""
        txtRemarks.Text = ""
        IMS_Flash lblMessage, "Transfer recorded for " & IMS_NormalizeBarcode(txtBarcode.Text) & ".", False
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
    If code = "" Then Exit Sub
    On Error GoTo Fail
    Set asset = IMS_GetAsset(code)
    If asset Is Nothing Then
        ClearView
        IMS_Flash lblMessage, "Barcode " & code & " is not in the lab database. Use Inward to add it.", True
        Exit Sub
    End If
    mLoaded = True
    lblOwnership.Caption = "Ownership: " & DictStr(asset, "Ownership")
    lblType.Caption = "Type / model: " & DictStr(asset, "AssetType") & "  " & DictStr(asset, "Model")
    If DictStr(asset, "CurrentAssigneeID") = "" Then
        lblCurrent.Caption = "Current holder: nobody (in the lab)"
    Else
        lblCurrent.Caption = "Current holder: " & DictStr(asset, "CurrentAssigneeID") & "  " & DictStr(asset, "CurrentAssigneeName")
    End If
    lblState.Caption = "Status: " & DictStr(asset, "Status") & IMS_ExpiryNote(DictStr(asset, "ExpiryDate"))
    LoadHistory code
    IMS_Flash lblMessage, "Asset loaded. Choose the new status and record the transfer.", False
    Exit Sub
Fail:
    IMS_Flash lblMessage, "Could not read the lab database. " & Err.Description, True
End Sub

Private Sub LoadHistory(ByVal code As String)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    lstHistory.Clear
    data = IMS_AssetHistory(code)
    If Not IsArray(data) Then
        If gLastError <> "" Then
            lblHistoryNote.Caption = gLastError
        Else
            lblHistoryNote.Caption = "No earlier movements are recorded for " & code & "."
        End If
        Exit Sub
    End If
    For r = 1 To UBound(data, 1)
        lstHistory.AddItem CStr(data(r, 1))
        For c = 2 To 6
            lstHistory.List(lstHistory.ListCount - 1, c - 1) = CStr(data(r, c))
        Next c
    Next r
    lblHistoryNote.Caption = CStr(UBound(data, 1)) & " recorded movements for " & code & ", newest first."
End Sub

Private Sub ClearView()
    mLoaded = False
    lblOwnership.Caption = "Ownership:"
    lblType.Caption = "Type / model:"
    lblCurrent.Caption = "Current holder:"
    lblState.Caption = "Status:"
    lstHistory.Clear
    lblHistoryNote.Caption = "Every earlier holder appears here after a scan, newest first."
    IMS_Flash lblMessage, "Scan a barcode to begin.", False
End Sub

Private Function DictStr(ByVal d As Object, ByVal key As String) As String
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsNull(d(key)) Then Exit Function
    DictStr = Trim$(CStr(d(key)))
End Function
