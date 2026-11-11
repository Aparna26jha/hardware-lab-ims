VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard
   Caption         =   "Hardware Lab IMS"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15600
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Controls are created in UserForm_Initialize (see modUi).
' The cards and the attention panel use every asset this role may see.
' The filters and search apply to the asset list only.

Private Const FORM_W As Single = 980
Private Const RIGHT_EDGE As Single = 958

Private mReady As Boolean

Private lblWho As MSForms.Label
Private lblHint As MSForms.Label
Private lblCount As MSForms.Label
Private lblMessage As MSForms.Label
Private lblAttnHead As MSForms.Label
Private lblActHead As MSForms.Label
Private cardTotal As MSForms.Label
Private cardAvailable As MSForms.Label
Private cardAssigned As MSForms.Label
Private cardMaint As MSForms.Label
Private cardAMD As MSForms.Label
Private cardWipro As MSForms.Label
Private lstAttention As MSForms.ListBox
Private lstActivity As MSForms.ListBox
Private lstAssets As MSForms.ListBox
Private WithEvents cboOwnership As MSForms.ComboBox
Private WithEvents cboStatus As MSForms.ComboBox
Private WithEvents cboType As MSForms.ComboBox
Private WithEvents txtSearch As MSForms.TextBox
Private WithEvents btnRefresh As MSForms.CommandButton
Private WithEvents btnInward As MSForms.CommandButton
Private WithEvents btnTransfer As MSForms.CommandButton
Private WithEvents btnUsers As MSForms.CommandButton
Private WithEvents btnAccount As MSForms.CommandButton
Private WithEvents btnBackup As MSForms.CommandButton
Private WithEvents btnSignOut As MSForms.CommandButton
Private WithEvents btnExport As MSForms.CommandButton
Private WithEvents btnExportAudit As MSForms.CommandButton
Private WithEvents btnReset As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim widths As Variant
    Dim cardW As Single
    Dim isCustodian As Boolean
    isCustodian = IMS_IsCustodianRole(gRole)
    UiSizeForm Me, "Hardware Lab IMS", FORM_W, 512

    Set lblWho = UiLabel(Me, "lblWho", WhoText(), 18, 10, 440, 16)
    lblWho.Font.Bold = True
    lblWho.Font.Size = 11
    lblWho.ForeColor = UI_TEAL
    lblWho.WordWrap = False
    Set btnInward = UiButton(Me, "btnInward", "Inward", 0, 6, 66)
    Set btnTransfer = UiButton(Me, "btnTransfer", "Transfer", 0, 6, 70)
    Set btnUsers = UiButton(Me, "btnUsers", "Accounts", 0, 6, 70)
    Set btnAccount = UiButton(Me, "btnAccount", "My account", 0, 6, 76)
    Set btnBackup = UiButton(Me, "btnBackup", "Back up now", 0, 6, 84)
    Set btnSignOut = UiButton(Me, "btnSignOut", "Sign out", 0, 6, 66)
    btnInward.Visible = isCustodian
    btnTransfer.Visible = isCustodian
    btnUsers.Visible = IMS_IsSuperRole(gRole)
    btnBackup.Visible = IMS_IsSuperRole(gRole)
    PlaceTopButtons
    Set lblHint = UiHint(Me, "lblHint", HintText(), 18, 30, RIGHT_EDGE - 18, 12)

    cardW = (RIGHT_EDGE - 18 - 5 * 8) / 6
    Set cardTotal = UiCard(Me, "cardTotal", "Total assets", UI_CARD, 18, 46, cardW, 42)
    Set cardAvailable = UiCard(Me, "cardAvailable", "Available in lab", UI_CARD, 18 + 1 * (cardW + 8), 46, cardW, 42)
    Set cardAssigned = UiCard(Me, "cardAssigned", "Assigned to engineers", UI_CARD, 18 + 2 * (cardW + 8), 46, cardW, 42)
    Set cardMaint = UiCard(Me, "cardMaint", "Under maintenance", UI_CARD, 18 + 3 * (cardW + 8), 46, cardW, 42)
    Set cardAMD = UiCard(Me, "cardAMD", "AMD-Specific", UI_AMD, 18 + 4 * (cardW + 8), 46, cardW, 42)
    Set cardWipro = UiCard(Me, "cardWipro", "Wipro-Specific", UI_WIPRO, 18 + 5 * (cardW + 8), 46, cardW, 42)

    Set lblAttnHead = UiLabel(Me, "lblAttnHead", "Needs your attention", 18, 94, 464, 12)
    lblAttnHead.Font.Bold = True
    If isCustodian Then
        Set lstAttention = UiList(Me, "lstAttention", 18, 108, 464, 74, "104 pt;350 pt")
        Set lblActHead = UiLabel(Me, "lblActHead", "Recent activity (last 10 moves)", 494, 94, 464, 12)
        lblActHead.Font.Bold = True
        Set lstActivity = UiList(Me, "lstActivity", 494, 108, 464, 74, "66 pt;86 pt;92 pt;92 pt;40 pt;80 pt")
    Else
        Set lstAttention = UiList(Me, "lstAttention", 18, 108, RIGHT_EDGE - 18, 74, "104 pt;820 pt")
    End If

    UiLabel Me, "lblOwner", "Ownership", 18, 190, 150, 12
    Set cboOwnership = UiCombo(Me, "cboOwnership", 18, 202, 150)
    UiLabel Me, "lblStatusFilter", "Status", 176, 190, 170, 12
    Set cboStatus = UiCombo(Me, "cboStatus", 176, 202, 170)
    UiLabel Me, "lblTypeFilter", "Asset type", 354, 190, 150, 12
    Set cboType = UiCombo(Me, "cboType", 354, 202, 150)
    UiLabel Me, "lblSearch", "Search barcode, serial, model, PO, name, or description, then press Enter", 512, 190, 360, 12
    Set txtSearch = UiText(Me, "txtSearch", 512, 202, 360, 64)
    Set btnRefresh = UiButton(Me, "btnRefresh", "Search", 880, 200, 78)
    IMS_FillCombo cboOwnership, IMS_OwnershipList(), True
    IMS_FillCombo cboStatus, IMS_StatusList(), True
    IMS_FillCombo cboType, IMS_AssetTypeList(), True

    widths = Array(100, 84, 80, 90, 90, 90, 160, 66, 134)
    UiColumnHeads Me, "lblCol", Array("Barcode", "Ownership", "Type", "Model", "Serial", "PO", "Status", "Holder", "Holder name"), widths, 18, 228
    Set lstAssets = UiList(Me, "lstAssets", 18, 242, RIGHT_EDGE - 18, 176, UiWidthsText(widths))

    Set lblCount = UiLabel(Me, "lblCount", "0 assets", 18, 426, 300, 14)
    Set btnExport = UiButton(Me, "btnExport", "Export inventory CSV", 0, 422, 130)
    Set btnExportAudit = UiButton(Me, "btnExportAudit", "Export audit CSV", 0, 422, 116)
    Set btnReset = UiButton(Me, "btnReset", "Reset sample data", 0, 422, 116)
    btnExportAudit.Visible = isCustodian
    btnReset.Visible = IMS_IsSuperRole(gRole)
    PlaceBottomButtons
    Set lblMessage = UiLabel(Me, "lblMessage", "", 18, 450, RIGHT_EDGE - 18, 28)
    mReady = True
End Sub

Private Sub UserForm_Activate()
    RefreshAll
End Sub

Private Sub btnRefresh_Click()
    LoadAssets
End Sub

Private Sub cboOwnership_Change()
    If mReady Then LoadAssets
End Sub

Private Sub cboStatus_Change()
    If mReady Then LoadAssets
End Sub

Private Sub cboType_Change()
    If mReady Then LoadAssets
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        LoadAssets
    End If
End Sub

Private Sub btnInward_Click()
    frmInward.Show
    RefreshAll
End Sub

Private Sub btnTransfer_Click()
    frmTransfer.Show
    RefreshAll
End Sub

Private Sub btnUsers_Click()
    frmUserAdmin.Show
    RefreshAll
End Sub

Private Sub btnAccount_Click()
    frmAccount.Show
End Sub

Private Sub btnBackup_Click()
    Dim saved As String
    saved = IMS_BackupNow()
    If saved = "" Then
        IMS_Flash lblMessage, gLastError, True
    Else
        IMS_Flash lblMessage, "Backup saved: " & saved, False
    End If
End Sub

Private Sub btnSignOut_Click()
    IMS_SignOut
    Unload Me
End Sub

Private Sub btnReset_Click()
    IMS_ResetBackend
    RefreshAll
End Sub

Private Sub btnExport_Click()
    Dim data As Variant
    Dim path As Variant
    On Error GoTo Fail
    data = CurrentRows()
    path = IMS_SuggestExportPath("LabInventory")
    If VarType(path) = vbBoolean Then Exit Sub
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
    Dim c As Long
    Dim kept As Long
    Dim i As Long
    On Error GoTo Fail
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
                    For c = 1 To 7
                        filtered(i, c) = auditRows(r, c)
                    Next c
                End If
            Next r
        End If
    End If
    path = IMS_SuggestExportPath("LabAudit")
    If VarType(path) = vbBoolean Then Exit Sub
    IMS_WriteCsv CStr(path), IMS_AuditHeaders(), filtered
    IMS_Flash lblMessage, "Audit snapshot for the assets on screen written to " & CStr(path), False
    Exit Sub
Fail:
    IMS_Flash lblMessage, Err.Description, True
End Sub

Private Sub RefreshAll()
    LoadSummary
    LoadAssets
End Sub

Private Sub LoadSummary()
    Dim everything As Variant
    Dim counts As Object
    On Error GoTo Fail
    everything = IMS_QueryAssets("", "", "", "", "")
    Set counts = IMS_CountAssets(everything)
    cardTotal.Caption = CStr(counts("Total"))
    If counts("Scrapped") > 0 Then
        Me.Controls("cardTotalCap").Caption = "Total assets (" & CStr(counts("Scrapped")) & " scrapped)"
    Else
        Me.Controls("cardTotalCap").Caption = "Total assets"
    End If
    cardAvailable.Caption = CStr(counts("Available"))
    cardAssigned.Caption = CStr(counts("Assigned"))
    cardMaint.Caption = CStr(counts("Maintenance"))
    cardAMD.Caption = CStr(counts("AMD"))
    cardWipro.Caption = CStr(counts("Wipro"))
    LoadAttention everything
    If Not lstActivity Is Nothing Then LoadActivity
    Exit Sub
Fail:
    IMS_Flash lblMessage, "Could not read the summary. " & Err.Description, True
End Sub

Private Sub LoadAttention(ByVal everything As Variant)
    Dim pending As Variant
    Dim expiring As Variant
    Dim r As Long
    lstAttention.Clear
    If IMS_IsSuperRole(gRole) Then
        pending = IMS_PendingRequests()
        If IsArray(pending) Then
            For r = 1 To UBound(pending, 1)
                AddAttention CStr(pending(r, 1)), CStr(pending(r, 2)) & " is waiting for approval (asked for " & CStr(pending(r, 3)) & "). Open Accounts to approve or reject."
            Next r
        End If
    End If
    expiring = IMS_ExpiryRows(everything)
    If IsArray(expiring) Then
        For r = 1 To UBound(expiring, 1)
            AddAttention CStr(expiring(r, 1)), CStr(expiring(r, 2)) & "  -  " & CStr(expiring(r, 3)) & "  -  " & CStr(expiring(r, 4))
        Next r
    End If
    If lstAttention.ListCount = 0 Then
        AddAttention "All clear", "Nothing needs your attention right now."
        lblAttnHead.Caption = "Needs your attention"
    Else
        lblAttnHead.Caption = "Needs your attention (" & CStr(lstAttention.ListCount) & ")"
    End If
End Sub

Private Sub AddAttention(ByVal item As String, ByVal detail As String)
    lstAttention.AddItem item
    lstAttention.List(lstAttention.ListCount - 1, 1) = detail
End Sub

Private Sub LoadActivity()
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    lstActivity.Clear
    data = IMS_RecentActivity(10)
    If Not IsArray(data) Then
        lstActivity.AddItem ""
        If gLastError <> "" Then
            lstActivity.List(0, 1) = gLastError
        Else
            lstActivity.List(0, 1) = "No moves recorded yet."
        End If
        Exit Sub
    End If
    For r = 1 To UBound(data, 1)
        lstActivity.AddItem CStr(data(r, 1))
        For c = 2 To 6
            lstActivity.List(lstActivity.ListCount - 1, c - 1) = CStr(data(r, c))
        Next c
    Next r
End Sub

Private Sub LoadAssets()
    Dim data As Variant
    Dim r As Long
    On Error GoTo Fail
    data = CurrentRows()
    lstAssets.Clear
    If Not IsArray(data) Then
        lblCount.Caption = "0 assets match this view"
        Exit Sub
    End If
    For r = 1 To UBound(data, 1)
        lstAssets.AddItem CStr(data(r, 1))
        lstAssets.List(lstAssets.ListCount - 1, 1) = CStr(data(r, 2))
        lstAssets.List(lstAssets.ListCount - 1, 2) = CStr(data(r, 3))
        lstAssets.List(lstAssets.ListCount - 1, 3) = CStr(data(r, 4))
        lstAssets.List(lstAssets.ListCount - 1, 4) = CStr(data(r, 10))
        lstAssets.List(lstAssets.ListCount - 1, 5) = CStr(data(r, 5))
        lstAssets.List(lstAssets.ListCount - 1, 6) = CStr(data(r, 6))
        lstAssets.List(lstAssets.ListCount - 1, 7) = CStr(data(r, 7))
        lstAssets.List(lstAssets.ListCount - 1, 8) = CStr(data(r, 8))
    Next r
    lblCount.Caption = CStr(UBound(data, 1)) & " assets match this view"
    Exit Sub
Fail:
    lblCount.Caption = "0 assets match this view"
    IMS_Flash lblMessage, Err.Description, True
End Sub

Private Function CurrentRows() As Variant
    CurrentRows = IMS_QueryAssets(cboOwnership.Text, cboStatus.Text, cboType.Text, txtSearch.Text, "")
End Function

Private Function WhoText() As String
    Select Case gRole
        Case "User": WhoText = gEmpName & "  (" & gEmpID & ")  -  Engineer view: your assigned hardware"
        Case "Lead Admin": WhoText = gEmpName & "  (" & gEmpID & ")  -  Lead view: your team's hardware"
        Case "Admin": WhoText = gEmpName & "  (" & gEmpID & ")  -  Inventory custodian"
        Case Else: WhoText = gEmpName & "  (" & gEmpID & ")  -  " & gRole
    End Select
End Function

Private Function HintText() As String
    Select Case gRole
        Case "Super Admin"
            HintText = "Approve waiting requests in Accounts. Inward adds hardware, Transfer hands it out or takes it back. Back up now copies the database to the backup folder."
        Case "Admin"
            HintText = "Inward adds new hardware. Transfer hands it out or takes it back. Search finds any barcode, serial, model, PO, or person."
        Case "Lead Admin"
            HintText = "You see hardware held by you and your team. Search, filter, or export the list. The inventory custodian moves hardware."
        Case Else
            HintText = "You see hardware assigned to you. To return or swap a board, contact the inventory custodian."
    End Select
End Function

' Lines up the visible top buttons from the right edge, so hidden ones leave no gap.
Private Sub PlaceTopButtons()
    Dim order As Variant
    Dim i As Long
    Dim x As Single
    order = Array(btnSignOut, btnAccount, btnBackup, btnUsers, btnTransfer, btnInward)
    x = RIGHT_EDGE
    For i = LBound(order) To UBound(order)
        If order(i).Visible Then
            x = x - order(i).Width
            order(i).Left = x
            x = x - 6
        End If
    Next i
End Sub

Private Sub PlaceBottomButtons()
    Dim order As Variant
    Dim i As Long
    Dim x As Single
    order = Array(btnReset, btnExportAudit, btnExport)
    x = RIGHT_EDGE
    For i = LBound(order) To UBound(order)
        If order(i).Visible Then
            x = x - order(i).Width
            order(i).Left = x
            x = x - 6
        End If
    Next i
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then IMS_SignOut
End Sub
