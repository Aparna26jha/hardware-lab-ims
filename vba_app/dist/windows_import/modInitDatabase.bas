Attribute VB_Name = "modInitDatabase"
Option Explicit

' One-click setup. Run IMS_InitializeBackend from Alt+F8 on the admin PC.
' The macro creates IMS_Backend.accdb at BackendPath, builds the three tables,
' and loads sample users and assets only when the user table is empty.

Public Sub IMS_InitializeBackend()
    Dim conn As Object
    On Error GoTo Fail
    If Not IMS_LoadConfig() Then
        MsgBox gLastError, vbExclamation, IMS_APP_NAME
        Exit Sub
    End If

    IMS_CreateDatabaseFile
    Set conn = IMS_Connect()
    IMS_EnsureSchema conn
    IMS_Disconnect conn
    Set conn = Nothing

    If IMS_UserCount() = 0 Then
        InsertSeed
        MsgBox "Lab database created and sample data loaded." & vbCrLf & vbCrLf & gBackendPath, vbInformation, IMS_APP_NAME
    Else
        MsgBox "Database is ready. Existing users were left in place, so sample data was not loaded again." & vbCrLf & vbCrLf & gBackendPath, vbInformation, IMS_APP_NAME
    End If
    Exit Sub
Fail:
    MsgBox "Setup failed. " & Err.Description, vbCritical, IMS_APP_NAME
    If Not conn Is Nothing Then IMS_Disconnect conn
End Sub

Public Sub IMS_ResetBackend()
    Dim conn As Object
    Dim typed As String
    On Error GoTo Fail
    If Not IMS_IsSuperRole(gRole) Then
        MsgBox "Only a Super Admin can reset the lab database.", vbExclamation, IMS_APP_NAME
        Exit Sub
    End If
    typed = UCase$(Trim$(InputBox("This erases every asset, transfer, and account, then reloads the sample set." & vbCrLf & "Type RESET to continue.", "Reset lab database")))
    If typed <> "RESET" Then Exit Sub

    Set conn = IMS_Connect()
    conn.BeginTrans
    IMS_DeleteAllRows conn
    conn.CommitTrans
    IMS_Disconnect conn
    Set conn = Nothing
    InsertSeed
    MsgBox "Sample lab data has been restored.", vbInformation, IMS_APP_NAME
    Exit Sub
Fail:
    Dim failText As String
    failText = Err.Description
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.RollbackTrans
        IMS_Disconnect conn
    End If
    MsgBox "Reset failed. " & failText, vbCritical, IMS_APP_NAME
End Sub

Private Sub InsertSeed()
    Dim conn As Object
    Dim users As Variant
    Dim assets As Variant
    Dim auditRows As Variant
    Dim r As Long
    Dim flagSql As String

    Set conn = IMS_Connect()
    On Error GoTo Fail

    users = SeedUsers()
    For r = 1 To UBound(users, 1)
        If CBool(users(r, 5)) Then flagSql = "True" Else flagSql = "False"
        IMS_ExecOn conn, _
            "INSERT INTO [User_Directory] ([EmpID], [EmpName], [Role], [AccessCode], [Active]) VALUES (?, ?, ?, ?, " & flagSql & ")", _
            Array(CStr(users(r, 1)), CStr(users(r, 2)), CStr(users(r, 3)), CStr(users(r, 4)))
    Next r

    assets = SeedAssets()
    For r = 1 To UBound(assets, 1)
        IMS_ExecOn conn, AssetInsertSql(assets, r)
    Next r

    auditRows = SeedAudit()
    For r = 1 To UBound(auditRows, 1)
        IMS_ExecOn conn, _
            "INSERT INTO [Transfer_Audit_Log] ([Timestamp], [AssetBarcode], [PreviousAssigneeID], [NewAssigneeID], [TransferredBy_AdminID], [Remarks]) " & _
            "VALUES (" & IMS_SqlStampOrNull(CStr(auditRows(r, 1))) & ", ?, ?, ?, ?, ?)", _
            Array(CStr(auditRows(r, 2)), CStr(auditRows(r, 3)), CStr(auditRows(r, 4)), CStr(auditRows(r, 5)), CStr(auditRows(r, 6)))
    Next r

    IMS_Disconnect conn
    Exit Sub
Fail:
    Dim failNumber As Long
    Dim failText As String
    failNumber = Err.Number
    failText = Err.Description
    If Not conn Is Nothing Then IMS_Disconnect conn
    Err.Raise failNumber, "IMS", failText
End Sub

Private Function AssetInsertSql(ByVal assets As Variant, ByVal r As Long) As String
    ' Columns from SeedAssets: barcode, ownership, type, manufacturer, mfr serial, make, model,
    ' capacity, description, PO, arrival, expiry, assignee id, assignee name, status.
    AssetInsertSql = _
        "INSERT INTO [Assets_Master] (" & _
        "[Barcode], [Ownership], [AssetType], [Manufacturer], [MfsSerialNo], [Make], [Model], " & _
        "[Capacity_Rev], [Description], [PONumber], [ArrivalDate], [ExpiryDate], " & _
        "[CurrentAssigneeID], [CurrentAssigneeName], [Status], [CreatedBy], [CreatedAt]) VALUES (" & _
        Q(CStr(assets(r, 1))) & ", " & Q(CStr(assets(r, 2))) & ", " & Q(CStr(assets(r, 3))) & ", " & _
        Q(CStr(assets(r, 4))) & ", " & Q(CStr(assets(r, 5))) & ", " & Q(CStr(assets(r, 6))) & ", " & _
        Q(CStr(assets(r, 7))) & ", " & Q(CStr(assets(r, 8))) & ", " & Q(CStr(assets(r, 9))) & ", " & _
        Q(CStr(assets(r, 10))) & ", " & IMS_SqlDateOrNull(CStr(assets(r, 11))) & ", " & _
        IMS_SqlDateOrNull(CStr(assets(r, 12))) & ", " & Q(CStr(assets(r, 13))) & ", " & _
        Q(CStr(assets(r, 14))) & ", " & Q(CStr(assets(r, 15))) & ", 'SEED', Now())"
End Function

Private Function Q(ByVal value As String) As String
    Q = "'" & Replace(Replace(value, vbNullChar, ""), "'", "''") & "'"
End Function
