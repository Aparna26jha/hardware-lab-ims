Attribute VB_Name = "modADODB_Backend"
Option Explicit

' Hardware Lab IMS backend.
' Late-bound ADODB only. Do not add a Tools > References entry.
' Each call opens the Access file, runs one short statement, and drops the connection.
' OLE DB connection pooling is turned off so the .laccdb lock file is released immediately.
' Pooling is useful for a server database. On a shared Access file it keeps the file locked
' after Close and is the usual cause of "file already in use" across a lab.

Private Const SQL_USERS As String = _
    "CREATE TABLE [User_Directory] (" & _
    "[EmpID] TEXT(32) PRIMARY KEY, " & _
    "[EmpName] TEXT(120), [Role] TEXT(32), [AccessCode] TEXT(80), [Active] YESNO, " & _
    "[Email] TEXT(120), [RequestedRole] TEXT(32), [Approval] TEXT(16), " & _
    "[LeadID] TEXT(32), [WindowsUser] TEXT(64))"

Private Const SQL_DIRECTORY As String = _
    "CREATE TABLE [Allowed_Directory] (" & _
    "[Alias] TEXT(32) PRIMARY KEY, [Email] TEXT(120), [FullName] TEXT(120))"

Private Const SQL_ASSETS As String = _
    "CREATE TABLE [Assets_Master] (" & _
    "[Barcode] TEXT(64) PRIMARY KEY, " & _
    "[Ownership] TEXT(32), [AssetType] TEXT(64), [Manufacturer] TEXT(120), " & _
    "[MfsSerialNo] TEXT(64), [Make] TEXT(80), [Model] TEXT(80), [Capacity_Rev] TEXT(80), " & _
    "[Description] MEMO, [PONumber] TEXT(64), [ArrivalDate] DATETIME, [ExpiryDate] DATETIME, " & _
    "[CurrentAssigneeID] TEXT(32), [CurrentAssigneeName] TEXT(120), [Status] TEXT(40), " & _
    "[CreatedBy] TEXT(32), [CreatedAt] DATETIME, [Revision] LONG)"

Private Const SQL_LOG As String = _
    "CREATE TABLE [Transfer_Audit_Log] (" & _
    "[LogID] COUNTER PRIMARY KEY, " & _
    "[Timestamp] DATETIME, [AssetBarcode] TEXT(64), [PreviousAssigneeID] TEXT(32), " & _
    "[NewAssigneeID] TEXT(32), [TransferredBy_AdminID] TEXT(32), [Remarks] MEMO)"

Public Function IMS_LoadConfig() As Boolean
    On Error GoTo Fail
    gLastError = ""
    gBackendPath = Trim$(CStr(ThisWorkbook.Names("BackendPath").RefersToRange.Value))
    gBackupPath = ""
    On Error Resume Next
    gBackupPath = Trim$(CStr(ThisWorkbook.Names("BackupFolder").RefersToRange.Value))
    Err.Clear
    On Error GoTo Fail
    If Len(gBackendPath) < 6 Then
        gLastError = "Set the BackendPath cell on the Config sheet to the full path of IMS_Backend.accdb on the lab share."
        Exit Function
    End If
    If LCase$(Right$(gBackendPath, 6)) <> ".accdb" Then
        gLastError = "BackendPath must end with .accdb."
        Exit Function
    End If
    IMS_LoadConfig = True
    Exit Function
Fail:
    gLastError = "This workbook is missing the BackendPath name. Use the seed workbook from build_workbook.py."
End Function

Public Function IMS_BackendFileExists() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    IMS_BackendFileExists = fso.FileExists(gBackendPath)
End Function

Public Function IMS_Connect() As Object
    Dim conn As Object
    Dim attempt As Long
    Dim lastErr As Long
    Dim lastDesc As String

    If Len(gBackendPath) = 0 Then
        Err.Raise vbObjectError + 1, "IMS", "Backend path is empty."
    End If

    For attempt = 1 To 4
        On Error Resume Next
        Set conn = CreateObject("ADODB.Connection")
        conn.ConnectionTimeout = 15
        conn.CommandTimeout = 20
        conn.Mode = 16 ' adModeShareDenyNone
        conn.Open ConnString()
        lastErr = Err.Number
        lastDesc = Err.Description
        On Error GoTo 0
        If lastErr = 0 Then
            Set IMS_Connect = conn
            Exit Function
        End If
        Set conn = Nothing
        If Not IsLockError(lastErr) Or attempt = 4 Then
            Err.Raise vbObjectError + 2, "IMS", ProviderMessage(lastDesc)
        End If
        PauseBrief 0.15 * attempt
    Next attempt
End Function

Public Sub IMS_Disconnect(ByRef conn As Object)
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
    End If
    Set conn = Nothing
End Sub

Public Function IMS_OpenRs(ByVal conn As Object, ByVal sql As String, Optional ByVal params As Variant) As Object
    Dim cmd As Object
    Dim rs As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = 1 ' adCmdText
    cmd.CommandText = sql
    BindParams cmd, params
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    rs.Open cmd, , 3, 1 ' adOpenStatic, adLockReadOnly
    Set IMS_OpenRs = rs
End Function

Public Sub IMS_ExecOn(ByVal conn As Object, ByVal sql As String, Optional ByVal params As Variant)
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = 1
    cmd.CommandText = sql
    BindParams cmd, params
    cmd.Execute
End Sub

Public Function IMS_FetchOne(ByVal sql As String, Optional ByVal params As Variant) As Object
    Dim conn As Object
    Dim rs As Object
    Set conn = IMS_Connect()
    If IsMissing(params) Then
        Set rs = IMS_OpenRs(conn, sql)
    Else
        Set rs = IMS_OpenRs(conn, sql, params)
    End If
    If rs.EOF Then
        Set IMS_FetchOne = Nothing
    Else
        Set IMS_FetchOne = RsToDict(rs)
    End If
    rs.Close
    IMS_Disconnect conn
End Function

Public Function IMS_FetchDicts(ByVal sql As String, Optional ByVal params As Variant) As Collection
    Dim conn As Object
    Dim rs As Object
    Dim rows As Collection
    Set rows = New Collection
    Set conn = IMS_Connect()
    If IsMissing(params) Then
        Set rs = IMS_OpenRs(conn, sql)
    Else
        Set rs = IMS_OpenRs(conn, sql, params)
    End If
    Do While Not rs.EOF
        rows.Add RsToDict(rs)
        rs.MoveNext
    Loop
    rs.Close
    IMS_Disconnect conn
    Set IMS_FetchDicts = rows
End Function

Public Sub IMS_CreateDatabaseFile()
    Dim cat As Object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(gBackendPath) Then Exit Sub
    EnsureParentFolder gBackendPath
    Set cat = CreateObject("ADOX.Catalog")
    cat.Create ConnString()
    Set cat = Nothing
End Sub

Public Sub IMS_EnsureSchema(ByVal conn As Object)
    If Not TableExists(conn, "User_Directory") Then IMS_ExecOn conn, SQL_USERS
    If Not TableExists(conn, "Assets_Master") Then IMS_ExecOn conn, SQL_ASSETS
    If Not TableExists(conn, "Transfer_Audit_Log") Then IMS_ExecOn conn, SQL_LOG
    If Not TableExists(conn, "Allowed_Directory") Then IMS_ExecOn conn, SQL_DIRECTORY
    EnsureIndex conn, "ALTER TABLE [User_Directory] ALTER COLUMN [AccessCode] TEXT(80)"
    EnsureIndex conn, "ALTER TABLE [User_Directory] ADD COLUMN [Email] TEXT(120)"
    EnsureIndex conn, "ALTER TABLE [User_Directory] ADD COLUMN [RequestedRole] TEXT(32)"
    EnsureIndex conn, "ALTER TABLE [User_Directory] ADD COLUMN [Approval] TEXT(16)"
    EnsureIndex conn, "ALTER TABLE [User_Directory] ADD COLUMN [LeadID] TEXT(32)"
    EnsureIndex conn, "ALTER TABLE [User_Directory] ADD COLUMN [WindowsUser] TEXT(64)"
    EnsureIndex conn, "UPDATE [User_Directory] SET [Approval]='Approved' WHERE [Approval] IS NULL OR [Approval]=''"
    EnsureIndex conn, "ALTER TABLE [Assets_Master] ADD COLUMN [Revision] LONG"
    EnsureIndex conn, "UPDATE [Assets_Master] SET [Revision]=0 WHERE [Revision] IS NULL"
    EnsureIndex conn, "UPDATE [Assets_Master] SET [CurrentAssigneeID]=NULL WHERE [CurrentAssigneeID]=''"
    EnsureIndex conn, "UPDATE [Assets_Master] SET [MfsSerialNo]=NULL WHERE [MfsSerialNo]=''"
    EnsureIndex conn, "UPDATE [Transfer_Audit_Log] SET [PreviousAssigneeID]=NULL WHERE [PreviousAssigneeID]=''"
    EnsureIndex conn, "UPDATE [Transfer_Audit_Log] SET [NewAssigneeID]=NULL WHERE [NewAssigneeID]=''"
    EnsureIndex conn, "CREATE INDEX [idx_asset_assignee] ON [Assets_Master] ([CurrentAssigneeID])"
    EnsureIndex conn, "CREATE INDEX [idx_asset_owner] ON [Assets_Master] ([Ownership])"
    EnsureIndex conn, "CREATE INDEX [idx_asset_status] ON [Assets_Master] ([Status])"
    EnsureIndex conn, "CREATE INDEX [idx_asset_type] ON [Assets_Master] ([AssetType])"
    EnsureIndex conn, "CREATE UNIQUE INDEX [idx_asset_serial] ON [Assets_Master] ([MfsSerialNo])"
    EnsureIndex conn, "CREATE INDEX [idx_log_barcode] ON [Transfer_Audit_Log] ([AssetBarcode])"
    EnsureIndex conn, "ALTER TABLE [Assets_Master] ADD CONSTRAINT [fk_asset_assignee] FOREIGN KEY ([CurrentAssigneeID]) REFERENCES [User_Directory] ([EmpID])"
    EnsureIndex conn, "ALTER TABLE [Transfer_Audit_Log] ADD CONSTRAINT [fk_log_barcode] FOREIGN KEY ([AssetBarcode]) REFERENCES [Assets_Master] ([Barcode])"
    EnsureIndex conn, "ALTER TABLE [Transfer_Audit_Log] ADD CONSTRAINT [fk_log_prev] FOREIGN KEY ([PreviousAssigneeID]) REFERENCES [User_Directory] ([EmpID])"
    EnsureIndex conn, "ALTER TABLE [Transfer_Audit_Log] ADD CONSTRAINT [fk_log_new] FOREIGN KEY ([NewAssigneeID]) REFERENCES [User_Directory] ([EmpID])"
    EnsureIndex conn, "ALTER TABLE [Transfer_Audit_Log] ADD CONSTRAINT [fk_log_by] FOREIGN KEY ([TransferredBy_AdminID]) REFERENCES [User_Directory] ([EmpID])"
End Sub

Public Function IMS_UserCount() As Long
    Dim row As Object
    Set row = IMS_FetchOne("SELECT COUNT(*) AS N FROM [User_Directory]")
    IMS_UserCount = CLng(row("N"))
End Function

Public Sub IMS_DeleteAllRows(ByVal conn As Object)
    IMS_ExecOn conn, "DELETE FROM [Transfer_Audit_Log]"
    IMS_ExecOn conn, "DELETE FROM [Assets_Master]"
    IMS_ExecOn conn, "DELETE FROM [User_Directory]"
    IMS_ExecOn conn, "DELETE FROM [Allowed_Directory]"
End Sub

Public Function IMS_GetAsset(ByVal barcode As String) As Object
    Dim row As Object
    Dim person As Object
    Dim assigneeId As String
    Set row = IMS_FetchOne( _
        "SELECT * FROM [Assets_Master] WHERE [Barcode]=?", _
        Array(IMS_NormalizeBarcode(barcode)))
    If row Is Nothing Then Exit Function
    assigneeId = DictStr(row, "CurrentAssigneeID")
    If assigneeId <> "" Then
        Set person = IMS_FetchOne( _
            "SELECT [EmpName] FROM [User_Directory] WHERE [EmpID]=?", Array(assigneeId))
        If Not person Is Nothing Then row("CurrentAssigneeName") = DictStr(person, "EmpName")
    End If
    Set IMS_GetAsset = row
End Function

Public Function IMS_QueryAssets( _
    ByVal ownership As String, _
    ByVal statusName As String, _
    ByVal assetType As String, _
    ByVal searchText As String, _
    ByVal onlyEmpId As String) As Variant

    Dim sql As String
    Dim vals As Collection
    Dim rows As Collection
    Dim params() As Variant
    Dim i As Long
    Dim r As Long
    Dim d As Object
    Dim data As Variant
    Dim token As String

    If gEmpID = "" Then
        gLastError = "Sign in first."
        IMS_QueryAssets = Empty
        Exit Function
    End If
    If gRole = "User" Then onlyEmpId = gEmpID

    sql = "SELECT a.[Barcode] AS [Barcode], a.[Ownership] AS [Ownership], a.[AssetType] AS [AssetType], " & _
          "a.[Model] AS [Model], a.[PONumber] AS [PONumber], a.[Status] AS [Status], " & _
          "a.[CurrentAssigneeID] AS [CurrentAssigneeID], " & _
          "Nz(u.[EmpName], Nz(a.[CurrentAssigneeName], '')) AS [CurrentAssigneeName], " & _
          "a.[Manufacturer] AS [Manufacturer], a.[MfsSerialNo] AS [MfsSerialNo], " & _
          "a.[Make] AS [Make], a.[Capacity_Rev] AS [Capacity_Rev], a.[Description] AS [Description], " & _
          "a.[ArrivalDate] AS [ArrivalDate], a.[ExpiryDate] AS [ExpiryDate] " & _
          "FROM [Assets_Master] AS a LEFT JOIN [User_Directory] AS u " & _
          "ON a.[CurrentAssigneeID] = u.[EmpID] WHERE 1=1"
    Set vals = New Collection
    If gRole = "Lead Admin" Then
        sql = sql & " AND (UCase(Nz(a.[CurrentAssigneeID],''))=? OR UCase(Nz(u.[LeadID],''))=?)"
        vals.Add UCase$(gEmpID)
        vals.Add UCase$(gEmpID)
    End If

    If Trim$(ownership) <> "" Then
        sql = sql & " AND a.[Ownership]=?"
        vals.Add IMS_CanonicalChoice(ownership, IMS_OwnershipList())
    End If
    If Trim$(statusName) <> "" Then
        sql = sql & " AND a.[Status]=?"
        vals.Add IMS_CanonicalChoice(statusName, IMS_StatusList())
    End If
    If Trim$(assetType) <> "" Then
        sql = sql & " AND a.[AssetType]=?"
        vals.Add IMS_CanonicalChoice(assetType, IMS_AssetTypeList())
    End If
    If Trim$(onlyEmpId) <> "" Then
        sql = sql & " AND UCase(Nz(a.[CurrentAssigneeID],''))=?"
        vals.Add UCase$(Trim$(onlyEmpId))
    End If
    token = UCase$(Trim$(searchText))
    If token <> "" Then
        sql = sql & " AND (InStr(1, UCase(a.[Barcode]), ?) > 0" & _
            " OR InStr(1, UCase(Nz(a.[Model],'')), ?) > 0" & _
            " OR InStr(1, UCase(Nz(u.[EmpName], Nz(a.[CurrentAssigneeName],''))), ?) > 0" & _
            " OR InStr(1, UCase(Nz(a.[PONumber],'')), ?) > 0" & _
            " OR InStr(1, UCase(Nz(a.[CurrentAssigneeID],'')), ?) > 0" & _
            " OR InStr(1, UCase(Nz(a.[Manufacturer],'')), ?) > 0)"
        For i = 1 To 6
            vals.Add token
        Next i
    End If
    sql = sql & " ORDER BY a.[Barcode]"

    If vals.Count = 0 Then
        Set rows = IMS_FetchDicts(sql)
    Else
        ReDim params(0 To vals.Count - 1)
        For i = 1 To vals.Count
            params(i - 1) = vals(i)
        Next i
        Set rows = IMS_FetchDicts(sql, params)
    End If

    If rows.Count = 0 Then
        IMS_QueryAssets = Empty
        Exit Function
    End If

    ReDim data(1 To rows.Count, 1 To 15)
    For r = 1 To rows.Count
        Set d = rows(r)
        data(r, 1) = DictStr(d, "Barcode")
        data(r, 2) = DictStr(d, "Ownership")
        data(r, 3) = DictStr(d, "AssetType")
        data(r, 4) = DictStr(d, "Model")
        data(r, 5) = DictStr(d, "PONumber")
        data(r, 6) = DictStr(d, "Status") & IMS_ExpiryNote(DictStr(d, "ExpiryDate"))
        data(r, 7) = DictStr(d, "CurrentAssigneeID")
        data(r, 8) = DictStr(d, "CurrentAssigneeName")
        data(r, 9) = DictStr(d, "Manufacturer")
        data(r, 10) = DictStr(d, "MfsSerialNo")
        data(r, 11) = DictStr(d, "Make")
        data(r, 12) = DictStr(d, "Capacity_Rev")
        data(r, 13) = DictStr(d, "Description")
        data(r, 14) = DictStr(d, "ArrivalDate")
        data(r, 15) = DictStr(d, "ExpiryDate")
    Next r
    IMS_QueryAssets = data
End Function

Public Function IMS_QueryAudit() As Variant
    Dim rows As Collection
    Dim data As Variant
    Dim r As Long
    Dim d As Object
    If gEmpID = "" Then
        gLastError = "Sign in first."
        IMS_QueryAudit = Empty
        Exit Function
    End If
    If gRole = "User" Then
        gLastError = "Audit export is limited to inventory custodians."
        IMS_QueryAudit = Empty
        Exit Function
    End If
    Set rows = IMS_FetchDicts( _
        "SELECT [LogID], [Timestamp], [AssetBarcode], [PreviousAssigneeID], [NewAssigneeID], " & _
        "[TransferredBy_AdminID], [Remarks] FROM [Transfer_Audit_Log] ORDER BY [LogID]")
    If rows.Count = 0 Then
        IMS_QueryAudit = Empty
        Exit Function
    End If
    ReDim data(1 To rows.Count, 1 To 7)
    For r = 1 To rows.Count
        Set d = rows(r)
        data(r, 1) = DictStr(d, "LogID")
        data(r, 2) = DictStr(d, "Timestamp")
        data(r, 3) = DictStr(d, "AssetBarcode")
        data(r, 4) = DictStr(d, "PreviousAssigneeID")
        data(r, 5) = DictStr(d, "NewAssigneeID")
        data(r, 6) = DictStr(d, "TransferredBy_AdminID")
        data(r, 7) = DictStr(d, "Remarks")
    Next r
    IMS_QueryAudit = data
End Function

Public Function IMS_SaveNewAsset(ByVal fields As Object) As Boolean
    Dim conn As Object
    Dim barcode As String
    Dim assigneeId As String
    Dim assigneeName As String
    Dim statusName As String
    Dim attempt As Long
    Dim lastErr As Long
    Dim lastDesc As String

    gLastError = ""
    If Not IMS_IsCustodianRole(gRole) Then
        gLastError = "Only inventory custodians can inward hardware."
        Exit Function
    End If
    If Not ValidAssetFields(fields, True) Then Exit Function

    barcode = IMS_NormalizeBarcode(DictStr(fields, "Barcode"))
    If Not IMS_GetAsset(barcode) Is Nothing Then
        gLastError = "Barcode " & barcode & " is already in the lab database."
        Exit Function
    End If
    If SerialInUse(DictStr(fields, "MfsSerialNo"), barcode) Then
        gLastError = "Manufacturer serial number " & Trim$(DictStr(fields, "MfsSerialNo")) & " is already recorded on another asset."
        Exit Function
    End If

    assigneeId = UCase$(Trim$(DictStr(fields, "AssigneeID")))
    If assigneeId = "" Then
        statusName = "Available in Lab"
        assigneeName = ""
    Else
        assigneeName = LookupActiveName(assigneeId)
        If assigneeName = "" Then
            gLastError = "Initial assignee " & assigneeId & " is not an active employee."
            Exit Function
        End If
        statusName = "Assigned to Engineer"
    End If

    For attempt = 1 To 4
        Set conn = Nothing
        On Error Resume Next
        Set conn = IMS_Connect()
        If Err.Number = 0 Then conn.BeginTrans
        If Err.Number = 0 Then
            IMS_ExecOn conn, InsertAssetSql(fields, barcode, assigneeId, assigneeName, statusName)
        End If
        If Err.Number = 0 Then
            IMS_ExecOn conn, _
                "INSERT INTO [Transfer_Audit_Log] ([Timestamp], [AssetBarcode], [PreviousAssigneeID], [NewAssigneeID], [TransferredBy_AdminID], [Remarks]) " & _
                "VALUES (Now(), " & SqlText(barcode) & ", NULL, " & SqlIdOrNull(assigneeId) & ", " & _
                SqlIdOrNull(gEmpID) & ", " & SqlText("Inward receipt") & ")"
        End If
        If Err.Number = 0 Then conn.CommitTrans
        lastErr = Err.Number
        lastDesc = Err.Description
        If lastErr <> 0 Then
            Err.Clear
            If Not conn Is Nothing Then
                If conn.State = 1 Then conn.RollbackTrans
            End If
        End If
        IMS_Disconnect conn
        On Error GoTo 0
        If lastErr = 0 Then
            IMS_SaveNewAsset = True
            Exit Function
        End If
        If Not IsLockError(lastErr) Or attempt = 4 Then
            gLastError = "Inward did not save. " & lastDesc
            Exit Function
        End If
        PauseBrief 0.15 * attempt
    Next attempt
End Function

Public Function IMS_UpdateAssetMetadata(ByVal fields As Object) As Boolean
    Dim conn As Object
    Dim asset As Object
    Dim barcode As String
    Dim arrivalSql As String
    Dim expirySql As String
    Dim revision As Long
    Dim saved As Long

    gLastError = ""
    If Not IMS_IsCustodianRole(gRole) Then
        gLastError = "Only inventory custodians can edit hardware metadata."
        Exit Function
    End If
    If Not ValidAssetFields(fields, True) Then Exit Function
    barcode = IMS_NormalizeBarcode(DictStr(fields, "Barcode"))
    Set asset = IMS_GetAsset(barcode)
    If asset Is Nothing Then
        gLastError = "Barcode " & barcode & " was not found."
        Exit Function
    End If
    If SerialInUse(DictStr(fields, "MfsSerialNo"), barcode) Then
        gLastError = "Manufacturer serial number " & Trim$(DictStr(fields, "MfsSerialNo")) & " is already recorded on another asset."
        Exit Function
    End If

    On Error GoTo Fail
    revision = RevisionOf(asset)
    arrivalSql = IMS_SqlDateOrNull(DictStr(fields, "ArrivalDate"))
    expirySql = IMS_SqlDateOrNull(DictStr(fields, "ExpiryDate"))
    Set conn = IMS_Connect()
    saved = IMS_ExecCount(conn, _
        "UPDATE [Assets_Master] SET [Ownership]=" & SqlText(DictStr(fields, "Ownership")) & _
        ", [AssetType]=" & SqlText(DictStr(fields, "AssetType")) & _
        ", [Manufacturer]=" & SqlText(DictStr(fields, "Manufacturer")) & _
        ", [MfsSerialNo]=" & SqlTextOrNull(DictStr(fields, "MfsSerialNo")) & _
        ", [Make]=" & SqlText(DictStr(fields, "Make")) & _
        ", [Model]=" & SqlText(DictStr(fields, "Model")) & _
        ", [Capacity_Rev]=" & SqlText(DictStr(fields, "Capacity_Rev")) & _
        ", [Description]=" & SqlText(DictStr(fields, "Description")) & _
        ", [PONumber]=" & SqlText(DictStr(fields, "PONumber")) & _
        ", [ArrivalDate]=" & arrivalSql & ", [ExpiryDate]=" & expirySql & _
        ", [Revision]=Nz([Revision],0)+1 WHERE [Barcode]=" & SqlText(barcode) & _
        " AND Nz([Revision],0)=" & CStr(revision))
    IMS_Disconnect conn
    If saved = 0 Then
        gLastError = "Someone else saved this asset. Scan it again and retry."
        Exit Function
    End If
    IMS_UpdateAssetMetadata = True
    Exit Function
Fail:
    gLastError = "Metadata update failed. " & Err.Description
    IMS_Disconnect conn
End Function

Public Function IMS_TransferAsset( _
    ByVal barcode As String, _
    ByVal newEmpId As String, _
    ByVal remarks As String, _
    ByVal newStatus As String) As Boolean

    Dim conn As Object
    Dim asset As Object
    Dim previousId As String
    Dim currentStatus As String
    Dim storedId As String
    Dim storedName As String
    Dim statusName As String
    Dim revision As Long
    Dim saved As Long
    Dim attempt As Long
    Dim lastErr As Long
    Dim lastDesc As String

    gLastError = ""
    If Not IMS_IsCustodianRole(gRole) Then
        gLastError = "Only inventory custodians can transfer hardware."
        Exit Function
    End If

    barcode = IMS_NormalizeBarcode(barcode)
    newEmpId = UCase$(Trim$(newEmpId))
    remarks = Trim$(remarks)
    statusName = IMS_CanonicalChoice(newStatus, IMS_StatusList())
    If statusName = "" Then statusName = "Assigned to Engineer"

    If Not IMS_ValidBarcode(barcode) Then
        gLastError = "Scan a barcode of 4 to 64 letters, digits, hyphens, or underscores."
        Exit Function
    End If
    If statusName = "Scrapped/Returned" And Len(remarks) = 0 Then
        gLastError = "Enter remarks when scrapping or returning an asset."
        Exit Function
    End If

    Set asset = IMS_GetAsset(barcode)
    If asset Is Nothing Then
        gLastError = "Barcode " & barcode & " is not in the lab database."
        Exit Function
    End If
    previousId = UCase$(DictStr(asset, "CurrentAssigneeID"))
    currentStatus = DictStr(asset, "Status")
    revision = RevisionOf(asset)
    If Not IMS_StatusMoveAllowed(currentStatus, statusName, gRole) Then
        gLastError = "Status cannot move from " & currentStatus & " to " & statusName & "."
        Exit Function
    End If

    Select Case statusName
        Case "Assigned to Engineer"
            If newEmpId = "" Then
                gLastError = "Enter the target employee ID."
                Exit Function
            End If
            storedName = LookupActiveName(newEmpId)
            If storedName = "" Then
                gLastError = "Employee " & newEmpId & " is not an active user."
                Exit Function
            End If
            storedId = newEmpId
        Case "Available in Lab", "Scrapped/Returned"
            If newEmpId <> "" Then
                gLastError = "Clear the target employee ID when returning a board to the cage or scrapping it."
                Exit Function
            End If
            storedId = ""
            storedName = ""
        Case "Under Maintenance"
            If newEmpId = "" Then
                storedId = previousId
                storedName = DictStr(asset, "CurrentAssigneeName")
            Else
                storedName = LookupActiveName(newEmpId)
                If storedName = "" Then
                    gLastError = "Employee " & newEmpId & " is not an active user."
                    Exit Function
                End If
                storedId = newEmpId
            End If
    End Select

    If statusName = currentStatus And storedId = previousId Then
        gLastError = "This transfer does not change the assignee or the status."
        Exit Function
    End If

    For attempt = 1 To 4
        Set conn = Nothing
        saved = 0
        On Error Resume Next
        Set conn = IMS_Connect()
        If Err.Number = 0 Then conn.BeginTrans
        If Err.Number = 0 Then
            saved = IMS_ExecCount(conn, _
                "UPDATE [Assets_Master] SET [CurrentAssigneeID]=" & SqlIdOrNull(storedId) & _
                ", [CurrentAssigneeName]=" & SqlText(storedName) & _
                ", [Status]=" & SqlText(statusName) & _
                ", [Revision]=Nz([Revision],0)+1 WHERE [Barcode]=" & SqlText(barcode) & _
                " AND Nz([Revision],0)=" & CStr(revision))
            If saved = 0 And Err.Number = 0 Then
                Err.Raise vbObjectError + 2048, "IMS", "Someone else saved this asset. Scan it again and retry."
            End If
        End If
        If Err.Number = 0 Then
            IMS_ExecOn conn, _
                "INSERT INTO [Transfer_Audit_Log] ([Timestamp], [AssetBarcode], [PreviousAssigneeID], [NewAssigneeID], [TransferredBy_AdminID], [Remarks]) " & _
                "VALUES (Now(), " & SqlText(barcode) & ", " & SqlIdOrNull(previousId) & ", " & _
                SqlIdOrNull(storedId) & ", " & SqlIdOrNull(gEmpID) & ", " & SqlText(remarks) & ")"
        End If
        If Err.Number = 0 Then conn.CommitTrans
        lastErr = Err.Number
        lastDesc = Err.Description
        If lastErr <> 0 Then
            Err.Clear
            If Not conn Is Nothing Then
                If conn.State = 1 Then conn.RollbackTrans
            End If
        End If
        IMS_Disconnect conn
        On Error GoTo 0
        If lastErr = 0 Then
            IMS_TransferAsset = True
            Exit Function
        End If
        If Not IsLockError(lastErr) Or attempt = 4 Then
            gLastError = "Transfer did not save. " & lastDesc
            Exit Function
        End If
        PauseBrief 0.15 * attempt
    Next attempt
End Function

Private Function InsertAssetSql( _
    ByVal fields As Object, _
    ByVal barcode As String, _
    ByVal assigneeId As String, _
    ByVal assigneeName As String, _
    ByVal statusName As String) As String

    ' The caller executes the returned SQL. Dates are validated literals, never raw user text.
    InsertAssetSql = _
        "INSERT INTO [Assets_Master] (" & _
        "[Barcode], [Ownership], [AssetType], [Manufacturer], [MfsSerialNo], [Make], [Model], " & _
        "[Capacity_Rev], [Description], [PONumber], [ArrivalDate], [ExpiryDate], " & _
        "[CurrentAssigneeID], [CurrentAssigneeName], [Status], [CreatedBy], [CreatedAt], [Revision]) VALUES (" & _
        SqlText(barcode) & ", " & SqlText(DictStr(fields, "Ownership")) & ", " & _
        SqlText(DictStr(fields, "AssetType")) & ", " & SqlText(DictStr(fields, "Manufacturer")) & ", " & _
        SqlTextOrNull(DictStr(fields, "MfsSerialNo")) & ", " & SqlText(DictStr(fields, "Make")) & ", " & _
        SqlText(DictStr(fields, "Model")) & ", " & SqlText(DictStr(fields, "Capacity_Rev")) & ", " & _
        SqlText(DictStr(fields, "Description")) & ", " & SqlText(DictStr(fields, "PONumber")) & ", " & _
        IMS_SqlDateOrNull(DictStr(fields, "ArrivalDate")) & ", " & _
        IMS_SqlDateOrNull(DictStr(fields, "ExpiryDate")) & ", " & _
        SqlIdOrNull(assigneeId) & ", " & SqlText(assigneeName) & ", " & SqlText(statusName) & ", " & _
        SqlText(gEmpID) & ", Now(), 0)"
End Function

Private Function ValidAssetFields(ByVal fields As Object, ByVal requireBarcode As Boolean) As Boolean
    Dim barcode As String
    Dim ownership As String
    Dim assetType As String
    barcode = IMS_NormalizeBarcode(DictStr(fields, "Barcode"))
    ownership = IMS_CanonicalChoice(DictStr(fields, "Ownership"), IMS_OwnershipList())
    assetType = IMS_CanonicalChoice(DictStr(fields, "AssetType"), IMS_AssetTypeList())
    If requireBarcode And Not IMS_ValidBarcode(barcode) Then
        gLastError = "Barcode must be 4 to 64 letters, digits, hyphens, or underscores."
        Exit Function
    End If
    If ownership = "" Then
        gLastError = "Choose ownership: AMD-Specific or Wipro-Specific."
        Exit Function
    End If
    If assetType = "" Then
        gLastError = "Choose an asset type."
        Exit Function
    End If
    If Trim$(DictStr(fields, "PONumber")) = "" Then
        gLastError = "Purchase order number is required."
        Exit Function
    End If
    If Trim$(DictStr(fields, "ArrivalDate")) <> "" And Not IMS_IsIsoDate(DictStr(fields, "ArrivalDate")) Then
        gLastError = "Arrival date must be YYYY-MM-DD."
        Exit Function
    End If
    If Trim$(DictStr(fields, "ExpiryDate")) <> "" And Not IMS_IsIsoDate(DictStr(fields, "ExpiryDate")) Then
        gLastError = "Expiry date must be YYYY-MM-DD."
        Exit Function
    End If
    fields("Ownership") = ownership
    fields("AssetType") = assetType
    fields("Barcode") = barcode
    ValidAssetFields = True
End Function

Private Function SerialInUse(ByVal serialNo As String, ByVal barcode As String) As Boolean
    Dim row As Object
    serialNo = UCase$(Trim$(serialNo))
    If serialNo = "" Then Exit Function
    Set row = IMS_FetchOne( _
        "SELECT [Barcode] FROM [Assets_Master] WHERE UCase(Nz([MfsSerialNo],''))=? AND [Barcode]<>?", _
        Array(serialNo, barcode))
    SerialInUse = Not row Is Nothing
End Function

Private Function RevisionOf(ByVal asset As Object) As Long
    Dim raw As String
    raw = DictStr(asset, "Revision")
    If raw <> "" And IsNumeric(raw) Then RevisionOf = CLng(raw)
End Function

Public Function IMS_ExecCount(ByVal conn As Object, ByVal sql As String) As Long
    Dim affected As Variant
    affected = 0
    conn.Execute sql, affected
    If IsNull(affected) Or IsEmpty(affected) Then Exit Function
    ' Zero means the revision did not match. A negative value means the provider
    ' did not report a count, so the save is kept.
    IMS_ExecCount = CLng(affected)
End Function

Private Function SqlIdOrNull(ByVal id As String) As String
    id = UCase$(Trim$(id))
    If id = "" Then
        SqlIdOrNull = "NULL"
    Else
        SqlIdOrNull = SqlText(id)
    End If
End Function

Private Function SqlTextOrNull(ByVal value As String) As String
    If Trim$(value) = "" Then
        SqlTextOrNull = "NULL"
    Else
        SqlTextOrNull = SqlText(Trim$(value))
    End If
End Function

Private Function LookupActiveName(ByVal empId As String) As String
    Dim row As Object
    Set row = IMS_FetchOne( _
        "SELECT [EmpName], [Active] FROM [User_Directory] WHERE [EmpID]=?", _
        Array(UCase$(Trim$(empId))))
    If row Is Nothing Then Exit Function
    If Not IsTruthy(row("Active")) Then Exit Function
    LookupActiveName = DictStr(row, "EmpName")
End Function

Private Function ConnString() As String
    ' OLE DB Services=-2 keeps session pooling off so Close releases the Access lock file.
    ' Mode=16 is Share Deny None. Database Locking Mode=1 is row-level locking.
    ConnString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & gBackendPath & ";" & _
        "Persist Security Info=False;" & _
        "Mode=16;" & _
        "OLE DB Services=-2;" & _
        "Jet OLEDB:Database Locking Mode=1;" & _
        "Jet OLEDB:Engine Type=6;"
End Function

Private Sub BindParams(ByVal cmd As Object, Optional ByVal params As Variant)
    Dim i As Long
    Dim s As String
    Dim size As Long
    Dim typ As Long
    Dim p As Object
    If IsMissing(params) Then Exit Sub
    If Not IsArray(params) Then Exit Sub
    For i = LBound(params) To UBound(params)
        s = CStr(params(i))
        If Len(s) > 255 Then
            typ = 203 ' adLongVarWChar
            size = Len(s)
        Else
            typ = 202 ' adVarWChar
            size = 255
        End If
        If size < 1 Then size = 1
        Set p = cmd.CreateParameter("p" & CStr(i), typ, 1, size, s)
        cmd.Parameters.Append p
    Next i
End Sub

Private Function RsToDict(ByVal rs As Object) As Object
    Dim d As Object
    Dim f As Object
    Dim dv As Date
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' vbTextCompare
    For Each f In rs.Fields
        If IsNull(f.Value) Then
            d(f.Name) = ""
        ElseIf f.Type = 7 Or f.Type = 133 Or f.Type = 134 Or f.Type = 135 Then
            dv = CDate(f.Value)
            If TimeValue(dv) = 0 Then
                d(f.Name) = Format$(dv, "yyyy-mm-dd")
            Else
                d(f.Name) = Format$(dv, "yyyy-mm-dd") & "T" & Format$(dv, "hh:nn:ss")
            End If
        Else
            d(f.Name) = f.Value
        End If
    Next f
    Set RsToDict = d
End Function

Private Function DictStr(ByVal d As Object, ByVal key As String) As String
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsNull(d(key)) Then Exit Function
    DictStr = Trim$(CStr(d(key)))
End Function

Private Function IsTruthy(ByVal v As Variant) As Boolean
    If IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        IsTruthy = v
    ElseIf IsNumeric(v) Then
        IsTruthy = (CLng(v) <> 0)
    Else
        IsTruthy = (UCase$(CStr(v)) = "TRUE" Or UCase$(CStr(v)) = "YES" Or CStr(v) = "-1")
    End If
End Function

Private Function TableExists(ByVal conn As Object, ByVal tableName As String) As Boolean
    Dim rs As Object
    On Error Resume Next
    Set rs = conn.Execute("SELECT TOP 1 * FROM [" & tableName & "]")
    If Err.Number = 0 Then
        TableExists = True
        If Not rs Is Nothing Then
            If rs.State <> 0 Then rs.Close
        End If
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub EnsureIndex(ByVal conn As Object, ByVal ddl As String)
    On Error Resume Next
    IMS_ExecOn conn, ddl
    Err.Clear
    On Error GoTo 0
End Sub

Private Function IsLockError(ByVal n As Long) As Boolean
    Select Case n
        Case 3218, 3260, 3050, 3188, -2147467259, -2147217865, -2147217871, -2147217887
            IsLockError = True
    End Select
End Function

Private Function ProviderMessage(ByVal raw As String) As String
    If InStr(1, raw, "provider", vbTextCompare) > 0 Or InStr(1, raw, "ACE", vbTextCompare) > 0 Then
        ProviderMessage = "Excel could not start Microsoft.ACE.OLEDB.12.0. That provider is part of Microsoft Office, not a separate product. " & raw
    Else
        ProviderMessage = "Could not open the lab database. " & raw
    End If
End Function

Private Sub PauseBrief(ByVal seconds As Single)
    Dim until As Single
    until = Timer + seconds
    Do While Timer < until
        DoEvents
    Loop
End Sub

Private Sub EnsureParentFolder(ByVal filePath As String)
    Dim fso As Object
    Dim folder As String
    Dim parts() As String
    Dim cur As String
    Dim startAt As Long
    Dim i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetParentFolderName(filePath)
    If Len(folder) = 0 Then Exit Sub
    If fso.FolderExists(folder) Then Exit Sub
    If Left$(folder, 2) = "\\" Then
        parts = Split(Mid$(folder, 3), "\")
        If UBound(parts) < 1 Then Exit Sub
        cur = "\\" & parts(0) & "\" & parts(1)
        startAt = 2
    Else
        parts = Split(folder, "\")
        cur = parts(0)
        startAt = 1
    End If
    For i = startAt To UBound(parts)
        cur = cur & "\" & parts(i)
        If Not fso.FolderExists(cur) Then fso.CreateFolder cur
    Next i
End Sub

' Quotes a value that has already been constrained to lab field rules.
' Asset text still goes through this only after ValidAssetFields. Apostrophes are doubled.
Private Function SqlText(ByVal value As String) As String
    Dim s As String
    s = Replace(value, vbNullChar, "")
    s = Replace(s, "'", "''")
    SqlText = "'" & s & "'"
End Function
