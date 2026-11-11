Attribute VB_Name = "Module_AccessControl"
Option Explicit

' Role gates for the lab workbook.
' Self-registration can create a User only. Admin and Super Admin are assigned from frmUserAdmin.
' These checks stop the forms from showing the wrong actions. They are not a substitute for
' Windows permissions on the shared folder: anyone who can open the .accdb can edit tables.

Public Function IMS_TryLogin(ByVal empId As String, ByVal accessCode As String) As Boolean
    Dim row As Object
    gEmpID = ""
    gEmpName = ""
    gRole = ""
    gLastError = ""
    empId = UCase$(Trim$(empId))
    accessCode = Trim$(accessCode)

    If empId = "" Or accessCode = "" Then
        gLastError = "Enter your alias ID and lab access code."
        Exit Function
    End If

    On Error GoTo Fail
    Dim conn As Object
    Set conn = IMS_Connect()
    IMS_EnsureSchema conn
    IMS_Disconnect conn
    Set conn = Nothing
    Set row = IMS_FetchOne( _
        "SELECT [EmpID], [EmpName], [Role], [AccessCode], [Active], [Approval] FROM [User_Directory] WHERE [EmpID]=?", _
        Array(empId))
    If row Is Nothing Then
        gLastError = "Alias ID or access code is not recognized."
        Exit Function
    End If
    Dim approval As String
    approval = UCase$(Trim$(CStr(row("Approval") & "")))
    If approval = "" Then approval = "APPROVED"
    If approval = "PENDING" Then
        gLastError = "This alias is waiting for Super Admin approval."
        Exit Function
    End If
    If approval = "REJECTED" Then
        gLastError = "This request was not approved. Ask a Super Admin."
        Exit Function
    End If
    If Not FlagOn(row("Active")) Then
        gLastError = "This alias is inactive. Ask a Super Admin to turn it back on."
        Exit Function
    End If
    If Not IMS_AccessCodeMatches(accessCode, CStr(row("AccessCode"))) Then
        gLastError = "Alias ID or access code is not recognized."
        Exit Function
    End If
    If Not IMS_IsListed(CStr(row("Role")), IMS_RoleList()) Then
        gLastError = "This account has an unknown role. Ask a Super Admin to correct it."
        Exit Function
    End If

    gEmpID = CStr(row("EmpID"))
    gEmpName = CStr(row("EmpName"))
    gRole = IMS_CanonicalChoice(CStr(row("Role")), IMS_RoleList())
    IMS_TryLogin = True
    If Left$(CStr(row("AccessCode")), 3) <> "v1:" Then
        On Error Resume Next
        Set conn = IMS_Connect()
        IMS_ExecOn conn, _
            "UPDATE [User_Directory] SET [AccessCode]=? WHERE [EmpID]=?", _
            Array(IMS_HashAccessCode(accessCode), empId)
        IMS_Disconnect conn
        On Error GoTo 0
    End If
    Exit Function
Fail:
    gEmpID = ""
    gEmpName = ""
    gRole = ""
    gLastError = "Sign-in could not reach the lab database. " & Err.Description
End Function

Public Function IMS_RegisterUser( _
    ByVal empId As String, _
    ByVal empName As String, _
    ByVal email As String, _
    ByVal requestedRole As String, _
    ByVal accessCode As String, _
    ByVal confirmCode As String) As Boolean

    Dim existing As Object
    gLastError = ""
    empId = UCase$(Trim$(empId))
    empName = Trim$(empName)
    email = LCase$(Trim$(email))
    requestedRole = IMS_CanonicalChoice(requestedRole, IMS_RoleList())
    If requestedRole = "" Then requestedRole = "User"
    accessCode = Trim$(accessCode)

    If Len(empId) < 2 Or Len(empId) > 32 Then
        gLastError = "Alias ID must be 2 to 32 characters."
        Exit Function
    End If
    If Not IsSafeId(empId) Then
        gLastError = "Alias ID may contain letters, digits, and hyphens."
        Exit Function
    End If
    If Len(empName) = 0 Or Len(empName) > 120 Then
        gLastError = "Enter the full name."
        Exit Function
    End If
    If Not IsMailId(email) Then
        gLastError = "Enter the mail ID in the form name@company.com."
        Exit Function
    End If
    If Len(accessCode) < 4 Or Len(accessCode) > 32 Then
        gLastError = "The lab access code must be 4 to 32 characters."
        Exit Function
    End If
    If accessCode <> Trim$(confirmCode) Then
        gLastError = "The access code and the confirmation do not match."
        Exit Function
    End If

    On Error GoTo Fail
    Dim schemaConn As Object
    Set schemaConn = IMS_Connect()
    IMS_EnsureSchema schemaConn
    IMS_Disconnect schemaConn
    Dim allowed As Object
    Set allowed = IMS_FetchOne( _
        "SELECT [Alias], [Email], [FullName] FROM [Allowed_Directory] WHERE [Alias]=?", Array(empId))
    If allowed Is Nothing Then
        gLastError = "That alias is not on the lab directory. A Super Admin adds a person before they can request access."
        Exit Function
    End If
    If StrComp(LCase$(Trim$(CStr(allowed("Email")))), email, vbBinaryCompare) <> 0 Then
        gLastError = "The mail ID does not match the lab directory for this alias."
        Exit Function
    End If
    If Len(Trim$(CStr(allowed("FullName")))) > 0 Then
        If StrComp(Trim$(CStr(allowed("FullName"))), empName, vbTextCompare) <> 0 Then
            gLastError = "The full name does not match the lab directory for this alias."
            Exit Function
        End If
    End If
    Set existing = IMS_FetchOne( _
        "SELECT [EmpID], [Approval] FROM [User_Directory] WHERE [EmpID]=?", Array(empId))
    If Not existing Is Nothing Then
        If UCase$(Trim$(CStr(existing("Approval") & ""))) = "REJECTED" Then
            gLastError = "This alias was not approved. Ask a Super Admin to review it again."
        Else
            gLastError = "That alias is already registered. Sign in, or wait for Super Admin approval."
        End If
        Exit Function
    End If

    Dim conn As Object
    Set conn = IMS_Connect()
    IMS_ExecOn conn, _
        "INSERT INTO [User_Directory] ([EmpID], [EmpName], [Role], [AccessCode], [Active], [Email], [RequestedRole], [Approval], [WindowsUser]) " & _
        "VALUES (?, ?, 'User', ?, False, ?, ?, 'Pending', ?)", _
        Array(empId, empName, IMS_HashAccessCode(accessCode), email, requestedRole, Environ$("USERNAME"))
    IMS_Disconnect conn
    IMS_RegisterUser = True
    Exit Function
Fail:
    gLastError = "Registration did not save. " & Err.Description
End Function

Public Function IMS_ListUsers() As Variant
    Dim rows As Collection
    Dim data As Variant
    Dim r As Long
    Dim d As Object
    gLastError = ""
    If Not IMS_IsSuperRole(gRole) Then
        gLastError = "Only a Super Admin can manage accounts."
        IMS_ListUsers = Empty
        Exit Function
    End If
    On Error GoTo Fail
    Set rows = IMS_FetchDicts( _
        "SELECT [EmpID], [EmpName], [Role], [Active], [Approval], [RequestedRole], [LeadID], [Email] " & _
        "FROM [User_Directory] ORDER BY IIf([Approval]='Pending', 0, 1), [EmpID]")
    If rows.Count = 0 Then
        IMS_ListUsers = Empty
        Exit Function
    End If
    ReDim data(1 To rows.Count, 1 To 6)
    For r = 1 To rows.Count
        Set d = rows(r)
        data(r, 1) = CStr(d("EmpID"))
        data(r, 2) = CStr(d("EmpName"))
        data(r, 3) = CStr(d("Role"))
        data(r, 4) = AccountStatus(d)
        data(r, 5) = CStr(d("LeadID") & "")
        data(r, 6) = CStr(d("RequestedRole") & "")
    Next r
    IMS_ListUsers = data
    Exit Function
Fail:
    gLastError = "Could not read the user directory. " & Err.Description
    IMS_ListUsers = Empty
End Function

Public Function IMS_SetUserRole( _
    ByVal empId As String, _
    ByVal newRole As String, _
    ByVal makeActive As Boolean, _
    Optional ByVal leadId As String = "") As Boolean

    Dim target As Object
    Dim countRow As Object
    Dim superCount As Long
    Dim conn As Object

    gLastError = ""
    If Not IMS_IsSuperRole(gRole) Then
        gLastError = "Only a Super Admin can change roles."
        Exit Function
    End If
    empId = UCase$(Trim$(empId))
    newRole = IMS_CanonicalChoice(newRole, IMS_RoleList())
    If newRole = "" Then
        gLastError = "Choose Super Admin, Admin, Lead Admin, or User."
        Exit Function
    End If
    leadId = UCase$(Trim$(leadId))
    If newRole <> "User" Then leadId = ""
    If StrComp(empId, gEmpID, vbTextCompare) = 0 And Not makeActive Then
        gLastError = "You cannot turn off your own account."
        Exit Function
    End If

    On Error GoTo Fail
    Set target = IMS_FetchOne( _
        "SELECT [EmpID], [Role], [Active] FROM [User_Directory] WHERE [EmpID]=?", Array(empId))
    If target Is Nothing Then
        gLastError = "Employee " & empId & " was not found."
        Exit Function
    End If

    Set countRow = IMS_FetchOne( _
        "SELECT COUNT(*) AS N FROM [User_Directory] WHERE [Role]='Super Admin' AND [Active]=True")
    superCount = CLng(countRow("N"))
    If CStr(target("Role")) = "Super Admin" And FlagOn(target("Active")) Then
        If superCount <= 1 And (newRole <> "Super Admin" Or Not makeActive) Then
            gLastError = "Keep at least one active Super Admin."
            Exit Function
        End If
    End If

    If leadId <> "" Then
        Dim lead As Object
        Set lead = IMS_FetchOne( _
            "SELECT [Role], [Active], [Approval] FROM [User_Directory] WHERE [EmpID]=?", Array(leadId))
        If lead Is Nothing Or CStr(lead("Role")) <> "Lead Admin" Or Not FlagOn(lead("Active")) Then
            gLastError = "Choose an active Lead Admin for this person's team."
            Exit Function
        End If
    End If

    Set conn = IMS_Connect()
    If makeActive Then
        IMS_ExecOn conn, _
            "UPDATE [User_Directory] SET [Role]=?, [Active]=True, [Approval]='Approved', [LeadID]=? WHERE [EmpID]=?", _
            Array(newRole, leadId, empId)
    Else
        IMS_ExecOn conn, _
            "UPDATE [User_Directory] SET [Role]=?, [Active]=False, [LeadID]=? WHERE [EmpID]=?", _
            Array(newRole, leadId, empId)
    End If
    IMS_Disconnect conn
    IMS_SetUserRole = True
    Exit Function
Fail:
    gLastError = "The role change did not save. " & Err.Description
End Function

Public Function IMS_RejectUser(ByVal empId As String) As Boolean
    gLastError = ""
    If Not IMS_IsSuperRole(gRole) Then
        gLastError = "Only a Super Admin can reject a request."
        Exit Function
    End If
    empId = UCase$(Trim$(empId))
    If StrComp(empId, gEmpID, vbTextCompare) = 0 Then
        gLastError = "You cannot reject your own account."
        Exit Function
    End If
    On Error GoTo Fail
    Dim conn As Object
    Set conn = IMS_Connect()
    IMS_ExecOn conn, _
        "UPDATE [User_Directory] SET [Active]=False, [Approval]='Rejected' WHERE [EmpID]=?", _
        Array(empId)
    IMS_Disconnect conn
    IMS_RejectUser = True
    Exit Function
Fail:
    gLastError = "The request was not rejected. " & Err.Description
End Function

Public Function IMS_AddAllowedPerson( _
    ByVal aliasId As String, _
    ByVal email As String, _
    ByVal fullName As String) As Boolean

    Dim existing As Object
    gLastError = ""
    If Not IMS_IsSuperRole(gRole) Then
        gLastError = "Only a Super Admin can add a person to the lab directory."
        Exit Function
    End If
    aliasId = UCase$(Trim$(aliasId))
    email = LCase$(Trim$(email))
    fullName = Trim$(fullName)
    If Not IsSafeId(aliasId) Or Len(aliasId) < 2 Then
        gLastError = "Alias ID must be 2 to 32 letters, digits, or hyphens."
        Exit Function
    End If
    If Not IsMailId(email) Then
        gLastError = "Enter the mail ID in the form name@company.com."
        Exit Function
    End If
    If Len(fullName) = 0 Then
        gLastError = "Enter the full name on the company record."
        Exit Function
    End If
    On Error GoTo Fail
    Set existing = IMS_FetchOne("SELECT [Alias] FROM [Allowed_Directory] WHERE [Alias]=?", Array(aliasId))
    Dim conn As Object
    Set conn = IMS_Connect()
    If existing Is Nothing Then
        IMS_ExecOn conn, _
            "INSERT INTO [Allowed_Directory] ([Alias], [Email], [FullName]) VALUES (?, ?, ?)", _
            Array(aliasId, email, fullName)
    Else
        IMS_ExecOn conn, _
            "UPDATE [Allowed_Directory] SET [Email]=?, [FullName]=? WHERE [Alias]=?", _
            Array(email, fullName, aliasId)
    End If
    IMS_Disconnect conn
    IMS_AddAllowedPerson = True
    Exit Function
Fail:
    gLastError = "The directory row did not save. " & Err.Description
End Function

' Requests waiting for approval. Columns: alias, name, requested role.
Public Function IMS_PendingRequests() As Variant
    Dim rows As Collection
    Dim data As Variant
    Dim r As Long
    Dim d As Object
    gLastError = ""
    IMS_PendingRequests = Empty
    If Not IMS_IsSuperRole(gRole) Then Exit Function
    On Error GoTo Fail
    Set rows = IMS_FetchDicts( _
        "SELECT [EmpID], [EmpName], [RequestedRole] FROM [User_Directory] " & _
        "WHERE [Approval]='Pending' ORDER BY [EmpID]")
    If rows.Count = 0 Then Exit Function
    ReDim data(1 To rows.Count, 1 To 3)
    For r = 1 To rows.Count
        Set d = rows(r)
        data(r, 1) = CStr(d("EmpID") & "")
        data(r, 2) = CStr(d("EmpName") & "")
        data(r, 3) = CStr(d("RequestedRole") & "")
    Next r
    IMS_PendingRequests = data
    Exit Function
Fail:
    gLastError = "Could not read access requests. " & Err.Description
End Function

' The signed-in person's own record. Keys: EmpID, EmpName, Role, Email, Lead.
Public Function IMS_GetMyAccount() As Object
    Dim row As Object
    Dim lead As Object
    Dim info As Object
    Dim leadId As String
    gLastError = ""
    If gEmpID = "" Then
        gLastError = "Sign in first."
        Exit Function
    End If
    On Error GoTo Fail
    Set row = IMS_FetchOne( _
        "SELECT [EmpID], [EmpName], [Role], [Email], [LeadID] FROM [User_Directory] WHERE [EmpID]=?", _
        Array(gEmpID))
    If row Is Nothing Then
        gLastError = "Your account was not found."
        Exit Function
    End If
    Set info = CreateObject("Scripting.Dictionary")
    info.CompareMode = 1
    info("EmpID") = CStr(row("EmpID") & "")
    info("EmpName") = CStr(row("EmpName") & "")
    info("Role") = CStr(row("Role") & "")
    info("Email") = CStr(row("Email") & "")
    leadId = UCase$(Trim$(CStr(row("LeadID") & "")))
    info("Lead") = ""
    If leadId <> "" Then
        Set lead = IMS_FetchOne("SELECT [EmpName] FROM [User_Directory] WHERE [EmpID]=?", Array(leadId))
        If lead Is Nothing Then
            info("Lead") = leadId
        Else
            info("Lead") = leadId & "  " & CStr(lead("EmpName") & "")
        End If
    End If
    Set IMS_GetMyAccount = info
    Exit Function
Fail:
    gLastError = "Could not read your account. " & Err.Description
End Function

Public Function IMS_ChangeAccessCode( _
    ByVal currentCode As String, _
    ByVal newCode As String, _
    ByVal confirmCode As String) As Boolean

    Dim row As Object
    Dim conn As Object
    gLastError = ""
    If gEmpID = "" Then
        gLastError = "Sign in first."
        Exit Function
    End If
    currentCode = Trim$(currentCode)
    newCode = Trim$(newCode)
    If currentCode = "" Then
        gLastError = "Enter your current access code."
        Exit Function
    End If
    If Len(newCode) < 4 Or Len(newCode) > 32 Then
        gLastError = "The new access code must be 4 to 32 characters."
        Exit Function
    End If
    If newCode <> Trim$(confirmCode) Then
        gLastError = "The new code and the confirmation do not match."
        Exit Function
    End If
    If newCode = currentCode Then
        gLastError = "Choose a new code that differs from the current one."
        Exit Function
    End If

    On Error GoTo Fail
    Set row = IMS_FetchOne("SELECT [AccessCode] FROM [User_Directory] WHERE [EmpID]=?", Array(gEmpID))
    If row Is Nothing Then
        gLastError = "Your account was not found."
        Exit Function
    End If
    If Not IMS_AccessCodeMatches(currentCode, CStr(row("AccessCode") & "")) Then
        gLastError = "The current access code is not correct."
        Exit Function
    End If
    Set conn = IMS_Connect()
    IMS_ExecOn conn, _
        "UPDATE [User_Directory] SET [AccessCode]=? WHERE [EmpID]=?", _
        Array(IMS_HashAccessCode(newCode), gEmpID)
    IMS_Disconnect conn
    IMS_ChangeAccessCode = True
    Exit Function
Fail:
    gLastError = "The access code did not change. " & Err.Description
    IMS_Disconnect conn
End Function

Public Sub IMS_SignOut()
    gEmpID = ""
    gEmpName = ""
    gRole = ""
End Sub

Private Function AccountStatus(ByVal d As Object) As String
    Dim approval As String
    approval = UCase$(Trim$(CStr(d("Approval") & "")))
    If approval = "PENDING" Then
        AccountStatus = "Pending"
    ElseIf approval = "REJECTED" Then
        AccountStatus = "Rejected"
    ElseIf FlagOn(d("Active")) Then
        AccountStatus = "Active"
    Else
        AccountStatus = "Inactive"
    End If
End Function

Private Function IsMailId(ByVal email As String) As Boolean
    Dim atPos As Long
    atPos = InStr(1, email, "@")
    If atPos < 2 Then Exit Function
    If InStr(atPos + 1, email, ".") < atPos + 2 Then Exit Function
    If Len(email) > 120 Then Exit Function
    IsMailId = True
End Function

Private Function FlagOn(ByVal v As Variant) As Boolean
    If IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        FlagOn = CBool(v)
    ElseIf IsNumeric(v) Then
        FlagOn = (CLng(v) <> 0)
    Else
        FlagOn = (UCase$(CStr(v)) = "TRUE" Or UCase$(CStr(v)) = "YES" Or CStr(v) = "-1")
    End If
End Function

Private Function IsSafeId(ByVal value As String) As Boolean
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "-") Then Exit Function
    Next i
    IsSafeId = True
End Function
