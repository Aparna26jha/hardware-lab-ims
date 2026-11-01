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
        gLastError = "Enter an employee ID and a lab access code."
        Exit Function
    End If

    On Error GoTo Fail
    Set row = IMS_FetchOne( _
        "SELECT [EmpID], [EmpName], [Role], [AccessCode], [Active] FROM [User_Directory] WHERE [EmpID]=?", _
        Array(empId))
    If row Is Nothing Then
        gLastError = "Employee ID or access code is not recognized."
        Exit Function
    End If
    If Not FlagOn(row("Active")) Then
        gLastError = "This employee ID is inactive. Ask a Super Admin to turn it back on."
        Exit Function
    End If
    If StrComp(CStr(row("AccessCode")), accessCode, vbBinaryCompare) <> 0 Then
        gLastError = "Employee ID or access code is not recognized."
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
    ByVal accessCode As String, _
    ByVal confirmCode As String) As Boolean

    Dim existing As Object
    gLastError = ""
    empId = UCase$(Trim$(empId))
    empName = Trim$(empName)
    accessCode = Trim$(accessCode)

    If Len(empId) < 3 Or Len(empId) > 32 Then
        gLastError = "Employee ID must be 3 to 32 characters."
        Exit Function
    End If
    If Not IsSafeId(empId) Then
        gLastError = "Employee ID may contain letters, digits, and hyphens."
        Exit Function
    End If
    If Len(empName) = 0 Or Len(empName) > 120 Then
        gLastError = "Enter the employee name."
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
    Set existing = IMS_FetchOne("SELECT [EmpID] FROM [User_Directory] WHERE [EmpID]=?", Array(empId))
    If Not existing Is Nothing Then
        gLastError = "That employee ID is already registered. Sign in, or ask a Super Admin if the account is inactive."
        Exit Function
    End If

    Dim conn As Object
    Set conn = IMS_Connect()
    IMS_ExecOn conn, _
        "INSERT INTO [User_Directory] ([EmpID], [EmpName], [Role], [AccessCode], [Active]) VALUES (?, ?, 'User', ?, True)", _
        Array(empId, empName, accessCode)
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
        "SELECT [EmpID], [EmpName], [Role], [Active] FROM [User_Directory] ORDER BY [EmpID]")
    If rows.Count = 0 Then
        IMS_ListUsers = Empty
        Exit Function
    End If
    ReDim data(1 To rows.Count, 1 To 4)
    For r = 1 To rows.Count
        Set d = rows(r)
        data(r, 1) = CStr(d("EmpID"))
        data(r, 2) = CStr(d("EmpName"))
        data(r, 3) = CStr(d("Role"))
        If FlagOn(d("Active")) Then data(r, 4) = "Yes" Else data(r, 4) = "No"
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
    ByVal makeActive As Boolean) As Boolean

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
        gLastError = "Choose Super Admin, Admin, or User."
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

    Set conn = IMS_Connect()
    If makeActive Then
        IMS_ExecOn conn, _
            "UPDATE [User_Directory] SET [Role]=?, [Active]=True WHERE [EmpID]=?", _
            Array(newRole, empId)
    Else
        IMS_ExecOn conn, _
            "UPDATE [User_Directory] SET [Role]=?, [Active]=False WHERE [EmpID]=?", _
            Array(newRole, empId)
    End If
    IMS_Disconnect conn
    IMS_SetUserRole = True
    Exit Function
Fail:
    gLastError = "The role change did not save. " & Err.Description
End Function

Public Sub IMS_SignOut()
    gEmpID = ""
    gEmpName = ""
    gRole = ""
End Sub

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
