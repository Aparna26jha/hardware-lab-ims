Attribute VB_Name = "modGlobals"
Option Explicit

' Hardware Lab IMS
' Shared state for one Excel session. Each engineer runs a local copy of the workbook.
' The shared file on the network is the Access database, not this workbook.

Public Const IMS_APP_NAME As String = "Hardware Lab IMS"
Public Const IMS_SHOW_DEMO_FILL As Boolean = True

Public gEmpID As String
Public gEmpName As String
Public gRole As String
Public gBackendPath As String
Public gLastError As String

Public Function IMS_AssetTypeList() As Variant
    IMS_AssetTypeList = Array( _
        "Motherboard", "Add-in Card", "Debugger", "SoC Board", _
        "Power Supply", "Cable", "Other")
End Function

Public Function IMS_StatusList() As Variant
    IMS_StatusList = Array( _
        "Available in Lab", "Assigned to Engineer", _
        "Under Maintenance", "Scrapped/Returned")
End Function

Public Function IMS_RoleList() As Variant
    IMS_RoleList = Array("Super Admin", "Admin", "User")
End Function

Public Function IMS_OwnershipList() As Variant
    IMS_OwnershipList = Array("AMD-Specific", "Wipro-Specific")
End Function

Public Function IMS_IsListed(ByVal value As String, ByVal choices As Variant) As Boolean
    Dim i As Long
    For i = LBound(choices) To UBound(choices)
        If StrComp(CStr(choices(i)), Trim$(value), vbTextCompare) = 0 Then
            IMS_IsListed = True
            Exit Function
        End If
    Next i
End Function

Public Function IMS_CanonicalChoice(ByVal value As String, ByVal choices As Variant) As String
    Dim i As Long
    For i = LBound(choices) To UBound(choices)
        If StrComp(CStr(choices(i)), Trim$(value), vbTextCompare) = 0 Then
            IMS_CanonicalChoice = CStr(choices(i))
            Exit Function
        End If
    Next i
End Function

Public Function IMS_IsCustodianRole(ByVal roleName As String) As Boolean
    IMS_IsCustodianRole = (roleName = "Admin" Or roleName = "Super Admin")
End Function

Public Function IMS_IsSuperRole(ByVal roleName As String) As Boolean
    IMS_IsSuperRole = (roleName = "Super Admin")
End Function

Public Function IMS_IsIsoDate(ByVal s As String) As Boolean
    Dim d As Date
    s = Trim$(s)
    If Len(s) <> 10 Then Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Exit Function
    If Not IsNumeric(Left$(s, 4)) Then Exit Function
    If Not IsNumeric(Mid$(s, 6, 2)) Then Exit Function
    If Not IsNumeric(Right$(s, 2)) Then Exit Function
    On Error GoTo Bad
    d = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Right$(s, 2)))
    IMS_IsIsoDate = (Format$(d, "yyyy-mm-dd") = s)
    Exit Function
Bad:
    IMS_IsIsoDate = False
End Function

Public Function IMS_IsIsoStamp(ByVal s As String) As Boolean
    Dim d As Date
    s = Trim$(s)
    s = Replace(s, "T", " ")
    If Len(s) <> 19 Then Exit Function
    If Not IMS_IsIsoDate(Left$(s, 10)) Then Exit Function
    If Mid$(s, 11, 1) <> " " Then Exit Function
    On Error GoTo Bad
    d = TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), CInt(Mid$(s, 18, 2)))
    IMS_IsIsoStamp = (Format$(d, "hh:nn:ss") = Right$(s, 8))
    Exit Function
Bad:
    IMS_IsIsoStamp = False
End Function

' Returns a Jet date literal or the keyword NULL. Only a validated calendar value can pass.
Public Function IMS_SqlDateOrNull(ByVal iso As String) As String
    iso = Trim$(iso)
    If iso = "" Then
        IMS_SqlDateOrNull = "NULL"
    ElseIf Not IMS_IsIsoDate(iso) Then
        Err.Raise vbObjectError + 20, "IMS", "Use the date format YYYY-MM-DD."
    Else
        IMS_SqlDateOrNull = "#" & iso & "#"
    End If
End Function

Public Function IMS_SqlStampOrNull(ByVal iso As String) As String
    Dim stamp As String
    stamp = Replace(Trim$(iso), "T", " ")
    If stamp = "" Then
        IMS_SqlStampOrNull = "NULL"
    ElseIf Not IMS_IsIsoStamp(stamp) Then
        Err.Raise vbObjectError + 21, "IMS", "Use the timestamp format YYYY-MM-DDThh:mm:ss."
    Else
        IMS_SqlStampOrNull = "#" & stamp & "#"
    End If
End Function

Public Sub IMS_Flash(ByVal lbl As Object, ByVal message As String, ByVal isError As Boolean)
    lbl.Caption = message
    If isError Then
        lbl.ForeColor = RGB(140, 32, 32)
    Else
        lbl.ForeColor = RGB(15, 110, 78)
    End If
End Sub

Public Sub IMS_FillCombo(ByVal cbo As Object, ByVal choices As Variant, ByVal includeBlank As Boolean)
    Dim i As Long
    cbo.Clear
    If includeBlank Then cbo.AddItem ""
    For i = LBound(choices) To UBound(choices)
        cbo.AddItem CStr(choices(i))
    Next i
End Sub
