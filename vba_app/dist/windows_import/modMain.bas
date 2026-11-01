Attribute VB_Name = "modMain"
Option Explicit

' Alt+F8 entry points:
'   IMS_Start               open the sign-in form
'   IMS_InitializeBackend   create the shared database (also in modInitDatabase)
'   IMS_EditBackendPath     point this workbook at the share
'   IMS_TestConnection      read one row and report the result

Public Sub IMS_Start()
    If Not IMS_LoadConfig() Then
        MsgBox gLastError, vbExclamation, IMS_APP_NAME
        Exit Sub
    End If
    If Not IMS_BackendFileExists() Then
        MsgBox "No database file was found at:" & vbCrLf & vbCrLf & gBackendPath & vbCrLf & vbCrLf & _
            "A Super Admin creates it once by running IMS_InitializeBackend from Alt+F8.", _
            vbExclamation, IMS_APP_NAME
        Exit Sub
    End If

    Do
        IMS_SignOut
        frmAuth.Show
        If gEmpID = "" Then Exit Do
        frmDashboard.Show
    Loop
End Sub

Public Sub IMS_EditBackendPath()
    Dim entered As String
    Dim currentPath As String
    On Error GoTo Fail
    currentPath = CStr(ThisWorkbook.Names("BackendPath").RefersToRange.Value)
    entered = Trim$(InputBox("Full path to IMS_Backend.accdb on the lab share.", "Backend path", currentPath))
    If entered = "" Then Exit Sub
    If LCase$(Right$(entered, 6)) <> ".accdb" Then
        MsgBox "The path must end with .accdb.", vbExclamation, IMS_APP_NAME
        Exit Sub
    End If
    ThisWorkbook.Names("BackendPath").RefersToRange.Value = entered
    gBackendPath = entered
    ThisWorkbook.Save
    MsgBox "Saved backend path:" & vbCrLf & entered, vbInformation, IMS_APP_NAME
    Exit Sub
Fail:
    MsgBox Err.Description, vbCritical, IMS_APP_NAME
End Sub

Public Sub IMS_TestConnection()
    Dim row As Object
    On Error GoTo Fail
    If Not IMS_LoadConfig() Then
        MsgBox gLastError, vbExclamation, IMS_APP_NAME
        Exit Sub
    End If
    Set row = IMS_FetchOne("SELECT COUNT(*) AS N FROM [User_Directory]")
    MsgBox "Connected. User accounts in the directory: " & CStr(row("N")) & vbCrLf & gBackendPath, vbInformation, IMS_APP_NAME
    Exit Sub
Fail:
    MsgBox "Connection failed. " & Err.Description, vbCritical, IMS_APP_NAME
End Sub
