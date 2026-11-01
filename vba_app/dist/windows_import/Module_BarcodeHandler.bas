Attribute VB_Name = "Module_BarcodeHandler"
Option Explicit

' USB barcode scanners type the code and finish with Enter.
' Forms listen for vbKeyReturn and call these helpers. Do not set a Default button
' on the inward or transfer forms, or Enter will click that button instead of scanning.

Public Function IMS_NormalizeBarcode(ByVal raw As String) As String
    Dim s As String
    s = UCase$(Trim$(raw))
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Replace(s, " ", "")
    IMS_NormalizeBarcode = s
End Function

Public Function IMS_ValidBarcode(ByVal code As String) As Boolean
    Dim i As Long
    Dim ch As String
    code = IMS_NormalizeBarcode(code)
    If Len(code) < 4 Or Len(code) > 64 Then Exit Function
    For i = 1 To Len(code)
        ch = Mid$(code, i, 1)
        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "_") Then Exit Function
    Next i
    IMS_ValidBarcode = True
End Function

Public Function IMS_TypeFromPrefix(ByVal barcode As String) As String
    Dim token As String
    Dim dashAt As Long
    barcode = IMS_NormalizeBarcode(barcode)
    dashAt = InStr(barcode, "-")
    If dashAt > 1 Then
        token = Left$(barcode, dashAt - 1)
    Else
        token = Left$(barcode, 3)
    End If
    Select Case token
        Case "MB": IMS_TypeFromPrefix = "Motherboard"
        Case "AIC": IMS_TypeFromPrefix = "Add-in Card"
        Case "DBG": IMS_TypeFromPrefix = "Debugger"
        Case "SOC": IMS_TypeFromPrefix = "SoC Board"
        Case "PSU": IMS_TypeFromPrefix = "Power Supply"
        Case "CBL": IMS_TypeFromPrefix = "Cable"
        Case Else: IMS_TypeFromPrefix = ""
    End Select
End Function

' MB-7800X-001 -> MB-7800X so the next board in that family can reuse make and model.
Public Function IMS_FamilyPrefix(ByVal barcode As String) As String
    Dim dashAt As Long
    barcode = IMS_NormalizeBarcode(barcode)
    dashAt = InStrRev(barcode, "-")
    If dashAt > 1 Then
        IMS_FamilyPrefix = Left$(barcode, dashAt - 1)
    Else
        IMS_FamilyPrefix = barcode
    End If
End Function

Public Function IMS_SuggestFromBarcode(ByVal barcode As String) As Object
    Dim hint As Object
    Dim row As Object
    Dim family As String
    Dim assetType As String

    Set hint = CreateObject("Scripting.Dictionary")
    hint.CompareMode = 1
    barcode = IMS_NormalizeBarcode(barcode)
    assetType = IMS_TypeFromPrefix(barcode)
    If assetType <> "" Then hint("AssetType") = assetType

    family = IMS_FamilyPrefix(barcode)
    If Len(family) < 3 Then
        Set IMS_SuggestFromBarcode = hint
        Exit Function
    End If

    On Error GoTo Done
    Set row = IMS_FetchOne( _
        "SELECT TOP 1 [AssetType], [Manufacturer], [Make], [Model], [Capacity_Rev] " & _
        "FROM [Assets_Master] WHERE InStr(1, UCase([Barcode]), ?)=1 AND UCase([Barcode])<>? " & _
        "ORDER BY [CreatedAt] DESC", _
        Array(family, barcode))
    If Not row Is Nothing Then
        CopyIfPresent hint, row, "AssetType"
        CopyIfPresent hint, row, "Manufacturer"
        CopyIfPresent hint, row, "Make"
        CopyIfPresent hint, row, "Model"
        CopyIfPresent hint, row, "Capacity_Rev"
    End If
Done:
    Set IMS_SuggestFromBarcode = hint
End Function

Public Function IMS_SuggestFromModel(ByVal modelName As String) As Object
    Dim hint As Object
    Dim row As Object
    Set hint = CreateObject("Scripting.Dictionary")
    hint.CompareMode = 1
    modelName = UCase$(Trim$(modelName))
    If modelName = "" Then
        Set IMS_SuggestFromModel = hint
        Exit Function
    End If
    On Error GoTo Done
    Set row = IMS_FetchOne( _
        "SELECT TOP 1 [AssetType], [Manufacturer], [Make], [Model], [Capacity_Rev] " & _
        "FROM [Assets_Master] WHERE UCase([Model])=? ORDER BY [CreatedAt] DESC", _
        Array(modelName))
    If Not row Is Nothing Then
        CopyIfPresent hint, row, "AssetType"
        CopyIfPresent hint, row, "Manufacturer"
        CopyIfPresent hint, row, "Make"
        CopyIfPresent hint, row, "Model"
        CopyIfPresent hint, row, "Capacity_Rev"
    End If
Done:
    Set IMS_SuggestFromModel = hint
End Function

Private Sub CopyIfPresent(ByVal target As Object, ByVal source As Object, ByVal key As String)
    Dim value As String
    If source Is Nothing Then Exit Sub
    If Not source.Exists(key) Then Exit Sub
    If IsNull(source(key)) Then Exit Sub
    value = Trim$(CStr(source(key)))
    If value <> "" Then target(key) = value
End Sub
