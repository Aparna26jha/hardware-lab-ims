Attribute VB_Name = "modExport"
Option Explicit

' Writes a UTF-8 CSV through the ADODB.Stream object that ships with Windows.
' The snapshot is a new file. It is not linked back to the lab database.

Public Function IMS_AssetHeaders() As Variant
    IMS_AssetHeaders = Array( _
        "Barcode", "Ownership", "Asset Type", "Model", "PO Number", "Status", _
        "Assignee ID", "Assignee Name", "Manufacturer", "Mfr Serial", "Make", _
        "Capacity / Rev", "Description", "Arrival Date", "Expiry Date")
End Function

Public Function IMS_AuditHeaders() As Variant
    IMS_AuditHeaders = Array( _
        "Log ID", "Timestamp", "Asset Barcode", "Previous Assignee ID", _
        "New Assignee ID", "Transferred By", "Remarks")
End Function

Public Sub IMS_WriteCsv(ByVal filePath As String, ByVal headers As Variant, ByVal data As Variant)
    Dim stm As Object
    Dim line As String
    Dim r As Long
    Dim c As Long
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText HeaderLine(headers), 1 ' adWriteLine

    If IsArray(data) Then
        For r = LBound(data, 1) To UBound(data, 1)
            line = ""
            For c = LBound(data, 2) To UBound(data, 2)
                If c > LBound(data, 2) Then line = line & ","
                line = line & CsvCell(CStr(data(r, c)))
            Next c
            stm.WriteText line, 1
        Next r
    End If

    stm.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stm.Close
End Sub

Public Function IMS_SuggestExportPath(ByVal prefix As String) As Variant
    IMS_SuggestExportPath = Application.GetSaveAsFilename( _
        InitialFileName:=prefix & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".csv", _
        FileFilter:="CSV (*.csv), *.csv")
End Function

Private Function HeaderLine(ByVal headers As Variant) As String
    Dim i As Long
    Dim line As String
    For i = LBound(headers) To UBound(headers)
        If i > LBound(headers) Then line = line & ","
        line = line & CsvCell(CStr(headers(i)))
    Next i
    HeaderLine = line
End Function

Private Function CsvCell(ByVal value As String) As String
    value = Replace(value, vbCrLf, " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, """", """""")
    If Len(value) > 0 Then
        Select Case Left$(value, 1)
            Case "=", "+", "-", "@", vbTab, vbCr
                value = "'" & value
        End Select
    End If
    CsvCell = """" & value & """"
End Function
