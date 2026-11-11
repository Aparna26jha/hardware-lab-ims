Attribute VB_Name = "modUi"
Option Explicit

' Every screen builds its controls in code when it opens.
' A form therefore needs no designer (.frx) file: an imported .frm, or an empty
' UserForm with the code pasted in, shows the same screen.
' Sizes are in points.

Public Const UI_FONT As String = "Segoe UI"
Public Const UI_TEAL As Long = &H565C0C          ' RGB(12, 92, 86)
Public Const UI_INK As Long = &H1C2117           ' RGB(23, 33, 28)
Public Const UI_MUTED As Long = &H646B5C         ' RGB(92, 107, 100)
Public Const UI_PANEL As Long = &HF9FBF8         ' RGB(248, 251, 249)
Public Const UI_LINE As Long = &HCCD4C5          ' RGB(197, 212, 204)
Public Const UI_CARD As Long = &HF1F3E7          ' RGB(231, 243, 241)
Public Const UI_AMD As Long = &HEAE8F8           ' RGB(248, 232, 234)
Public Const UI_WIPRO As Long = &HF8EEE7         ' RGB(231, 238, 248)

Public Sub UiSizeForm(ByVal host As Object, ByVal caption As String, ByVal widthPt As Single, ByVal heightPt As Single)
    host.Caption = caption
    host.BackColor = UI_PANEL
    host.Width = widthPt
    host.Height = heightPt
End Sub

Public Function UiAdd( _
    ByVal host As Object, ByVal progId As String, ByVal ctlName As String, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single, ByVal heightPt As Single) As Object

    Dim c As Object
    Set c = host.Controls.Add(progId, ctlName, True)
    c.Left = leftPt
    c.Top = topPt
    c.Width = widthPt
    c.Height = heightPt
    On Error Resume Next
    c.Font.Name = UI_FONT
    c.Font.Size = 9
    On Error GoTo 0
    Set UiAdd = c
End Function

Public Function UiLabel( _
    ByVal host As Object, ByVal ctlName As String, ByVal caption As String, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single, ByVal heightPt As Single) As Object

    Dim c As Object
    Set c = UiAdd(host, "Forms.Label.1", ctlName, leftPt, topPt, widthPt, heightPt)
    c.Caption = caption
    c.WordWrap = True
    c.BackStyle = 0 ' fmBackStyleTransparent
    c.ForeColor = UI_INK
    Set UiLabel = c
End Function

Public Function UiTitle(ByVal host As Object, ByVal ctlName As String, ByVal caption As String, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single) As Object

    Dim c As Object
    Set c = UiLabel(host, ctlName, caption, leftPt, topPt, widthPt, 22)
    c.Font.Size = 14
    c.Font.Bold = True
    c.ForeColor = UI_TEAL
    Set UiTitle = c
End Function

' The one-line instruction under each screen title.
Public Function UiHint(ByVal host As Object, ByVal ctlName As String, ByVal caption As String, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single, ByVal heightPt As Single) As Object

    Dim c As Object
    Set c = UiLabel(host, ctlName, caption, leftPt, topPt, widthPt, heightPt)
    c.ForeColor = UI_MUTED
    Set UiHint = c
End Function

Public Function UiText( _
    ByVal host As Object, ByVal ctlName As String, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single, _
    Optional ByVal maxLen As Long = 0, Optional ByVal secret As Boolean = False) As Object

    Dim c As Object
    Set c = UiAdd(host, "Forms.TextBox.1", ctlName, leftPt, topPt, widthPt, 18)
    If maxLen > 0 Then c.MaxLength = maxLen
    If secret Then c.PasswordChar = "*"
    Set UiText = c
End Function

Public Function UiMemo( _
    ByVal host As Object, ByVal ctlName As String, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single, ByVal heightPt As Single) As Object

    Dim c As Object
    Set c = UiAdd(host, "Forms.TextBox.1", ctlName, leftPt, topPt, widthPt, heightPt)
    c.MultiLine = True
    c.WordWrap = True
    c.EnterKeyBehavior = True
    c.ScrollBars = 2 ' fmScrollBarsVertical
    Set UiMemo = c
End Function

Public Function UiButton( _
    ByVal host As Object, ByVal ctlName As String, ByVal caption As String, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single) As Object

    Dim c As Object
    Set c = UiAdd(host, "Forms.CommandButton.1", ctlName, leftPt, topPt, widthPt, 22)
    c.Caption = caption
    c.TakeFocusOnClick = False
    Set UiButton = c
End Function

Public Function UiCombo( _
    ByVal host As Object, ByVal ctlName As String, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single) As Object

    Dim c As Object
    Set c = UiAdd(host, "Forms.ComboBox.1", ctlName, leftPt, topPt, widthPt, 18)
    c.Style = 2 ' fmStyleDropDownList
    Set UiCombo = c
End Function

Public Function UiList( _
    ByVal host As Object, ByVal ctlName As String, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single, ByVal heightPt As Single, _
    ByVal columnWidths As String) As Object

    Dim c As Object
    Dim n As Long
    Set c = UiAdd(host, "Forms.ListBox.1", ctlName, leftPt, topPt, widthPt, heightPt)
    n = UBound(Split(columnWidths, ";")) + 1
    c.ColumnCount = n
    c.ColumnWidths = columnWidths
    c.IntegralHeight = False
    Set UiList = c
End Function

Public Function UiCheck( _
    ByVal host As Object, ByVal ctlName As String, ByVal caption As String, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single) As Object

    Dim c As Object
    Set c = UiAdd(host, "Forms.CheckBox.1", ctlName, leftPt, topPt, widthPt, 18)
    c.Caption = caption
    c.BackStyle = 0
    Set UiCheck = c
End Function

Public Function UiOption( _
    ByVal host As Object, ByVal ctlName As String, ByVal caption As String, ByVal groupName As String, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single) As Object

    Dim c As Object
    Set c = UiAdd(host, "Forms.OptionButton.1", ctlName, leftPt, topPt, widthPt, 18)
    c.Caption = caption
    c.GroupName = groupName
    c.BackStyle = 0
    Set UiOption = c
End Function

' Column captions above a list box, lined up with its ColumnWidths (points).
Public Sub UiColumnHeads( _
    ByVal host As Object, ByVal prefix As String, ByVal captions As Variant, ByVal widths As Variant, _
    ByVal leftPt As Single, ByVal topPt As Single)

    Dim i As Long
    Dim x As Single
    Dim c As Object
    x = leftPt + 3
    For i = LBound(captions) To UBound(captions)
        If CSng(widths(i)) > 0 Then
            Set c = UiLabel(host, prefix & CStr(i), CStr(captions(i)), x, topPt, CSng(widths(i)), 12)
            c.Font.Bold = True
            c.Font.Size = 8
            c.ForeColor = UI_MUTED
            c.WordWrap = False
        End If
        x = x + CSng(widths(i))
    Next i
End Sub

Public Function UiWidthsText(ByVal widths As Variant) As String
    Dim i As Long
    Dim s As String
    For i = LBound(widths) To UBound(widths)
        If i > LBound(widths) Then s = s & ";"
        s = s & CStr(widths(i)) & " pt"
    Next i
    UiWidthsText = s
End Function

' One summary card: a coloured box with a large number and a caption under it.
' Returns the number label. The caption label is named ctlName & "Cap".
Public Function UiCard( _
    ByVal host As Object, ByVal ctlName As String, ByVal caption As String, ByVal backColor As Long, _
    ByVal leftPt As Single, ByVal topPt As Single, ByVal widthPt As Single, ByVal heightPt As Single) As Object

    Dim box As Object
    Dim num As Object
    Dim cap As Object
    Set box = UiAdd(host, "Forms.Label.1", ctlName & "Box", leftPt, topPt, widthPt, heightPt)
    box.Caption = ""
    box.BackStyle = 1 ' fmBackStyleOpaque
    box.BackColor = backColor
    box.BorderStyle = 1 ' fmBorderStyleSingle
    box.BorderColor = UI_LINE
    Set num = UiLabel(host, ctlName, "0", leftPt + 8, topPt + 4, widthPt - 16, 22)
    num.Font.Size = 16
    num.Font.Bold = True
    num.ForeColor = UI_TEAL
    num.WordWrap = False
    Set cap = UiLabel(host, ctlName & "Cap", caption, leftPt + 8, topPt + 27, widthPt - 16, 14)
    cap.Font.Size = 8
    cap.ForeColor = UI_MUTED
    cap.WordWrap = False
    Set UiCard = num
End Function

' Selects a value in a drop-down list without raising an error when it is not listed.
Public Sub UiPick(ByVal cbo As Object, ByVal value As String)
    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If StrComp(CStr(cbo.List(i)), value, vbTextCompare) = 0 Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i
    cbo.ListIndex = -1
End Sub

Public Function UiStamp(ByVal raw As String) As String
    UiStamp = Left$(Replace(raw, "T", " "), 16)
End Function
