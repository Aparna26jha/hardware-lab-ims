Attribute VB_Name = "modSecurity"
Option Explicit

' Lab access codes are stored as SHA-256 hashes. The workbook never calls an
' outside program: the hash is computed in VBA so a locked-down laptop can run it.
' Sample codes typed at sign-in stay the same. The Access file stores v1: plus 64 hex digits.

Public Function IMS_HashAccessCode(ByVal accessCode As String) As String
    IMS_HashAccessCode = "v1:" & Sha256Hex("HWLAB|" & accessCode)
End Function

Public Function IMS_AccessCodeMatches(ByVal typedCode As String, ByVal storedCode As String) As Boolean
    storedCode = CStr(storedCode)
    If Left$(storedCode, 3) = "v1:" Then
        IMS_AccessCodeMatches = (StrComp(storedCode, IMS_HashAccessCode(typedCode), vbBinaryCompare) = 0)
    Else
        IMS_AccessCodeMatches = (StrComp(storedCode, typedCode, vbBinaryCompare) = 0)
    End If
End Function

' Custodians follow the lab flow. A Super Admin may correct any status because that role has complete access.
Public Function IMS_StatusMoveAllowed( _
    ByVal fromStatus As String, _
    ByVal toStatus As String, _
    ByVal roleName As String) As Boolean

    If roleName = "Super Admin" Then
        IMS_StatusMoveAllowed = True
        Exit Function
    End If
    If fromStatus = toStatus Then
        IMS_StatusMoveAllowed = True
        Exit Function
    End If
    If fromStatus = "Scrapped/Returned" Then Exit Function

    Select Case fromStatus
        Case "Available in Lab"
            IMS_StatusMoveAllowed = (toStatus = "Assigned to Engineer" _
                Or toStatus = "Under Maintenance" Or toStatus = "Scrapped/Returned")
        Case "Assigned to Engineer"
            IMS_StatusMoveAllowed = (toStatus = "Available in Lab" _
                Or toStatus = "Under Maintenance" Or toStatus = "Scrapped/Returned")
        Case "Under Maintenance"
            IMS_StatusMoveAllowed = (toStatus = "Available in Lab" _
                Or toStatus = "Assigned to Engineer" Or toStatus = "Scrapped/Returned")
    End Select
End Function

Public Function IMS_ExpiryNote(ByVal rawDate As String) As String
    Dim iso As String
    Dim exp As Date
    iso = Left$(Trim$(rawDate), 10)
    If Not IMS_IsIsoDate(iso) Then Exit Function
    exp = DateSerial(CLng(Left$(iso, 4)), CLng(Mid$(iso, 6, 2)), CLng(Mid$(iso, 9, 2)))
    If exp < Date Then
        IMS_ExpiryNote = " (Expired)"
    ElseIf exp <= Date + 30 Then
        IMS_ExpiryNote = " (Expiring soon)"
    End If
End Function

Private Function Sha256Hex(ByVal text As String) As String
    Dim src() As Byte
    Dim buf() As Byte
    Dim k(0 To 63) As Long
    Dim w(0 To 63) As Long
    Dim srcLen As Long
    Dim total As Long
    Dim pos As Long
    Dim i As Long
    Dim bitLen As Double
    Dim remain As Double
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long, g As Long, h As Long
    Dim h0 As Long, h1 As Long, h2 As Long, h3 As Long
    Dim h4 As Long, h5 As Long, h6 As Long, h7 As Long
    Dim s0 As Long, s1 As Long, t1 As Long, t2 As Long
    Dim ch As Long, maj As Long

    LoadSha256K k
    src = Utf8Bytes(text, srcLen)
    bitLen = CDbl(srcLen) * 8#
    total = srcLen + 1
    Do While (total Mod 64) <> 56
        total = total + 1
    Loop
    total = total + 8
    ReDim buf(0 To total - 1)
    For i = 0 To srcLen - 1
        buf(i) = src(i)
    Next i
    buf(srcLen) = 128
    remain = bitLen
    For i = 0 To 7
        buf(total - 1 - i) = CByte(remain - 256# * Int(remain / 256#))
        remain = Int(remain / 256#)
    Next i

    h0 = &H6A09E667: h1 = &HBB67AE85: h2 = &H3C6EF372: h3 = &HA54FF53A
    h4 = &H510E527F: h5 = &H9B05688C: h6 = &H1F83D9AB: h7 = &H5BE0CD19

    pos = 0
    Do While pos < total
        For i = 0 To 15
            w(i) = BeWord(buf, pos + i * 4)
        Next i
        For i = 16 To 63
            s0 = Ror32(w(i - 15), 7) Xor Ror32(w(i - 15), 18) Xor Shr32(w(i - 15), 3)
            s1 = Ror32(w(i - 2), 17) Xor Ror32(w(i - 2), 19) Xor Shr32(w(i - 2), 10)
            w(i) = Add5(w(i - 16), s0, w(i - 7), s1, 0)
        Next i
        a = h0: b = h1: c = h2: d = h3
        e = h4: f = h5: g = h6: h = h7
        For i = 0 To 63
            s1 = Ror32(e, 6) Xor Ror32(e, 11) Xor Ror32(e, 25)
            ch = (e And f) Xor ((Not e) And g)
            t1 = Add5(h, s1, ch, k(i), w(i))
            s0 = Ror32(a, 2) Xor Ror32(a, 13) Xor Ror32(a, 22)
            maj = (a And b) Xor (a And c) Xor (b And c)
            t2 = UAdd(s0, maj)
            h = g: g = f: f = e
            e = UAdd(d, t1)
            d = c: c = b: b = a
            a = UAdd(t1, t2)
        Next i
        h0 = UAdd(h0, a): h1 = UAdd(h1, b): h2 = UAdd(h2, c): h3 = UAdd(h3, d)
        h4 = UAdd(h4, e): h5 = UAdd(h5, f): h6 = UAdd(h6, g): h7 = UAdd(h7, h)
        pos = pos + 64
    Loop
    Sha256Hex = Hex8(h0) & Hex8(h1) & Hex8(h2) & Hex8(h3) & Hex8(h4) & Hex8(h5) & Hex8(h6) & Hex8(h7)
End Function

Private Sub LoadSha256K(ByRef k() As Long)
    k(0) = &H428A2F98: k(1) = &H71374491: k(2) = &HB5C0FBCF: k(3) = &HE9B5DBA5
    k(4) = &H3956C25B: k(5) = &H59F111F1: k(6) = &H923F82A4: k(7) = &HAB1C5ED5
    k(8) = &HD807AA98: k(9) = &H12835B01: k(10) = &H243185BE: k(11) = &H550C7DC3
    k(12) = &H72BE5D74: k(13) = &H80DEB1FE: k(14) = &H9BDC06A7: k(15) = &HC19BF174
    k(16) = &HE49B69C1: k(17) = &HEFBE4786: k(18) = &HFC19DC6: k(19) = &H240CA1CC
    k(20) = &H2DE92C6F: k(21) = &H4A7484AA: k(22) = &H5CB0A9DC: k(23) = &H76F988DA
    k(24) = &H983E5152: k(25) = &HA831C66D: k(26) = &HB00327C8: k(27) = &HBF597FC7
    k(28) = &HC6E00BF3: k(29) = &HD5A79147: k(30) = &H6CA6351: k(31) = &H14292967
    k(32) = &H27B70A85: k(33) = &H2E1B2138: k(34) = &H4D2C6DFC: k(35) = &H53380D13
    k(36) = &H650A7354: k(37) = &H766A0ABB: k(38) = &H81C2C92E: k(39) = &H92722C85
    k(40) = &HA2BFE8A1: k(41) = &HA81A664B: k(42) = &HC24B8B70: k(43) = &HC76C51A3
    k(44) = &HD192E819: k(45) = &HD6990624: k(46) = &HF40E3585: k(47) = &H106AA070
    k(48) = &H19A4C116: k(49) = &H1E376C08: k(50) = &H2748774C: k(51) = &H34B0BCB5
    k(52) = &H391C0CB3: k(53) = &H4ED8AA4A: k(54) = &H5B9CCA4F: k(55) = &H682E6FF3
    k(56) = &H748F82EE: k(57) = &H78A5636F: k(58) = &H84C87814: k(59) = &H8CC70208
    k(60) = &H90BEFFFA: k(61) = &HA4506CEB: k(62) = &HBEF9A3F7: k(63) = &HC67178F2
End Sub

Private Function Utf8Bytes(ByVal text As String, ByRef count As Long) As Byte()
    Dim tmp() As Byte
    Dim out() As Byte
    Dim i As Long
    Dim ch As Long
    Dim n As Long
    count = 0
    ReDim tmp(0 To Len(text) * 3 + 1)
    n = 0
    For i = 1 To Len(text)
        ch = AscW(Mid$(text, i, 1))
        If ch < 0 Then ch = ch + 65536
        If ch < 128 Then
            tmp(n) = CByte(ch)
            n = n + 1
        ElseIf ch < 2048 Then
            tmp(n) = CByte(192 + (ch \ 64))
            tmp(n + 1) = CByte(128 + (ch And 63))
            n = n + 2
        Else
            tmp(n) = CByte(224 + (ch \ 4096))
            tmp(n + 1) = CByte(128 + ((ch \ 64) And 63))
            tmp(n + 2) = CByte(128 + (ch And 63))
            n = n + 3
        End If
    Next i
    count = n
    If n = 0 Then
        ReDim out(0 To 0)
    Else
        ReDim out(0 To n - 1)
        For i = 0 To n - 1
            out(i) = tmp(i)
        Next i
    End If
    Utf8Bytes = out
End Function

Private Function BeWord(ByRef buf() As Byte, ByVal i As Long) As Long
    BeWord = ToLong( _
        CDbl(buf(i)) * 16777216# + _
        CDbl(buf(i + 1)) * 65536# + _
        CDbl(buf(i + 2)) * 256# + _
        CDbl(buf(i + 3)))
End Function

Private Function Unsigned(ByVal v As Long) As Double
    If v < 0 Then
        Unsigned = CDbl(v) + 4294967296#
    Else
        Unsigned = CDbl(v)
    End If
End Function

Private Function ToLong(ByVal u As Double) As Long
    u = u - 4294967296# * Int(u / 4294967296#)
    If u >= 2147483648# Then
        ToLong = CLng(u - 4294967296#)
    Else
        ToLong = CLng(u)
    End If
End Function

Private Function UAdd(ByVal a As Long, ByVal b As Long) As Long
    Dim s As Double
    s = Unsigned(a) + Unsigned(b)
    If s >= 4294967296# Then s = s - 4294967296#
    UAdd = ToLong(s)
End Function

Private Function Add5(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal e As Long) As Long
    Add5 = UAdd(UAdd(UAdd(UAdd(a, b), c), d), e)
End Function

Private Function Ror32(ByVal v As Long, ByVal n As Long) As Long
    Dim u As Double
    Dim span As Double
    Dim low As Double
    Dim high As Double
    u = Unsigned(v)
    span = 2# ^ n
    high = Int(u / span)
    low = u - span * high
    Ror32 = ToLong(high + low * (2# ^ (32 - n)))
End Function

Private Function Shr32(ByVal v As Long, ByVal n As Long) As Long
    Shr32 = ToLong(Int(Unsigned(v) / (2# ^ n)))
End Function

Private Function Hex8(ByVal v As Long) As String
    Hex8 = Right$("00000000" & Hex$(v), 8)
End Function
