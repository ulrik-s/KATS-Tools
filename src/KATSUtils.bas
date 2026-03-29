Option Explicit

' ============================================================
' Shared helpers
' ============================================================

Public Function Dbl2Str(dbl As Double) As String
    Dbl2Str = Replace(CStr(dbl), ".", ",")
End Function


Public Function FindSumRow(ByVal t As Table) As Long
    Dim r As Long
    For r = 1 To t.rows.count
        If LCase$(Trim$(CellTextSafe(t, r, 1))) = "summa" Then
            FindSumRow = r
            Exit Function
        End If
    Next r
    FindSumRow = 0
End Function

' Safe table cell text read (merged/missing cells => "")
Public Function CellTextSafe(ByVal t As Table, ByVal row As Long, ByVal col As Long) As String
    On Error GoTo Missing
    Dim s As String
    s = t.Cell(row, col).Range.text
    s = Replace(s, Chr(7), "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    CellTextSafe = Trim$(s)
    Exit Function
Missing:
    CellTextSafe = ""
End Function

Public Sub CellSetTextSafe(ByVal t As Table, ByVal row As Long, ByVal col As Long, ByVal value As String)
    On Error GoTo Missing
    t.Cell(row, col).Range.text = value
    Exit Sub
Missing:
End Sub

Public Function LooksLikeIsoDate(ByVal s As String) As Boolean
    s = Trim$(s)
    LooksLikeIsoDate = (s Like "####-##-##")
End Function

' Robust Swedish-ish number parsing to Currency:
' Keeps digits, one decimal separator, optional leading '-'.
Public Function SvToCurrency(ByVal s As String) As Currency
    s = Trim$(s)
    If Len(s) = 0 Then SvToCurrency = 0@: Exit Function

    Dim i As Long, ch As String, out As String
    Dim sawDot As Boolean

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If ch = "-" And Len(out) = 0 Then
            out = "-"
        ElseIf ch >= "0" And ch <= "9" Then
            out = out & ch
        ElseIf (ch = "," Or ch = ".") And Not sawDot Then
            out = out & "."
            sawDot = True
        End If
    Next i

    If out = "" Or out = "-" Or out = "." Or out = "-." Then
        SvToCurrency = 0@
    Else
        SvToCurrency = CCur(Val(out))
    End If
End Function

' Half-away-from-zero rounding to integer
Public Function RoundHalfAwayFromZeroToLong(ByVal v As Currency) As Long
    If v >= 0@ Then
        RoundHalfAwayFromZeroToLong = CLng(Int(v + 0.5@))
    Else
        RoundHalfAwayFromZeroToLong = -CLng(Int(-v + 0.5@))
    End If
End Function

' 4209 -> "4 209"
Public Function FormatSvInt(ByVal n As Long) As String
    Dim s As String, neg As Boolean
    s = CStr(n)
    If Left$(s, 1) = "-" Then
        neg = True
        s = Mid$(s, 2)
    End If

    Dim out As String, i As Long, cnt As Long
    out = ""
    cnt = 0

    For i = Len(s) To 1 Step -1
        out = Mid$(s, i, 1) & out
        cnt = cnt + 1
        If (cnt Mod 3 = 0) And (i > 1) Then out = " " & out
    Next i

    If neg Then out = "-" & out
    FormatSvInt = out
End Function

' ============================================================
' Test helpers (optional)
' ============================================================

Public Sub Test_FillBookmarks()
    Dim doc As Document
    Set doc = ActiveDocument

    PutBookmarkText doc, "Mottagaradress", _
        "ACME AB" & vbCr & _
        "Gatan 1" & vbCr & _
        "123 45 Malmö"

    PutBookmarkText doc, "Arendenummer", "KATS-12345"
    PutBookmarkText doc, "Datum", format(Date, "yyyy-mm-dd")
    PutBookmarkText doc, "DagensDatum", format(Date, "yyyy-mm-dd")
End Sub

Private Sub PutBookmarkText(ByVal doc As Document, ByVal bm As String, Optional ByVal value As String = "")
    If Not doc.Bookmarks.Exists(bm) Then Exit Sub
    Dim rng As Range
    Set rng = doc.Bookmarks(bm).Range
    rng.text = value
    doc.Bookmarks.Add name:=bm, Range:=rng
End Sub


' ============================================================
' Helpers: parsing/time
' ============================================================

Public Function IsoDateToDate(ByVal s As String) As Date
    IsoDateToDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Right$(s, 2)))
End Function

Public Function TryExtractHHMM(ByVal s As String, ByRef hh As Long, ByRef mm As Long) As Boolean
    Dim i As Long, startAt As Long
    startAt = InStr(1, s, "kl", vbTextCompare)
    If startAt = 0 Then startAt = 1

    For i = startAt To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then Exit For
    Next i
    If i > Len(s) - 4 Then Exit Function

    If Not IsDigit(Mid$(s, i, 1)) Or Not IsDigit(Mid$(s, i + 1, 1)) Then Exit Function
    hh = CLng(Mid$(s, i, 2))

    Dim sep As String
    sep = Mid$(s, i + 2, 1)
    If sep <> ":" And sep <> "." Then Exit Function

    If Not IsDigit(Mid$(s, i + 3, 1)) Or Not IsDigit(Mid$(s, i + 4, 1)) Then Exit Function
    mm = CLng(Mid$(s, i + 3, 2))

    If hh < 0 Or hh > 23 Then Exit Function
    If mm < 0 Or mm > 59 Then Exit Function

    TryExtractHHMM = True
End Function

Public Function IsDigit(ByVal ch As String) As Boolean
    IsDigit = (ch >= "0" And ch <= "9")
End Function

Public Function HasAnyDigit(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            HasAnyDigit = True
            Exit Function
        End If
    Next i
    HasAnyDigit = False
End Function

Public Function FindTableRowContaining(ByVal t As Table, ByVal needle As String, Optional ByVal col As Long = 0) As Long
    Dim r As Long, c As Long
    For r = 1 To t.rows.count
        If col > 0 Then
            If InStr(1, CellTextSafe(t, r, col), needle, vbTextCompare) > 0 Then
                FindTableRowContaining = r
                Exit Function
            End If
        Else
            For c = 1 To t.Columns.count
                If InStr(1, CellTextSafe(t, r, c), needle, vbTextCompare) > 0 Then
                    FindTableRowContaining = r
                    Exit Function
                End If
            Next c
        End If
    Next r
    FindTableRowContaining = 0
End Function

Public Function TableRowContains(ByVal t As Table, ByVal row As Long, ByVal needle As String, Optional ByVal col As Long = 0) As Boolean
    Dim c As Long

    If col > 0 Then
        TableRowContains = (InStr(1, CellTextSafe(t, row, col), needle, vbTextCompare) > 0)
        Exit Function
    End If

    For c = 1 To t.Columns.count
        If InStr(1, CellTextSafe(t, row, c), needle, vbTextCompare) > 0 Then
            TableRowContains = True
            Exit Function
        End If
    Next c
End Function

Public Function RoundCurrencyToDecimals(ByVal v As Currency, ByVal decimals As Long) As Currency
    Dim s As Currency
    Select Case decimals
        Case 0: s = 1@
        Case 1: s = 10@
        Case 2: s = 100@
        Case Else
            ' keep it simple: support 0..2
            s = 100@
            decimals = 2
    End Select

    Dim scaled As Currency
    scaled = v * s

    Dim i As Long
    i = RoundHalfAwayFromZeroToLong(scaled)

    RoundCurrencyToDecimals = CCur(i) / s
End Function

Public Function ParseRateKr(ByVal s As String) As Currency
    Dim p As Long
    p = InStr(1, s, "à", vbTextCompare)
    If p = 0 Then p = InStr(1, s, "a", vbTextCompare) ' fallback om 'à' saknas
    If p > 0 Then
        ParseRateKr = SvToCurrency(Mid$(s, p + 1))
    Else
        ParseRateKr = SvToCurrency(s)
    End If
End Function

Public Function FormatSvDecimal(ByVal v As Currency, ByVal decimals As Long) As String
    Dim fmt As String
    fmt = "0"
    If decimals > 0 Then fmt = fmt & "." & String$(decimals, "0")

    Dim s As String
    s = format$(v, fmt)
    s = Replace(s, ".", ",")
    FormatSvDecimal = s
End Function

Public Function FormatSvMoney(ByVal v As Currency) As String
    v = RoundCurrencyToDecimals(v, 2)

    Dim neg As Boolean
    If v < 0@ Then
        neg = True
        v = -v
    End If

    Dim centsTotal As Currency
    centsTotal = v * 100@

    Dim cents As Long
    cents = RoundHalfAwayFromZeroToLong(centsTotal)

    Dim kronor As Long
    kronor = cents \ 100

    Dim ore As Long
    ore = cents Mod 100

    Dim s As String
    s = FormatSvInt(kronor) & "," & Right$("0" & CStr(ore), 2) & " kr"

    If neg Then s = "-" & s
    FormatSvMoney = s
End Function


Public Sub SetValueInLastCell(ByVal t As Table, ByVal row As Long, ByVal value As String)
    Dim lastCol As Long
    lastCol = t.Columns.count
    CellSetTextSafe t, row, lastCol, value
End Sub

Function SwedishDateText(Optional ByVal City As String = "Lund") As String
    Dim months As Variant
    months = Array("", "januari", "februari", "mars", "april", "maj", "juni", _
                      "juli", "augusti", "september", "oktober", "november", "december")

    SwedishDateText = City & " den " & Day(Date) & " " & months(Month(Date)) & " " & Year(Date)
End Function

