Option Explicit

' ============================================================
' Shared helpers
' ============================================================

Public Function RequireSingleTable(ByVal content As Range) As Table
    If content Is Nothing Then Exit Function
    If content.Tables.count = 0 Then Exit Function

    Set RequireSingleTable = content.Tables(1)
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

Public Sub ReplaceAllLiteral(ByVal rng As Range, ByVal findText As String, ByVal replaceText As String)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = findText
        .Replacement.text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
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
' Helpers: parsing/time
' ============================================================

Public Function IsoDateToDate(ByVal s As String) As Date
    IsoDateToDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Right$(s, 2)))
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
            If RegexContainsLoose(CellTextSafe(t, r, col), needle) Then
                FindTableRowContaining = r
                Exit Function
            End If
        Else
            For c = 1 To t.Columns.count
                If RegexContainsLoose(CellTextSafe(t, r, c), needle) Then
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
        TableRowContains = RegexContainsLoose(CellTextSafe(t, row, col), needle)
        Exit Function
    End If

    For c = 1 To t.Columns.count
        If RegexContainsLoose(CellTextSafe(t, row, c), needle) Then
            TableRowContains = True
            Exit Function
        End If
    Next c
End Function

Public Function RegexContainsLoose(ByVal haystack As String, ByVal needle As String) As Boolean
    Dim rx As RegexTy
    InitializeRegex rx, needle, True
    RegexContainsLoose = Test(rx, haystack)
End Function

Public Function FirstNonEmptyLine(ByRef lines() As String) As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            FirstNonEmptyLine = Trim$(lines(i))
            Exit Function
        End If
    Next i
    FirstNonEmptyLine = ""
End Function

Public Function TryExtractPostort(ByVal s As String, ByRef postort As String) As Boolean
    Dim i As Long
    Dim line As String

    line = Trim$(s)

    For i = 1 To Len(line) - 6
        If Mid$(line, i, 3) Like "###" _
           And Mid$(line, i + 3, 1) = " " _
           And Mid$(line, i + 4, 2) Like "##" _
           And Mid$(line, i + 6, 1) = " " Then

            postort = Trim$(Mid$(line, i + 7))
            postort = NormalizePostort(postort)
            TryExtractPostort = (Len(postort) > 0)
            Exit Function
        End If
    Next i

    postort = ""
    TryExtractPostort = False
End Function

Public Function NormalizePostort(ByVal s As String) As String
    Dim parts() As String
    Dim i As Long

    s = LCase$(Trim$(s))
    If Len(s) = 0 Then
        NormalizePostort = ""
        Exit Function
    End If

    parts = Split(s, " ")

    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            parts(i) = UCase$(Left$(parts(i), 1)) & Mid$(parts(i), 2)
        End If
    Next i

    NormalizePostort = Join(parts, " ")
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
    p = InStr(1, s, ChrW$(225), vbTextCompare) ' á
    If p = 0 Then p = InStr(1, s, ChrW$(224), vbTextCompare) ' à (bakåtkompatibilitet)
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

Public Function FirstNonEmptyString(ByVal primary As String, Optional ByVal fallback As String = "", Optional ByVal defaultValue As String = "") As String
    primary = Trim$(primary)
    If Len(primary) > 0 Then
        FirstNonEmptyString = primary
        Exit Function
    End If

    fallback = Trim$(fallback)
    If Len(fallback) > 0 Then
        FirstNonEmptyString = fallback
    Else
        FirstNonEmptyString = defaultValue
    End If
End Function

Public Function JoinPath(ByVal folderPath As String, ByVal fileName As String) As String
#If Mac Then
    If Right$(folderPath, 1) = "/" Then
        JoinPath = folderPath & fileName
    Else
        JoinPath = folderPath & "/" & fileName
    End If
#Else
    If Right$(folderPath, 1) = "\" Then
        JoinPath = folderPath & fileName
    Else
        JoinPath = folderPath & "\" & fileName
    End If
#End If
End Function

Public Sub RenderSignatureBlock(ByVal content As Range, ByVal postort As String, ByVal namn As String, ByVal titel As String)
    content.text = SwedishDateText(FirstNonEmptyString(postort, "", "Lund")) & vbCr & vbCr & namn & vbCr & titel
End Sub

Public Function SwedishDateText(Optional ByVal City As String = "Lund") As String
    Dim months As Variant
    months = Array("", "januari", "februari", "mars", "april", "maj", "juni", _
                      "juli", "augusti", "september", "oktober", "november", "december")

    SwedishDateText = City & " den " & Day(Date) & " " & months(Month(Date)) & " " & Year(Date)
End Function
