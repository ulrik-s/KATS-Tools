Option Explicit

' ============================================================
' Global state
' ============================================================

Private gHasHearingStart As Boolean
Private gHearingStart As Date
Private gPostort As String

Private Enum ArCategory
    arArvode = 1
    arArvodeHelg = 2
    arTidsspillan = 3
    arTidsspillanOvrigTid = 4
End Enum

Private gCategoryHours(1 To 4) As Currency
Private gHasCategoryHours(1 To 4) As Boolean

Private Enum MoneyStateKey
    msArvodeExMoms = 1
    msUtlaggExMoms = 2
    msUtlaggEjMoms = 3
End Enum

Private gMoneyState(1 To 3) As Currency
Private gHasMoneyState(1 To 3) As Boolean

Private gIsTaxemal As Boolean
Private gTaxLevel As Long
Private gHearingMinutes As Long

Private gRegexInitialized As Boolean
Private gRxHearingTime As RegexTy
Private gRxHearingTaxa As RegexTy

Private Sub EnsureHearingRegexInitialized()
    If gRegexInitialized Then Exit Sub

    ' Matchar hearing-raden och fångar tid som hör till hearingen
    InitializeRegex gRxHearingTime, _
        "medverkat vid (?:huvud)?förhandling från\s*(?:kl\.?\s*)?([0-9]{1,2})(?:\s*[:.]\s*([0-9]{2}))?", _
        True

    ' Matchar bara taxemål när "enligt taxa" kommer direkt efter hearing-tiden
    InitializeRegex gRxHearingTaxa, _
        "medverkat vid (?:huvud)?förhandling från\s*(?:kl\.?\s*)?[0-9]{1,2}(?:\s*[:.]\s*[0-9]{2})?\s*[,;:]?\s*enligt taxa\b", _
        True

    gRegexInitialized = True
End Sub

' ============================================================
' Reset
' ============================================================

Public Sub ResetProcessorState()
    Dim i As Long

    gHasHearingStart = False
    gHearingStart = 0
    gPostort = ""

    gIsTaxemal = False
    gTaxLevel = 1
    gHearingMinutes = 0

    For i = LBound(gCategoryHours) To UBound(gCategoryHours)
        gCategoryHours(i) = 0@
        gHasCategoryHours(i) = False
    Next i

    For i = LBound(gMoneyState) To UBound(gMoneyState)
        gMoneyState(i) = 0@
        gHasMoneyState(i) = False
    Next i
End Sub

Private Sub SetCategoryHours(ByVal key As ArCategory, ByVal hours As Currency)
    gCategoryHours(key) = hours
    gHasCategoryHours(key) = True
End Sub

Private Function HasCategoryHours(ByVal key As ArCategory) As Boolean
    HasCategoryHours = gHasCategoryHours(key)
End Function

Private Function GetCategoryHours(ByVal key As ArCategory) As Currency
    GetCategoryHours = gCategoryHours(key)
End Function

Private Sub SetMoneyState(ByVal key As MoneyStateKey, ByVal value As Currency)
    gMoneyState(key) = value
    gHasMoneyState(key) = True
End Sub

Private Function HasMoneyState(ByVal key As MoneyStateKey) As Boolean
    HasMoneyState = gHasMoneyState(key)
End Function

Private Function GetMoneyState(ByVal key As MoneyStateKey) As Currency
    GetMoneyState = gMoneyState(key)
End Function

Private Sub DetectTaxCaseFromARTable(ByVal t As Table)
    Dim r As Long
    Dim c As Long

    gIsTaxemal = False
    gTaxLevel = 1

    For r = 1 To t.rows.count
        For c = 1 To t.Columns.count
            Dim s As String
            s = CellTextSafe(t, r, c)

            If Len(s) > 0 Then
                If IsTaxaForHearingText(s) Then
                    gIsTaxemal = True
                    gTaxLevel = 1
                    Exit Sub
                End If
            End If
        Next c
    Next r
End Sub

Private Function TryExtractHearingTime(ByVal s As String, ByRef hh As Long, ByRef mm As Long) As Boolean
    Dim matcher As MatcherStateTy
    Dim h As String
    Dim m As String

    hh = 0
    mm = 0

    EnsureHearingRegexInitialized

    If Not Match(matcher, gRxHearingTime, s) Then Exit Function

    h = GetCapture(matcher, s, 1)
    m = GetCapture(matcher, s, 2)

    If Len(h) = 0 Then Exit Function

    hh = CLng(h)

    If Len(m) = 0 Then
        mm = 0
    Else
        mm = CLng(m)
    End If

    If hh < 0 Or hh > 23 Then Exit Function
    If mm < 0 Or mm > 59 Then Exit Function

    TryExtractHearingTime = True
End Function

Private Function IsTaxaForHearingText(ByVal s As String) As Boolean
    EnsureHearingRegexInitialized
    IsTaxaForHearingText = Test(gRxHearingTaxa, s)
End Function

' ============================================================
' PROCESSORS
' ============================================================

' ---- UTLÄGGSSPECIFIKATION ----
Public Sub Process_UTLAGGSSPECIFIKATION(ByVal content As Range)
    If content.Tables.count = 0 Then Exit Sub

    Dim t As Table
    Set t = content.Tables(1)
    If t.Columns.count < 5 Then Exit Sub
    ProcessExpenseSection t, "Utlägg", msUtlaggExMoms, True
    ProcessExpenseSection t, "Utlägg momsfri", msUtlaggEjMoms, False

    AutoFitUtlaggTable t
    AddAirBeforeSectionHeadings t, 12
End Sub

Private Sub ProcessExpenseSection(ByVal t As Table, ByVal heading As String, ByVal moneyKey As MoneyStateKey, ByVal applyMileageRule As Boolean)
    Const COL_DATE As Long = 1
    Const COL_DESC As Long = 2
    Const COL_QTY As Long = 3
    Const COL_RATE As Long = 4
    Const COL_AMT As Long = 5

    Dim headingRow As Long
    Dim summaryRow As Long

    headingRow = FindSectionHeadingRow(t, heading)
    summaryRow = FindSectionSummaryRowAfterHeading(t, heading)

    If headingRow = 0 Or summaryRow = 0 Then Exit Sub

    Dim r As Long
    Dim sumSek As Currency
    sumSek = 0@

    For r = headingRow + 1 To summaryRow - 1
        Dim d As String
        d = CellTextSafe(t, r, COL_DATE)

        If LooksLikeIsoDate(d) Then
            Dim desc As String
            Dim qty As Currency
            Dim rate As Currency
            Dim amtExisting As Currency
            Dim amtSek As Long

            desc = CellTextSafe(t, r, COL_DESC)
            qty = SvToCurrency(CellTextSafe(t, r, COL_QTY))
            rate = SvToCurrency(CellTextSafe(t, r, COL_RATE))
            amtExisting = SvToCurrency(CellTextSafe(t, r, COL_AMT))

            If applyMileageRule Then
                If InStr(1, desc, "Milersättning", vbTextCompare) > 0 Then
                    rate = GetMileage()
                    CellSetTextSafe t, r, COL_RATE, FormatSvDecimal(rate, 2)
                End If
            End If

            If qty <> 0@ And rate <> 0@ Then
                amtSek = RoundHalfAwayFromZeroToLong(qty * rate)
            Else
                amtSek = RoundHalfAwayFromZeroToLong(amtExisting)
            End If

            CellSetTextSafe t, r, COL_AMT, FormatSvInt(amtSek)
            sumSek = sumSek + CCur(amtSek)
        End If
    Next r

    SetMoneyState moneyKey, RoundCurrencyToDecimals(sumSek, 0)
    CellSetTextSafe t, summaryRow, COL_QTY, ""
    CellSetTextSafe t, summaryRow, COL_AMT, FormatSvInt(RoundHalfAwayFromZeroToLong(sumSek))
End Sub

' ---- ARGRUPPERTIDERDATUMANTALSUMMA ----
Public Sub Process_ARGRUPPERTIDERDATUMANTALSUMMA(ByVal content As Range)
    If content.Tables.count = 0 Then Exit Sub

    Dim t As Table
    Set t = content.Tables(1)

    DetectTaxCaseFromARTable t
    CaptureHearingStartFromARTable t

    UpdateCategoryFromHeading t, 3, "Arvode", arArvode
    UpdateCategoryFromHeading t, 3, "Arvode helg", arArvodeHelg
    UpdateCategoryFromHeading t, 3, "Tidsspillan", arTidsspillan
    UpdateCategoryFromHeading t, 3, "Tidsspillan övrig tid", arTidsspillanOvrigTid

    Dim totalRow As Long
    totalRow = FindTableRowContaining(t, "Ärende, total", 1)
    If totalRow > 0 Then
        CellSetTextSafe t, totalRow, 1, ""
        CellSetTextSafe t, totalRow, 2, ""
        CellSetTextSafe t, totalRow, 3, ""
    End If

    AddAirBeforeSectionHeadings t, 12
End Sub

Private Function IsHeadingRow(ByVal t As Table, ByVal rowIndex As Long) As Boolean
    IsHeadingRow = (t.rows(rowIndex).Cells.count = 1)
End Function

Private Sub AddAirBeforeSectionHeadings(ByVal t As Table, Optional ByVal pointsBefore As Single = 8)
    Dim r As Long
    Dim c As Long
    Dim firstHeadingSeen As Boolean

    firstHeadingSeen = False

    For r = 1 To t.rows.count
        If IsHeadingRow(t, r) Then
            If firstHeadingSeen Then
                For c = 1 To t.rows(r).Cells.count
                    With t.Cell(r, c).Range.ParagraphFormat
                        .SpaceBefore = pointsBefore
                        .SpaceAfter = 0
                    End With
                Next c
            Else
                firstHeadingSeen = True
            End If
        End If
    Next r
End Sub

Private Sub UpdateCategoryFromHeading(ByVal t As Table, ByVal col As Long, ByVal heading As String, ByVal category As ArCategory)
    Dim sumHours As Currency
    sumHours = GetSumColumnWithHeading(t, col, heading)

    SetSumColumnWithHeading t, col, heading, sumHours
    SetCategoryHours category, RoundCurrencyToDecimals(sumHours, 2)
End Sub

Public Sub SetSumColumnWithHeading(ByVal t As Table, ByVal col As Long, ByVal heading As String, ByVal sumHours As Currency)
    Dim summaryRow As Long
    summaryRow = FindSectionSummaryRowAfterHeading(t, heading)
    If summaryRow = 0 Then Exit Sub

    CellSetTextSafe t, summaryRow, col, FormatSvDecimal(sumHours, 2)
End Sub

Public Function GetSumColumnWithHeading(ByVal t As Table, ByVal col As Long, ByVal heading As String) As Currency
    Dim headingRow As Long
    Dim summaryRow As Long

    headingRow = FindSectionHeadingRow(t, heading)
    summaryRow = FindSectionSummaryRowAfterHeading(t, heading)

    If headingRow = 0 Or summaryRow = 0 Then
        GetSumColumnWithHeading = 0@
        Exit Function
    End If

    Dim sumHours As Currency
    sumHours = 0@

    Dim r As Long
    For r = headingRow + 1 To summaryRow - 1
        If LooksLikeIsoDate(CellTextSafe(t, r, 1)) Then
            sumHours = sumHours + SvToCurrency(CellTextSafe(t, r, col))
        End If
    Next r

    GetSumColumnWithHeading = RoundCurrencyToDecimals(sumHours, 2)
End Function

Private Function FindHearingStartTextPos(ByVal s As String) As Long
    Dim needle As String
    Dim p As Long

    needle = "medverkat vid huvudförhandling från"
    p = InStr(1, s, needle, vbTextCompare)
    If p > 0 Then
        FindHearingStartTextPos = p + Len(needle)
        Exit Function
    End If

    needle = "medverkat vid förhandling från"
    p = InStr(1, s, needle, vbTextCompare)
    If p > 0 Then
        FindHearingStartTextPos = p + Len(needle)
        Exit Function
    End If

    FindHearingStartTextPos = 0
End Function

Private Sub CaptureHearingStartFromARTable(ByVal t As Table)
    Dim r As Long
    Dim c As Long

    For r = 1 To t.rows.count
        For c = 1 To t.Columns.count
            Dim s As String
            s = CellTextSafe(t, r, c)

            If Len(s) > 0 Then
                Dim hh As Long
                Dim mm As Long

                If TryExtractHearingTime(s, hh, mm) Then
                    Dim baseDate As Date
                    baseDate = Date

                    Dim ds As String
                    ds = CellTextSafe(t, r, 1)
                    If LooksLikeIsoDate(ds) Then baseDate = IsoDateToDate(ds)

                    gHearingStart = DateSerial(Year(baseDate), Month(baseDate), Day(baseDate)) + TimeSerial(hh, mm, 0)
                    gHasHearingStart = True

                    Dim minutes As Long
                    minutes = DateDiff("n", gHearingStart, Now)
                    If minutes < 0 Then minutes = minutes + (24& * 60&)

                    gHearingMinutes = minutes

                    Dim hours As Currency
                    hours = RoundCurrencyToDecimals(CCur(minutes) / 60@, 2)

                    CellSetTextSafe t, r, 3, FormatSvDecimal(hours, 2)
                    Exit Sub
                End If
            End If
        Next c
    Next r
End Sub

Private Function ShouldUseTaxa() As Boolean
    ShouldUseTaxa = (gIsTaxemal And gHasHearingStart And gHearingMinutes <= 225)
End Function

Private Function GetTaxAmountLevel1(ByVal hearingMinutes As Long) As Currency
    Select Case hearingMinutes
        Case 0 To 14: GetTaxAmountLevel1 = 2809@
        Case 15 To 29: GetTaxAmountLevel1 = 2980@
        Case 30 To 44: GetTaxAmountLevel1 = 3509@
        Case 45 To 59: GetTaxAmountLevel1 = 4049@
        Case 60 To 74: GetTaxAmountLevel1 = 4583@
        Case 75 To 89: GetTaxAmountLevel1 = 5106@
        Case 90 To 104: GetTaxAmountLevel1 = 5635@
        Case 105 To 119: GetTaxAmountLevel1 = 6164@
        Case 120 To 134: GetTaxAmountLevel1 = 6704@
        Case 135 To 149: GetTaxAmountLevel1 = 7227@
        Case 150 To 164: GetTaxAmountLevel1 = 7767@
        Case 165 To 179: GetTaxAmountLevel1 = 8301@
        Case 180 To 194: GetTaxAmountLevel1 = 8824@
        Case 195 To 209: GetTaxAmountLevel1 = 9364@
        Case 210 To 225: GetTaxAmountLevel1 = 9887@
        Case Else: GetTaxAmountLevel1 = 0@
    End Select
End Function

Private Sub ApplyTaxaRow(ByVal t As Table, ByVal rowIndex As Long, ByVal taxAmount As Currency)
    Dim hours As Long
    hours = 0

    If gHearingMinutes >= 60 Then
        hours = Int(gHearingMinutes / 60)
    End If

    Dim minutes As Integer
    minutes = gHearingMinutes - (hours * 60)

    Dim spec As String
    spec = ""

    If hours > 0 Then
        spec = CStr(hours) & " tim "
    End If

    spec = spec & CStr(minutes) & " min enligt taxa"

    CellSetTextSafe t, rowIndex, 2, spec
    CellSetTextSafe t, rowIndex, 3, FormatSvMoney(taxAmount)
End Sub

Private Sub Process_ARVODE(ByVal content As Range)
    If content.Tables.count = 0 Then Exit Sub

    Dim t As Table
    Set t = content.Tables(1)

    If t.rows.count < 6 Or t.Columns.count < 3 Then Exit Sub

    Const ROW_ARVODE As Long = 2
    Const ROW_ARVODE_HELG As Long = 3
    Const ROW_TIDSSPILLAN As Long = 4
    Const ROW_TIDSSPILLAN_OVRIG As Long = 5
    Const ROW_UTLAGG As Long = 6

    ' ===== Hämta utlägg
    If HasMoneyState(msUtlaggExMoms) Then
        CellSetTextSafe t, ROW_UTLAGG, 2, ""
        CellSetTextSafe t, ROW_UTLAGG, 3, FormatSvMoney(GetMoneyState(msUtlaggExMoms))
    End If

    ' ===== TAXEMÅL steg 1 =====
    If ShouldUseTaxa() Then
        Dim removeTidsspillan As Boolean
        removeTidsspillan = True

        Dim taxAmount As Currency
        taxAmount = GetTaxAmountLevel1(gHearingMinutes)

        ApplyTaxaRow t, ROW_ARVODE, taxAmount

        SetMoneyState msArvodeExMoms, RoundCurrencyToDecimals(taxAmount + MoneyFromRow(t, ROW_UTLAGG), 2)

        If GetCategoryHours(arTidsspillan) > 1# Then
            removeTidsspillan = False
            CellSetTextSafe t, ROW_TIDSSPILLAN, 1, "TIDSSPILLAN överstigande 1 tim"
            ApplyHoursRow t, ROW_TIDSSPILLAN, HasCategoryHours(arTidsspillan), (GetCategoryHours(arTidsspillan) - 1)

            Dim utlagg As Currency
            utlagg = MoneyFromRow(t, ROW_UTLAGG)

            Dim tidsspillan As Currency
            tidsspillan = MoneyFromRow(t, ROW_TIDSSPILLAN)

            Dim total As Currency
            total = utlagg + tidsspillan + taxAmount
            SetMoneyState msArvodeExMoms, RoundCurrencyToDecimals(total, 2)
        End If

        ' Ta bort rader som inte ska vara med i sammanställningen i taxemål
        DeleteArvodeRowIfZeroAmount t, ROW_UTLAGG
        t.rows(ROW_TIDSSPILLAN_OVRIG).Delete

        If removeTidsspillan Then
            t.rows(ROW_TIDSSPILLAN).Delete
        End If

        t.rows(ROW_ARVODE_HELG).Delete
        Exit Sub
    End If

    ' ===== Vanlig modell =====
    ApplyHoursRow t, ROW_ARVODE, HasCategoryHours(arArvode), GetCategoryHours(arArvode)
    ApplyHoursRow t, ROW_ARVODE_HELG, HasCategoryHours(arArvodeHelg), GetCategoryHours(arArvodeHelg)
    ApplyHoursRow t, ROW_TIDSSPILLAN, HasCategoryHours(arTidsspillan), GetCategoryHours(arTidsspillan)
    ApplyHoursRow t, ROW_TIDSSPILLAN_OVRIG, HasCategoryHours(arTidsspillanOvrigTid), GetCategoryHours(arTidsspillanOvrigTid)

    Dim totalExMoms As Currency
    totalExMoms = 0@

    totalExMoms = totalExMoms + MoneyFromRow(t, ROW_ARVODE)
    totalExMoms = totalExMoms + MoneyFromRow(t, ROW_ARVODE_HELG)
    totalExMoms = totalExMoms + MoneyFromRow(t, ROW_TIDSSPILLAN)
    totalExMoms = totalExMoms + MoneyFromRow(t, ROW_TIDSSPILLAN_OVRIG)
    totalExMoms = totalExMoms + MoneyFromRow(t, ROW_UTLAGG)

    SetMoneyState msArvodeExMoms, RoundCurrencyToDecimals(totalExMoms, 2)

    DeleteArvodeRowIfZeroAmount t, ROW_UTLAGG
    DeleteArvodeRowIfZeroAmount t, ROW_TIDSSPILLAN_OVRIG
    DeleteArvodeRowIfZeroAmount t, ROW_TIDSSPILLAN
    DeleteArvodeRowIfZeroAmount t, ROW_ARVODE_HELG
    DeleteArvodeRowIfZeroAmount t, ROW_ARVODE
End Sub

Private Sub ApplyHoursRow(ByVal t As Table, ByVal rowIndex As Long, ByVal hasHours As Boolean, ByVal hours As Currency)
    If Not hasHours Then Exit Sub

    hours = RoundCurrencyToDecimals(hours, 2)
    If hours = 0@ Then Exit Sub

    Dim rate As Currency
    rate = ParseRateKr(CellTextSafe(t, rowIndex, 2))
    If rate = 0@ Then Exit Sub

    Dim amount As Currency
    amount = RoundCurrencyToDecimals(hours * rate, 0)

    CellSetTextSafe t, rowIndex, 2, FormatSvDecimal(hours, 2) & " à " & FormatSvInt(CLng(rate)) & " kr"
    CellSetTextSafe t, rowIndex, 3, FormatSvMoney(amount)
End Sub

Private Sub ApplyHearingRow(ByVal t As Table, ByVal rowIndex As Long)
    If Not gHasHearingStart Then Exit Sub

    Dim minutes As Long
    minutes = DateDiff("n", gHearingStart, Now)
    If minutes < 0 Then minutes = minutes + (24& * 60&)

    Dim hours As Currency
    hours = RoundCurrencyToDecimals(CCur(minutes) / 60@, 1)
    If hours = 0@ Then Exit Sub

    Dim rate As Currency
    rate = ParseRateKr(CellTextSafe(t, rowIndex, 2))
    If rate = 0@ Then rate = 1626@

    Dim amount As Currency
    amount = RoundCurrencyToDecimals(hours * rate, 2)

    CellSetTextSafe t, rowIndex, 2, FormatSvDecimal(hours, 2) & " tim à " & FormatSvInt(CLng(rate)) & " kr"
    CellSetTextSafe t, rowIndex, 3, FormatSvMoney(amount)
End Sub

Private Function MoneyFromRow(ByVal t As Table, ByVal rowIndex As Long) As Currency
    If rowIndex < 1 Or rowIndex > t.rows.count Then
        MoneyFromRow = 0@
        Exit Function
    End If

    MoneyFromRow = RoundCurrencyToDecimals(SvToCurrency(CellTextSafe(t, rowIndex, 3)), 2)
End Function

Private Sub DeleteArvodeRowIfZeroAmount(ByVal t As Table, ByVal rowIndex As Long)
    If rowIndex < 1 Or rowIndex > t.rows.count Then Exit Sub

    Dim amtText As String
    amtText = Trim$(CellTextSafe(t, rowIndex, 3))

    If Not HasAnyDigit(amtText) Then
        t.rows(rowIndex).Delete
        Exit Sub
    End If

    If RoundCurrencyToDecimals(SvToCurrency(amtText), 2) = 0@ Then
        t.rows(rowIndex).Delete
    End If
End Sub

Public Sub Process_ARVODE_TOTAL(ByVal content As Range)
    If content.Tables.count = 0 Then Exit Sub
    If Not HasMoneyState(msArvodeExMoms) Then Exit Sub

    Dim t As Table
    Set t = content.Tables(1)

    Dim rowBeloppExkl As Long
    Dim rowMoms As Long
    Dim rowUtlaggEjMoms As Long
    Dim rowBeloppInkl As Long

    rowBeloppExkl = FindTableRowContaining(t, "Belopp exkl. moms", 1)
    rowMoms = FindTableRowContaining(t, "Moms (25%)", 1)
    rowUtlaggEjMoms = FindTableRowContaining(t, "UTLÄGG EJ MOMS", 1)
    rowBeloppInkl = FindTableRowContaining(t, "Belopp inkl. moms", 1)

    Dim arvodeExMoms As Currency
    arvodeExMoms = GetMoneyState(msArvodeExMoms)

    If rowBeloppExkl > 0 Then
        CellSetTextSafe t, rowBeloppExkl, 3, FormatSvMoney(arvodeExMoms)
    End If

    Dim moms As Currency
    moms = RoundCurrencyToDecimals(arvodeExMoms * 0.25@, 0)
    If rowMoms > 0 Then
        CellSetTextSafe t, rowMoms, 3, FormatSvMoney(moms)
    End If

    Dim utlaggEjMoms As Currency
    utlaggEjMoms = 0@
    If HasMoneyState(msUtlaggEjMoms) Then
        utlaggEjMoms = GetMoneyState(msUtlaggEjMoms)
        If rowUtlaggEjMoms > 0 Then
            utlaggEjMoms = RoundCurrencyToDecimals(utlaggEjMoms, 0)
            CellSetTextSafe t, rowUtlaggEjMoms, 3, FormatSvMoney(utlaggEjMoms)
        End If
    End If

    Dim incl As Currency
    incl = RoundCurrencyToDecimals(arvodeExMoms + moms + utlaggEjMoms, 0)
    If rowBeloppInkl > 0 Then
        CellSetTextSafe t, rowBeloppInkl, 3, FormatSvMoney(incl)
    End If

    If rowUtlaggEjMoms > 0 Then
        DeleteArvodeRowIfZeroAmount t, rowUtlaggEjMoms
    End If
End Sub

' ---- SIGNATUR ----
Public Sub Process_SIGNATUR(ByVal content As Range)
    Dim namn As String
    namn = GetFullName()

    Dim titel As String
    titel = GetTitle()

    content.text = SwedishDateText(gPostort) & vbCr & vbCr & namn & vbCr & titel
End Sub

' ---- MOTTAGARE ----
Public Sub Process_MOTTAGARE(ByVal content As Range)
    If content.Tables.count = 0 Then Exit Sub

    Dim t As Table
    Set t = content.Tables(1)
    If t.rows.count <> 1 Or t.Columns.count < 2 Then Exit Sub

    Dim raw As String
    raw = CellTextSafe(t, 1, 2)
    If Len(raw) = 0 Then Exit Sub

    raw = Replace(raw, vbCrLf, vbCr)
    raw = Replace(raw, Chr(11), vbCr)
    raw = Replace(raw, Chr(7), "")

    Dim lines() As String
    lines = Split(raw, vbCr)

    Dim firstLine As String
    firstLine = FirstNonEmptyLine(lines)
    If Len(firstLine) = 0 Then Exit Sub

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim postort As String
        If TryExtractPostort(lines(i), postort) Then
            gPostort = postort
            Exit For
        End If
    Next i

    CellSetTextSafe t, 1, 2, firstLine & vbCr & "via e-post"
End Sub

' ============================================================
' Generic table helpers
' ============================================================

Public Sub AutoFitUtlaggTable(ByVal t As Table)
    t.AllowAutoFit = True
    t.AutoFitBehavior wdAutoFitContent
End Sub

Public Function FindTableRowContaining(ByVal t As Table, ByVal needle As String, Optional ByVal col As Long = 0) As Long
    Dim r As Long
    Dim c As Long

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

Private Function CellTextSafe(ByVal t As Table, ByVal row As Long, ByVal col As Long) As String
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

Private Sub CellSetTextSafe(ByVal t As Table, ByVal row As Long, ByVal col As Long, ByVal value As String)
    On Error GoTo Missing
    t.Cell(row, col).Range.text = value
    Exit Sub
Missing:
End Sub

Private Function FindSectionHeadingRow(ByVal t As Table, ByVal heading As String) As Long
    Dim r As Long
    For r = 1 To t.rows.count
        If RowMatchesHeading(t, r, heading) Then
            FindSectionHeadingRow = r
            Exit Function
        End If
    Next r
    FindSectionHeadingRow = 0
End Function

Private Function FindSectionSummaryRowAfterHeading(ByVal t As Table, ByVal heading As String) As Long
    Dim headingRow As Long
    headingRow = FindSectionHeadingRow(t, heading)
    If headingRow = 0 Then Exit Function

    Dim r As Long
    For r = headingRow + 1 To t.rows.count
        If StrComp(Trim$(CellTextSafe(t, r, 1)), "Summa", vbTextCompare) = 0 Then
            FindSectionSummaryRowAfterHeading = r
            Exit Function
        End If
    Next r

    FindSectionSummaryRowAfterHeading = 0
End Function

Private Function RowMatchesHeading(ByVal t As Table, ByVal rowIndex As Long, ByVal heading As String) As Boolean
    Dim firstValue As String
    Dim c As Long
    Dim nonEmptyCount As Long

    firstValue = ""

    For c = 1 To t.rows(rowIndex).Cells.count
        Dim s As String
        s = Trim$(CellTextSafe(t, rowIndex, c))

        If Len(s) > 0 Then
            nonEmptyCount = nonEmptyCount + 1

            If firstValue = "" Then
                firstValue = s
            ElseIf StrComp(s, firstValue, vbTextCompare) <> 0 Then
                RowMatchesHeading = False
                Exit Function
            End If
        End If
    Next c

    If nonEmptyCount = 0 Then
        RowMatchesHeading = False
    Else
        RowMatchesHeading = (StrComp(firstValue, heading, vbTextCompare) = 0)
    End If
End Function

Private Function FindSumRow(ByVal t As Table) As Long
    Dim r As Long
    For r = 1 To t.rows.count
        If LCase$(Trim$(CellTextSafe(t, r, 1))) = "summa" Then
            FindSumRow = r
            Exit Function
        End If
    Next r
    FindSumRow = 0
End Function

Private Function FirstNonEmptyLine(ByRef lines() As String) As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            FirstNonEmptyLine = Trim$(lines(i))
            Exit Function
        End If
    Next i
    FirstNonEmptyLine = ""
End Function

' ============================================================
' Parsing helpers
' ============================================================

Private Function LooksLikeIsoDate(ByVal s As String) As Boolean
    s = Trim$(s)
    LooksLikeIsoDate = (s Like "####-##-##")
End Function

Private Function IsoDateToDate(ByVal s As String) As Date
    IsoDateToDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Right$(s, 2)))
End Function

Private Function IsDigit(ByVal ch As String) As Boolean
    IsDigit = (ch >= "0" And ch <= "9")
End Function

Private Function HasAnyDigit(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            HasAnyDigit = True
            Exit Function
        End If
    Next i
    HasAnyDigit = False
End Function

Private Function TryExtractPostort(ByVal s As String, ByRef postort As String) As Boolean
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

Private Function NormalizePostort(ByVal s As String) As String
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

' ============================================================
' Numeric helpers
' ============================================================

Private Function SvToCurrency(ByVal s As String) As Currency
    s = Trim$(s)
    If Len(s) = 0 Then
        SvToCurrency = 0@
        Exit Function
    End If

    Dim i As Long
    Dim ch As String
    Dim out As String
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

Private Function RoundHalfAwayFromZeroToLong(ByVal v As Currency) As Long
    If v >= 0@ Then
        RoundHalfAwayFromZeroToLong = CLng(Int(v + 0.5@))
    Else
        RoundHalfAwayFromZeroToLong = -CLng(Int(-v + 0.5@))
    End If
End Function

Private Function RoundCurrencyToDecimals(ByVal v As Currency, ByVal decimals As Long) As Currency
    Dim factor As Currency

    Select Case decimals
        Case 0: factor = 1@
        Case 1: factor = 10@
        Case 2: factor = 100@
        Case Else: factor = 100@
    End Select

    Dim scaledValue As Currency
    scaledValue = v * factor

    Dim roundedValue As Long
    roundedValue = RoundHalfAwayFromZeroToLong(scaledValue)

    RoundCurrencyToDecimals = CCur(roundedValue) / factor
End Function

Private Function FormatSvInt(ByVal n As Long) As String
    Dim s As String
    Dim neg As Boolean

    s = CStr(n)
    If Left$(s, 1) = "-" Then
        neg = True
        s = Mid$(s, 2)
    End If

    Dim out As String
    Dim i As Long
    Dim cnt As Long

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

Private Function FormatSvDecimal(ByVal v As Currency, ByVal decimals As Long) As String
    Dim fmt As String
    fmt = "0"
    If decimals > 0 Then fmt = fmt & "." & String$(decimals, "0")

    Dim s As String
    s = format$(v, fmt)
    s = Replace(s, ".", ",")
    FormatSvDecimal = s
End Function

Private Function FormatSvMoney(ByVal v As Currency) As String
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

Private Function ParseRateKr(ByVal s As String) As Currency
    ' Plocka sista numeriska värdet i cellen.
    ' Ex: "26,00 à 3256 kr" -> 3256
    Dim i As Long
    Dim token As String
    Dim current As String
    Dim sawDot As Boolean

    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)

        If ch >= "0" And ch <= "9" Then
            current = current & ch
        ElseIf (ch = "," Or ch = ".") And Len(current) > 0 And Not sawDot Then
            current = current & "."
            sawDot = True
        Else
            If Len(current) > 0 Then
                token = current
                current = ""
                sawDot = False
            End If
        End If
    Next i

    If Len(current) > 0 Then token = current

    If token = "" Or token = "." Then
        ParseRateKr = 0@
    Else
        ParseRateKr = CCur(Val(token))
    End If
End Function

