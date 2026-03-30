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
Private gHearingMinutes As Long

Private gRegexInitialized As Boolean
Private gRxHearingTime As RegexTy
Private gRxHearingTaxa As RegexTy

Private Type ArvodeTotalReadModel
    RowBeloppExkl As Long
    RowMoms As Long
    RowUtlaggEjMoms As Long
    RowBeloppInkl As Long
    ArvodeExMoms As Currency
    HasArvode As Boolean
    UtlaggEjMoms As Currency
    HasUtlaggEjMoms As Boolean
End Type

Private Type ArvodeTotalComputedModel
    RowBeloppExkl As Long
    RowMoms As Long
    RowUtlaggEjMoms As Long
    RowBeloppInkl As Long
    ArvodeExMoms As Currency
    Moms As Currency
    UtlaggEjMoms As Currency
    Incl As Currency
End Type

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

Public Sub SetCurrentPostort(ByVal value As String)
    gPostort = value
End Sub

Public Function GetCurrentPostort() As String
    GetCurrentPostort = gPostort
End Function

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

    For r = 1 To t.rows.count
        For c = 1 To t.Columns.count
            Dim s As String
            s = CellTextSafe(t, r, c)

            If Len(s) > 0 Then
                If IsTaxaForHearingText(s) Then
                    gIsTaxemal = True
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
    Dim t As Table
    Set t = ReadUtlaggTable(content)
    If t Is Nothing Then Exit Sub

    TransformUtlaggTable t
    RenderUtlaggTable t
End Sub

Private Function ReadUtlaggTable(ByVal content As Range) As Table
    Dim t As Table
    Set t = RequireSingleTable(content)
    If t Is Nothing Then Exit Function
    If t.Columns.count < 5 Then Exit Function

    Set ReadUtlaggTable = t
End Function

Private Sub TransformUtlaggTable(ByVal t As Table)
    ProcessExpenseSection t, "Utlägg", msUtlaggExMoms, True
    ProcessExpenseSection t, "Utlägg momsfri", msUtlaggEjMoms, False
End Sub

Private Sub RenderUtlaggTable(ByVal t As Table)
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
    Dim t As Table
    Set t = ReadArgrupperTable(content)
    If t Is Nothing Then Exit Sub

    TransformArgrupperTable t
    RenderArgrupperTable t
End Sub

Private Function ReadArgrupperTable(ByVal content As Range) As Table
    Set ReadArgrupperTable = RequireSingleTable(content)
End Function

Private Sub TransformArgrupperTable(ByVal t As Table)
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
End Sub

Private Sub RenderArgrupperTable(ByVal t As Table)
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

Public Sub Process_ARVODE(ByVal content As Range)
    Dim t As Table
    Set t = ReadArvodeTable(content)
    If t Is Nothing Then Exit Sub

    TransformArvodeState t
    RenderArvodeTable t
End Sub

Private Function ReadArvodeTable(ByVal content As Range) As Table
    Set ReadArvodeTable = RequireSingleTable(content)
End Function

Private Sub TransformArvodeState(ByVal t As Table)
    ' Arvode bygger på state fångad i tidigare steg (ARGRUPPER + UTLÄGG).
    ' Denna hook gör processorflödet explicit: Read -> Transform -> Render.
End Sub

Private Sub RenderArvodeTable(ByVal t As Table)
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
    Dim t As Table
    Set t = RequireSingleTable(content)
    If t Is Nothing Then Exit Sub

    Dim inputModel As ArvodeTotalReadModel
    inputModel = ReadArvodeTotal(t)
    If Not inputModel.HasArvode Then Exit Sub

    Dim computed As ArvodeTotalComputedModel
    computed = ComputeArvodeTotal(inputModel)

    RenderArvodeTotal t, computed
End Sub

Private Function ReadArvodeTotal(ByVal t As Table) As ArvodeTotalReadModel
    ReadArvodeTotal.RowBeloppExkl = FindTableRowContaining(t, "Belopp exkl. moms", 1)
    ReadArvodeTotal.RowMoms = FindTableRowContaining(t, "Moms (25%)", 1)
    ReadArvodeTotal.RowUtlaggEjMoms = FindTableRowContaining(t, "UTLÄGG EJ MOMS", 1)
    ReadArvodeTotal.RowBeloppInkl = FindTableRowContaining(t, "Belopp inkl. moms", 1)

    ReadArvodeTotal.HasArvode = HasMoneyState(msArvodeExMoms)
    If ReadArvodeTotal.HasArvode Then
        ReadArvodeTotal.ArvodeExMoms = GetMoneyState(msArvodeExMoms)
    End If

    ReadArvodeTotal.HasUtlaggEjMoms = HasMoneyState(msUtlaggEjMoms)
    If ReadArvodeTotal.HasUtlaggEjMoms Then
        ReadArvodeTotal.UtlaggEjMoms = GetMoneyState(msUtlaggEjMoms)
    End If
End Function

Private Function ComputeArvodeTotal(inputModel As ArvodeTotalReadModel) As ArvodeTotalComputedModel
    ComputeArvodeTotal.RowBeloppExkl = inputModel.RowBeloppExkl
    ComputeArvodeTotal.RowMoms = inputModel.RowMoms
    ComputeArvodeTotal.RowUtlaggEjMoms = inputModel.RowUtlaggEjMoms
    ComputeArvodeTotal.RowBeloppInkl = inputModel.RowBeloppInkl

    ComputeArvodeTotal.ArvodeExMoms = inputModel.ArvodeExMoms
    ComputeArvodeTotal.Moms = RoundCurrencyToDecimals(inputModel.ArvodeExMoms * 0.25@, 0)

    If inputModel.HasUtlaggEjMoms Then
        ComputeArvodeTotal.UtlaggEjMoms = RoundCurrencyToDecimals(inputModel.UtlaggEjMoms, 0)
    Else
        ComputeArvodeTotal.UtlaggEjMoms = 0@
    End If

    ComputeArvodeTotal.Incl = RoundCurrencyToDecimals( _
        ComputeArvodeTotal.ArvodeExMoms + ComputeArvodeTotal.Moms + ComputeArvodeTotal.UtlaggEjMoms, 0)
End Function

Private Sub RenderArvodeTotal(ByVal t As Table, computed As ArvodeTotalComputedModel)
    If computed.RowBeloppExkl > 0 Then
        CellSetTextSafe t, computed.RowBeloppExkl, 3, FormatSvMoney(computed.ArvodeExMoms)
    End If

    If computed.RowMoms > 0 Then
        CellSetTextSafe t, computed.RowMoms, 3, FormatSvMoney(computed.Moms)
    End If

    If computed.RowUtlaggEjMoms > 0 Then
        CellSetTextSafe t, computed.RowUtlaggEjMoms, 3, FormatSvMoney(computed.UtlaggEjMoms)
    End If

    If computed.RowBeloppInkl > 0 Then
        CellSetTextSafe t, computed.RowBeloppInkl, 3, FormatSvMoney(computed.Incl)
    End If

    If computed.RowUtlaggEjMoms > 0 Then
        DeleteArvodeRowIfZeroAmount t, computed.RowUtlaggEjMoms
    End If
End Sub

' ============================================================
' Generic table helpers
' ============================================================

Public Sub AutoFitUtlaggTable(ByVal t As Table)
    t.AllowAutoFit = True
    t.AutoFitBehavior wdAutoFitContent
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
