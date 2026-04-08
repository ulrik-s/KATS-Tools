Option Explicit

Private gRegexInitialized As Boolean
Private gRxHearingTime As RegexTy
Private gRxHearingTaxa As RegexTy

' ============================================================
' Public entry point for the KR pipeline
' ============================================================

Public Sub KATS_ProcessKRPipelineRanges( _
    ByVal expensesContent As Range, _
    ByVal expensesNoVatContent As Range, _
    ByVal worklogContent As Range, _
    ByVal arvodeContent As Range, _
    ByVal arvodeTotalContent As Range, _
    ByVal recipientContent As Range, _
    ByVal signatureContent As Range)

    Dim model As KATSDocumentModel
    Dim state As KATSComputedState

    KATS_ReadDocumentFromRanges _
        model, _
        expensesContent, _
        expensesNoVatContent, _
        worklogContent, _
        arvodeContent, _
        recipientContent, _
        signatureContent

    KATS_TransformDocument model, state

    KATS_RenderDocument _
        model, _
        state, _
        expensesContent, _
        expensesNoVatContent, _
        worklogContent, _
        arvodeContent, _
        arvodeTotalContent, _
        recipientContent, _
        signatureContent
End Sub

' ============================================================
' READ
' ============================================================

Private Sub KATS_ReadDocumentFromRanges( _
    ByRef model As KATSDocumentModel, _
    ByVal expensesContent As Range, _
    ByVal expensesNoVatContent As Range, _
    ByVal worklogContent As Range, _
    ByVal arvodeContent As Range, _
    ByVal recipientContent As Range, _
    ByVal signatureContent As Range)

    KATS_ReadExpenseBlock expensesContent, "UTLAGGSSPECIFIKATION", "Utlägg", True, model.Expenses
    KATS_ReadExpenseBlock expensesNoVatContent, "UTLAGGSSPECIFIKATION", "Utlägg momsfri", False, model.ExpensesNoVat
    KATS_ReadWorklogSection worklogContent, model.Worklog
    KATS_ReadArvodeSection arvodeContent, model.Arvode
    KATS_ReadRecipientSection recipientContent, model.Recipient
    KATS_ReadSignatureSection signatureContent, model.Signature
End Sub

Private Sub KATS_ReadExpenseBlock(ByVal content As Range, ByVal tagName As String, ByVal heading As String, ByVal applyMileageRule As Boolean, ByRef block As KATSExpenseBlockInput)
    Const COL_DATE As Long = 1
    Const COL_DESC As Long = 2
    Const COL_QTY As Long = 3
    Const COL_RATE As Long = 4
    Const COL_AMT As Long = 5

    block.Section.Kind = IIf(applyMileageRule, katsSectionExpenses, katsSectionExpensesNoVat)
    block.Section.TagName = tagName
    block.Heading = heading
    block.ApplyMileageRule = applyMileageRule

    If content Is Nothing Then Exit Sub
    block.Section.Exists = True
    If content.Tables.Count = 0 Then Exit Sub
    block.Section.HasTable = True

    Dim t As Table
    Set t = content.Tables(1)
    If t.Columns.Count < 5 Then Exit Sub

    Dim headingRow As Long
    Dim summaryRow As Long
    Dim r As Long

    headingRow = FindSectionHeadingRow(t, heading)
    summaryRow = FindSectionSummaryRowAfterHeading(t, heading)

    If headingRow = 0 Or summaryRow = 0 Then Exit Sub
    block.SummaryRow = summaryRow

    For r = headingRow + 1 To summaryRow - 1
        If LooksLikeIsoDate(CellTextSafe(t, r, COL_DATE)) Then
            block.Count = block.Count + 1
            If block.Count > 200 Then Exit For

            block.Rows(block.Count).RowIndex = r
            block.Rows(block.Count).Desc = CellTextSafe(t, r, COL_DESC)
            block.Rows(block.Count).Qty = SvToCurrency(CellTextSafe(t, r, COL_QTY))
            block.Rows(block.Count).Rate = SvToCurrency(CellTextSafe(t, r, COL_RATE))
            block.Rows(block.Count).ExistingAmount = SvToCurrency(CellTextSafe(t, r, COL_AMT))
        End If
    Next r
End Sub

Private Sub KATS_ReadWorklogSection(ByVal content As Range, ByRef worklog As KATSWorklogInput)
    worklog.Section.Kind = katsSectionWorklog
    worklog.Section.TagName = "ARGRUPPERTIDERDATUMANTALSUMMA"

    If content Is Nothing Then Exit Sub
    worklog.Section.Exists = True
    If content.Tables.Count = 0 Then Exit Sub
    worklog.Section.HasTable = True

    Dim t As Table
    Set t = content.Tables(1)

    KATS_ReadWorklogCategory t, "Arvode", katsArvode, worklog
    KATS_ReadWorklogCategory t, "Arvode helg", katsArvodeHelg, worklog
    KATS_ReadWorklogCategory t, "Tidsspillan", katsTidsspillan, worklog
    KATS_ReadWorklogCategory t, "Tidsspillan övrig tid", katsTidsspillanOvrigTid, worklog

    Dim r As Long
    Dim c As Long

    For r = 1 To t.Rows.Count
        For c = 1 To t.Columns.Count
            Dim s As String
            s = CellTextSafe(t, r, c)

            If Len(s) > 0 Then
                Dim hh As Long
                Dim mm As Long
                Dim isTaxa As Boolean

                If KR_TryMatchHearing(s, hh, mm, isTaxa) Then
                    Dim baseDate As Date
                    baseDate = Date

                    Dim ds As String
                    ds = CellTextSafe(t, r, 1)
                    If LooksLikeIsoDate(ds) Then baseDate = IsoDateToDate(ds)

                    worklog.HearingFound = True
                    worklog.HearingRow = r
                    worklog.HearingStart = DateSerial(Year(baseDate), Month(baseDate), Day(baseDate)) + TimeSerial(hh, mm, 0)
                    worklog.HearingMinutes = DateDiff("n", worklog.HearingStart, Now)
                    If worklog.HearingMinutes < 0 Then worklog.HearingMinutes = worklog.HearingMinutes + (24& * 60&)
                    worklog.IsTaxa = isTaxa
                    Exit Sub
                End If
            End If
        Next c
    Next r
End Sub

Private Sub KATS_ReadWorklogCategory(ByVal t As Table, ByVal heading As String, ByVal category As KATSArCategory, ByRef worklog As KATSWorklogInput)
    Const COL_DATE As Long = 1
    Const COL_DESC As Long = 2
    Const COL_QTY As Long = 3
    Const COL_AMT As Long = 5

    Dim headingRow As Long
    Dim summaryRow As Long
    Dim r As Long

    headingRow = FindSectionHeadingRow(t, heading)
    summaryRow = FindSectionSummaryRowAfterHeading(t, heading)

    If headingRow = 0 Or summaryRow = 0 Then Exit Sub

    worklog.SummaryRow(category) = summaryRow

    For r = headingRow + 1 To summaryRow - 1
        If LooksLikeIsoDate(CellTextSafe(t, r, COL_DATE)) Then
            Dim desc As String
            desc = CellTextSafe(t, r, COL_DESC)

            If KR_IsAdjustmentDesc(desc) Then
                worklog.CategoryAdjustment(category) = worklog.CategoryAdjustment(category) + RoundCurrencyToDecimals(SvToCurrency(CellTextSafe(t, r, COL_AMT)), 2)
            Else
                worklog.CategoryHours(category) = worklog.CategoryHours(category) + SvToCurrency(CellTextSafe(t, r, COL_QTY))
            End If
        End If
    Next r

    worklog.CategoryHours(category) = RoundCurrencyToDecimals(worklog.CategoryHours(category), 2)
    worklog.CategoryAdjustment(category) = RoundCurrencyToDecimals(worklog.CategoryAdjustment(category), 2)
End Sub

Private Sub KATS_ReadArvodeSection(ByVal content As Range, ByRef arvode As KATSArvodeInput)
    arvode.Section.Kind = katsSectionArvode
    arvode.Section.TagName = "ARVODE"

    If content Is Nothing Then Exit Sub
    arvode.Section.Exists = True
    If content.Tables.Count = 0 Then Exit Sub
    arvode.Section.HasTable = True

    Dim t As Table
    Set t = content.Tables(1)
    If t.Rows.Count < 6 Or t.Columns.Count < 3 Then Exit Sub

    arvode.RateText(2) = CellTextSafe(t, 2, 2)
    arvode.RateText(3) = CellTextSafe(t, 3, 2)
    arvode.RateText(4) = CellTextSafe(t, 4, 2)
    arvode.RateText(5) = CellTextSafe(t, 5, 2)
    arvode.RateText(6) = CellTextSafe(t, 6, 2)
End Sub

Private Sub KATS_ReadRecipientSection(ByVal content As Range, ByRef recipient As KATSRecipientInput)
    recipient.Section.Kind = katsSectionRecipient
    recipient.Section.TagName = "MOTTAGARE"

    If content Is Nothing Then Exit Sub
    recipient.Section.Exists = True
    If content.Tables.Count = 0 Then Exit Sub
    recipient.Section.HasTable = True

    Dim t As Table
    Set t = content.Tables(1)
    If t.Rows.Count <> 1 Or t.Columns.Count < 2 Then Exit Sub

    recipient.RawText = t.Cell(1, 2).Range.Text
End Sub

Private Sub KATS_ReadSignatureSection(ByVal content As Range, ByRef signature As KATSSignatureInput)
    signature.Section.Kind = katsSectionSignature
    signature.Section.TagName = "SIGNATUR"

    If content Is Nothing Then Exit Sub
    signature.Section.Exists = True
End Sub

' ============================================================
' TRANSFORM
' ============================================================

Private Sub KATS_TransformDocument(ByRef model As KATSDocumentModel, ByRef state As KATSComputedState)
    KATS_TransformWorklog model.Worklog, state
    KATS_TransformExpenseBlock model.Expenses, state.ExpenseOutput, state.ExpenseExVatTotal
    KATS_TransformExpenseBlock model.ExpensesNoVat, state.ExpenseNoVatOutput, state.ExpenseNoVatTotal
    KATS_TransformRecipient model.Recipient, state
    KATS_TransformArvode model.Arvode, state
End Sub

Private Sub KATS_TransformWorklog(ByRef worklog As KATSWorklogInput, ByRef state As KATSComputedState)
    Dim i As Long

    For i = 1 To 4
        state.CategoryHours(i) = worklog.CategoryHours(i)
        state.CategoryAdjustment(i) = worklog.CategoryAdjustment(i)
    Next i

    state.HearingFound = worklog.HearingFound
    state.HearingStart = worklog.HearingStart
    state.HearingMinutes = worklog.HearingMinutes
    state.IsTaxa = worklog.IsTaxa
End Sub

Private Sub KATS_TransformExpenseBlock(ByRef block As KATSExpenseBlockInput, ByRef output As KATSExpenseBlockOutput, ByRef totalOut As Currency)
    Dim i As Long
    Dim amt As Currency
    Dim rate As Currency

    output.Count = block.Count
    output.Total = 0@

    For i = 1 To block.Count
        output.RowIndex(i) = block.Rows(i).RowIndex

        If KR_IsAdjustmentDesc(block.Rows(i).Desc) Then
            amt = RoundCurrencyToDecimals(block.Rows(i).ExistingAmount, 2)
        Else
            rate = block.Rows(i).Rate

            If block.ApplyMileageRule Then
                If InStr(1, block.Rows(i).Desc, "Milersättning", vbTextCompare) > 0 Then
                    rate = GetMileage()
                End If
            End If

            If block.Rows(i).Qty <> 0@ And rate <> 0@ Then
                amt = CCur(RoundHalfAwayFromZeroToLong(block.Rows(i).Qty * rate))
            Else
                amt = CCur(RoundHalfAwayFromZeroToLong(block.Rows(i).ExistingAmount))
            End If
        End If

        output.Amount(i) = amt
        output.Total = output.Total + amt
    Next i

    output.Total = RoundCurrencyToDecimals(output.Total, 2)
    totalOut = output.Total
End Sub

Private Sub KATS_TransformRecipient(ByRef recipient As KATSRecipientInput, ByRef state As KATSComputedState)
    Dim raw As String
    raw = recipient.RawText
    If Len(raw) = 0 Then Exit Sub

    raw = Replace(raw, vbCrLf, vbCr)
    raw = Replace(raw, Chr(11), vbCr)
    raw = Replace(raw, Chr(7), "")

    Dim lines() As String
    lines = Split(raw, vbCr)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim postort As String
        If TryExtractPostort(lines(i), postort) Then
            state.Postort = postort
            Exit For
        End If
    Next i
End Sub

Private Sub KATS_TransformArvode(ByRef arvode As KATSArvodeInput, ByRef state As KATSComputedState)
    Const ROW_ARVODE As Long = 2
    Const ROW_ARVODE_HELG As Long = 3
    Const ROW_TIDSSPILLAN As Long = 4
    Const ROW_TIDSSPILLAN_OVRIG As Long = 5
    Const ROW_UTLAGG As Long = 6

    Dim i As Long
    For i = 1 To 6
        state.ArvodeSpec(i) = ""
        state.ArvodeAmount(i) = 0@
        state.ArvodeKeepRow(i) = False
    Next i

    state.ArvodeAmount(ROW_UTLAGG) = state.ExpenseExVatTotal
    state.ArvodeKeepRow(ROW_UTLAGG) = (state.ExpenseExVatTotal <> 0@)

    If KATS_ShouldUseTaxa(state) Then
        KATS_TransformArvodeTaxa arvode, state
    Else
        KATS_TransformArvodeNormal arvode, state
    End If

    state.ArvodeExMoms = 0@
    For i = ROW_ARVODE To ROW_UTLAGG
        If state.ArvodeKeepRow(i) Then
            state.ArvodeExMoms = state.ArvodeExMoms + state.ArvodeAmount(i)
        End If
    Next i
    state.ArvodeExMoms = RoundCurrencyToDecimals(state.ArvodeExMoms, 2)
End Sub

Private Sub KATS_TransformArvodeNormal(ByRef arvode As KATSArvodeInput, ByRef state As KATSComputedState)
    KATS_BuildHoursRow arvode.RateText(2), state.CategoryHours(katsArvode), state.CategoryAdjustment(katsArvode), state.ArvodeSpec(2), state.ArvodeAmount(2), state.ArvodeKeepRow(2)
    KATS_BuildHoursRow arvode.RateText(3), state.CategoryHours(katsArvodeHelg), state.CategoryAdjustment(katsArvodeHelg), state.ArvodeSpec(3), state.ArvodeAmount(3), state.ArvodeKeepRow(3)
    KATS_BuildHoursRow arvode.RateText(4), state.CategoryHours(katsTidsspillan), state.CategoryAdjustment(katsTidsspillan), state.ArvodeSpec(4), state.ArvodeAmount(4), state.ArvodeKeepRow(4)
    KATS_BuildHoursRow arvode.RateText(5), state.CategoryHours(katsTidsspillanOvrigTid), state.CategoryAdjustment(katsTidsspillanOvrigTid), state.ArvodeSpec(5), state.ArvodeAmount(5), state.ArvodeKeepRow(5)
End Sub

Private Sub KATS_TransformArvodeTaxa(ByRef arvode As KATSArvodeInput, ByRef state As KATSComputedState)
    Dim taxAmount As Currency
    taxAmount = KATS_GetTaxAmountLevel1(state.HearingMinutes)

    state.ArvodeSpec(2) = KATS_FormatTaxaSpec(state.HearingMinutes, state.CategoryAdjustment(katsArvode))
    state.ArvodeAmount(2) = RoundCurrencyToDecimals(taxAmount + state.CategoryAdjustment(katsArvode), 2)
    state.ArvodeKeepRow(2) = (state.ArvodeAmount(2) <> 0@)

    KATS_BuildHoursRow arvode.RateText(3), state.CategoryHours(katsArvodeHelg), state.CategoryAdjustment(katsArvodeHelg), state.ArvodeSpec(3), state.ArvodeAmount(3), state.ArvodeKeepRow(3)

    Dim overtimeHours As Currency
    overtimeHours = state.CategoryHours(katsTidsspillan) - 1@
    If overtimeHours < 0@ Then overtimeHours = 0@

    KATS_BuildHoursRow arvode.RateText(4), overtimeHours, state.CategoryAdjustment(katsTidsspillan), state.ArvodeSpec(4), state.ArvodeAmount(4), state.ArvodeKeepRow(4)
    KATS_BuildHoursRow arvode.RateText(5), state.CategoryHours(katsTidsspillanOvrigTid), state.CategoryAdjustment(katsTidsspillanOvrigTid), state.ArvodeSpec(5), state.ArvodeAmount(5), state.ArvodeKeepRow(5)
End Sub

Private Sub KATS_BuildHoursRow(ByVal rateText As String, ByVal hours As Currency, ByVal adjustmentAmount As Currency, ByRef specOut As String, ByRef amountOut As Currency, ByRef keepOut As Boolean)
    Dim displayHours As Currency
    displayHours = RoundCurrencyToDecimals(hours, 2)

    If displayHours = 0@ And adjustmentAmount = 0@ Then
        specOut = ""
        amountOut = 0@
        keepOut = False
        Exit Sub
    End If

    If displayHours <> 0@ Then
        Dim rate As Currency
        rate = ParseRateKr(rateText)
        If rate = 0@ Then
            specOut = ""
            amountOut = 0@
            keepOut = False
            Exit Sub
        End If

        amountOut = RoundCurrencyToDecimals((displayHours * rate) + adjustmentAmount, 2)
        specOut = FormatSvDecimal(displayHours, 2) & " à " & FormatSvInt(CLng(rate)) & " kr"

        If adjustmentAmount <> 0@ Then
            specOut = specOut & " + justering " & FormatSvMoney(adjustmentAmount)
        End If
    Else
        specOut = "Justering"
        amountOut = RoundCurrencyToDecimals(adjustmentAmount, 2)
    End If

    keepOut = (amountOut <> 0@ Or Len(specOut) > 0)
End Sub

Private Function KATS_ShouldUseTaxa(ByRef state As KATSComputedState) As Boolean
    KATS_ShouldUseTaxa = (state.IsTaxa And state.HearingFound And state.HearingMinutes <= 225)
End Function

Private Function KATS_GetTaxAmountLevel1(ByVal hearingMinutes As Long) As Currency
    Select Case hearingMinutes
        Case 0 To 14: KATS_GetTaxAmountLevel1 = 2809@
        Case 15 To 29: KATS_GetTaxAmountLevel1 = 2980@
        Case 30 To 44: KATS_GetTaxAmountLevel1 = 3509@
        Case 45 To 59: KATS_GetTaxAmountLevel1 = 4049@
        Case 60 To 74: KATS_GetTaxAmountLevel1 = 4583@
        Case 75 To 89: KATS_GetTaxAmountLevel1 = 5106@
        Case 90 To 104: KATS_GetTaxAmountLevel1 = 5635@
        Case 105 To 119: KATS_GetTaxAmountLevel1 = 6164@
        Case 120 To 134: KATS_GetTaxAmountLevel1 = 6704@
        Case 135 To 149: KATS_GetTaxAmountLevel1 = 7227@
        Case 150 To 164: KATS_GetTaxAmountLevel1 = 7767@
        Case 165 To 179: KATS_GetTaxAmountLevel1 = 8301@
        Case 180 To 194: KATS_GetTaxAmountLevel1 = 8824@
        Case 195 To 209: KATS_GetTaxAmountLevel1 = 9364@
        Case 210 To 225: KATS_GetTaxAmountLevel1 = 9887@
        Case Else: KATS_GetTaxAmountLevel1 = 0@
    End Select
End Function

Private Function KATS_FormatTaxaSpec(ByVal totalMinutes As Long, ByVal adjustmentAmount As Currency) As String
    Dim wholeHours As Long
    Dim remMinutes As Long
    Dim s As String

    wholeHours = 0
    If totalMinutes >= 60 Then wholeHours = Int(totalMinutes / 60)
    remMinutes = totalMinutes - (wholeHours * 60)

    s = ""
    If wholeHours > 0 Then s = CStr(wholeHours) & " tim "
    s = s & CStr(remMinutes) & " min enligt taxa"

    If adjustmentAmount <> 0@ Then
        s = s & " + justering " & FormatSvMoney(adjustmentAmount)
    End If

    KATS_FormatTaxaSpec = s
End Function

' ============================================================
' RENDER
' ============================================================

Private Sub KATS_RenderDocument( _
    ByRef model As KATSDocumentModel, _
    ByRef state As KATSComputedState, _
    ByVal expensesContent As Range, _
    ByVal expensesNoVatContent As Range, _
    ByVal worklogContent As Range, _
    ByVal arvodeContent As Range, _
    ByVal arvodeTotalContent As Range, _
    ByVal recipientContent As Range, _
    ByVal signatureContent As Range)

    KATS_RenderExpenseBlock expensesContent, model.Expenses, state.ExpenseOutput
    KATS_RenderExpenseBlock expensesNoVatContent, model.ExpensesNoVat, state.ExpenseNoVatOutput
    KATS_RenderWorklogSection worklogContent, model.Worklog, state
    KATS_RenderArvodeSection arvodeContent, state
    KATS_RenderArvodeTotalSection arvodeTotalContent, state
    KATS_RenderRecipientSection recipientContent, state
    KATS_RenderSignatureSection signatureContent, state
End Sub

Private Sub KATS_RenderExpenseBlock(ByVal content As Range, ByRef block As KATSExpenseBlockInput, ByRef output As KATSExpenseBlockOutput)
    Const COL_QTY As Long = 3
    Const COL_AMT As Long = 5
    Const COL_RATE As Long = 4

    If content Is Nothing Then Exit Sub
    If content.Tables.Count = 0 Then Exit Sub

    Dim t As Table
    Set t = content.Tables(1)

    Dim i As Long
    For i = 1 To output.Count
        Dim r As Long
        r = output.RowIndex(i)

        If block.ApplyMileageRule Then
            If InStr(1, block.Rows(i).Desc, "Milersättning", vbTextCompare) > 0 Then
                CellSetTextSafe t, r, COL_RATE, FormatSvDecimal(GetMileage(), 2)
            End If
        End If

        CellSetTextSafe t, r, COL_AMT, KATS_FormatTableAmount(output.Amount(i))
    Next i

    If block.SummaryRow > 0 Then
        CellSetTextSafe t, block.SummaryRow, COL_QTY, ""
        CellSetTextSafe t, block.SummaryRow, COL_AMT, KATS_FormatTableAmount(output.Total)
    End If

    AutoFitUtlaggTable t
    AddAirBeforeSectionHeadings t, 12
End Sub

Private Sub KATS_RenderWorklogSection(ByVal content As Range, ByRef worklog As KATSWorklogInput, ByRef state As KATSComputedState)
    If content Is Nothing Then Exit Sub
    If content.Tables.Count = 0 Then Exit Sub

    Dim t As Table
    Set t = content.Tables(1)

    Dim i As Long
    For i = 1 To 4
        If worklog.SummaryRow(i) > 0 Then
            CellSetTextSafe t, worklog.SummaryRow(i), 3, FormatSvDecimal(state.CategoryHours(i), 2)
        End If
    Next i

    If worklog.HearingFound And worklog.HearingRow > 0 Then
        Dim hearingHours As Currency
        hearingHours = RoundCurrencyToDecimals(CCur(state.HearingMinutes) / 60@, 2)
        CellSetTextSafe t, worklog.HearingRow, 3, FormatSvDecimal(hearingHours, 2)
    End If

    Dim totalRow As Long
    totalRow = FindTableRowContaining(t, "Ärende, total", 1)
    If totalRow > 0 Then
        CellSetTextSafe t, totalRow, 1, ""
        CellSetTextSafe t, totalRow, 2, ""
        CellSetTextSafe t, totalRow, 3, ""
    End If

    AddAirBeforeSectionHeadings t, 12
End Sub

Private Sub KATS_RenderArvodeSection(ByVal content As Range, ByRef state As KATSComputedState)
    Const ROW_ARVODE As Long = 2
    Const ROW_ARVODE_HELG As Long = 3
    Const ROW_TIDSSPILLAN As Long = 4
    Const ROW_TIDSSPILLAN_OVRIG As Long = 5
    Const ROW_UTLAGG As Long = 6

    If content Is Nothing Then Exit Sub
    If content.Tables.Count = 0 Then Exit Sub

    Dim t As Table
    Set t = content.Tables(1)
    If t.Rows.Count < 6 Or t.Columns.Count < 3 Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = ROW_ARVODE To ROW_UTLAGG
        If state.ArvodeKeepRow(rowIndex) Then
            CellSetTextSafe t, rowIndex, 2, state.ArvodeSpec(rowIndex)
            CellSetTextSafe t, rowIndex, 3, FormatSvMoney(state.ArvodeAmount(rowIndex))
        Else
            CellSetTextSafe t, rowIndex, 3, ""
        End If
    Next rowIndex

    DeleteArvodeRowIfZeroAmount t, ROW_UTLAGG
    DeleteArvodeRowIfZeroAmount t, ROW_TIDSSPILLAN_OVRIG
    DeleteArvodeRowIfZeroAmount t, ROW_TIDSSPILLAN
    DeleteArvodeRowIfZeroAmount t, ROW_ARVODE_HELG
    DeleteArvodeRowIfZeroAmount t, ROW_ARVODE
End Sub

Private Sub KATS_RenderArvodeTotalSection(ByVal content As Range, ByRef state As KATSComputedState)
    If content Is Nothing Then Exit Sub
    If content.Tables.Count = 0 Then Exit Sub

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

    If rowBeloppExkl > 0 Then
        CellSetTextSafe t, rowBeloppExkl, 3, FormatSvMoney(state.ArvodeExMoms)
    End If

    Dim moms As Currency
    moms = RoundCurrencyToDecimals(state.ArvodeExMoms * 0.25@, 2)
    If rowMoms > 0 Then
        CellSetTextSafe t, rowMoms, 3, FormatSvMoney(moms)
    End If

    If rowUtlaggEjMoms > 0 Then
        CellSetTextSafe t, rowUtlaggEjMoms, 3, FormatSvMoney(state.ExpenseNoVatTotal)
    End If

    Dim incl As Currency
    incl = RoundCurrencyToDecimals(state.ArvodeExMoms + moms + state.ExpenseNoVatTotal, 0)
    If rowBeloppInkl > 0 Then
        CellSetTextSafe t, rowBeloppInkl, 3, FormatSvMoney(incl)
    End If

    If rowUtlaggEjMoms > 0 Then
        DeleteArvodeRowIfZeroAmount t, rowUtlaggEjMoms
    End If
End Sub

Private Sub KATS_RenderRecipientSection(ByVal content As Range, ByRef state As KATSComputedState)
    If content Is Nothing Then Exit Sub
    If content.Tables.Count = 0 Then Exit Sub

    Dim t As Table
    Set t = content.Tables(1)
    If t.Rows.Count <> 1 Or t.Columns.Count < 2 Then Exit Sub

    Dim raw As String
    raw = t.Cell(1, 2).Range.Text
    raw = Replace(raw, Chr(7), "")
    raw = Replace(raw, Chr(11), vbCr)
    raw = Replace(raw, vbCrLf, vbCr)

    Dim lines() As String
    lines = Split(raw, vbCr)

    Dim firstLine As String
    firstLine = FirstNonEmptyLine(lines)
    If Len(firstLine) = 0 Then Exit Sub

    CellSetTextSafe t, 1, 2, firstLine & vbCr & "via e-post"
End Sub

Private Sub KATS_RenderSignatureSection(ByVal content As Range, ByRef state As KATSComputedState)
    If content Is Nothing Then Exit Sub

    content.Text = SwedishDateText(FirstNonEmptyString(state.Postort, GetCity(), "Lund")) & vbCr & vbCr & GetFullName() & vbCr & GetTitle()
End Sub

Private Function KATS_FormatTableAmount(ByVal v As Currency) As String
    v = RoundCurrencyToDecimals(v, 2)

    If v = RoundCurrencyToDecimals(v, 0) Then
        KATS_FormatTableAmount = FormatSvInt(RoundHalfAwayFromZeroToLong(v))
    Else
        KATS_FormatTableAmount = FormatSvDecimal(v, 2)
    End If
End Function

' ============================================================
' Shared KR helpers
' ============================================================

Private Sub KR_EnsureRegexInitialized()
    If gRegexInitialized Then Exit Sub

    InitializeRegex gRxHearingTime, _
        "medverkat vid (?:huvud)?f.rhandling fr.n\s*(?:kl\.?\s*)?([0-9]{1,2})(?:\s*[:.]\s*([0-9]{2}))?", _
        True

    InitializeRegex gRxHearingTaxa, _
        "medverkat vid (?:huvud)?f.rhandling fr.n\s*(?:kl\.?\s*)?[0-9]{1,2}(?:\s*[:.]\s*[0-9]{2})?\s*[,;:]?\s*enligt taxa\b", _
        True

    gRegexInitialized = True
End Sub

Private Function KR_TryMatchHearing(ByVal s As String, ByRef hh As Long, ByRef mm As Long, ByRef isTaxa As Boolean) As Boolean
    Dim matcher As MatcherStateTy
    Dim h As String
    Dim m As String

    hh = 0
    mm = 0
    isTaxa = False

    KR_EnsureRegexInitialized

    If Not Match(matcher, gRxHearingTime, s) Then Exit Function

    h = GetCapture(matcher, s, 1)
    m = GetCapture(matcher, s, 2)

    If Len(h) = 0 Then Exit Function

    hh = CLng(h)
    If Len(m) > 0 Then mm = CLng(m)

    If hh < 0 Or hh > 23 Then Exit Function
    If mm < 0 Or mm > 59 Then Exit Function

    isTaxa = Test(gRxHearingTaxa, s)
    KR_TryMatchHearing = True
End Function

Private Function KR_IsAdjustmentDesc(ByVal desc As String) As Boolean
    KR_IsAdjustmentDesc = (InStr(1, desc, "justering", vbTextCompare) > 0)
End Function

Private Function FindSectionHeadingRow(ByVal t As Table, ByVal heading As String) As Long
    Dim r As Long
    For r = 1 To t.Rows.Count
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
    For r = headingRow + 1 To t.Rows.Count
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

    For c = 1 To t.Rows(rowIndex).Cells.Count
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
        RowMatchesHeading = RegexEqualsLoose(firstValue, heading)
    End If
End Function

Private Function RegexEqualsLoose(ByVal text As String, ByVal needle As String) As Boolean
    Dim rx As RegexTy
    ' NOTE: Our custom regex engine does not reliably support \s,
    ' so we explicitly allow leading/trailing ASCII spaces.
    InitializeRegex rx, "^ *" & SwedishLooseRegex(needle) & " *$", True
    RegexEqualsLoose = Test(rx, text)
End Function

Private Sub DeleteArvodeRowIfZeroAmount(ByVal t As Table, ByVal rowIndex As Long)
    If rowIndex < 1 Or rowIndex > t.Rows.Count Then Exit Sub

    Dim amtText As String
    amtText = Trim$(CellTextSafe(t, rowIndex, 3))

    If Not HasAnyDigit(amtText) Then
        t.Rows(rowIndex).Delete
        Exit Sub
    End If

    If RoundCurrencyToDecimals(SvToCurrency(amtText), 2) = 0@ Then
        t.Rows(rowIndex).Delete
    End If
End Sub

Private Sub AddAirBeforeSectionHeadings(ByVal t As Table, Optional ByVal pointsBefore As Single = 8)
    Dim r As Long
    Dim c As Long
    Dim firstHeadingSeen As Boolean

    firstHeadingSeen = False

    For r = 1 To t.Rows.Count
        If (t.Rows(r).Cells.Count = 1) Then
            If firstHeadingSeen Then
                For c = 1 To t.Rows(r).Cells.Count
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

Private Sub AutoFitUtlaggTable(ByVal t As Table)
    t.AllowAutoFit = True
    t.AutoFitBehavior wdAutoFitContent
End Sub
