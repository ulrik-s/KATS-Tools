Option Explicit

Public Sub RunKatsRibbon(control As IRibbonControl)
    KATS_ProcessAllTaggedBlocks
End Sub

Public Sub MailaTmpDokumentRibbon(control As IRibbonControl)
    MailaTmpDokument
End Sub

Public Sub CheckForUpdateRibbon(control As IRibbonControl)
    CheckForUpdate
End Sub

Public Sub KATS_ProcessAllTaggedBlocks()
    Dim doc As Document
    Set doc = ActiveDocument

    SaveDebugCopy ActiveDocument
    CleanupLeftoverPlaceholders ActiveDocument

    ResetProcessorState

    Dim oldTrack As Boolean
    oldTrack = doc.TrackRevisions
    doc.TrackRevisions = False

    KATS_ProcessKRTaggedBlocks doc

    ProcessTagEverywhere doc, "YTTRANDE_SIGNATUR", "Process_YTTRANDE_SIGNATUR"
    ProcessTagEverywhere doc, "YTTRANDE_PARTER", "Process_YTTRANDE_PARTER"

    doc.TrackRevisions = oldTrack
End Sub

Private Sub KATS_ProcessKRTaggedBlocks(ByVal doc As Document)
    Dim expensesContent As Range
    Dim expensesNoVatContent As Range
    Dim worklogContent As Range
    Dim arvodeContent As Range
    Dim arvodeTotalContent As Range
    Dim recipientContent As Range
    Dim signatureContent As Range

    Dim firstUtlagg As Range
    Dim secondUtlagg As Range

    Set firstUtlagg = ConsumeFirstTagRangeEverywhere(doc, "UTLAGGSSPECIFIKATION")
    Set secondUtlagg = ConsumeFirstTagRangeEverywhere(doc, "UTLAGGSSPECIFIKATION")

    If Not firstUtlagg Is Nothing Then
        Set expensesContent = firstUtlagg
    End If

    If Not secondUtlagg Is Nothing Then
        Set expensesNoVatContent = secondUtlagg
    End If

    Set worklogContent = ConsumeFirstTagRangeEverywhere(doc, "ARGRUPPERTIDERDATUMANTALSUMMA")
    Set arvodeContent = ConsumeFirstTagRangeEverywhere(doc, "ARVODE")
    Set arvodeTotalContent = ConsumeFirstTagRangeEverywhere(doc, "ARVODE_TOTAL")
    Set recipientContent = ConsumeFirstTagRangeEverywhere(doc, "MOTTAGARE")
    Set signatureContent = ConsumeFirstTagRangeEverywhere(doc, "SIGNATUR")

    KATS_ProcessKRPipelineRanges _
        expensesContent, _
        expensesNoVatContent, _
        worklogContent, _
        arvodeContent, _
        arvodeTotalContent, _
        recipientContent, _
        signatureContent
End Sub

Private Sub CleanupLeftoverPlaceholders(ByVal doc As Document)
    ReplaceAllLiteral doc.content, "[Utlägg2]", ""
    ReplaceAllLiteral doc.content, "[Utlägg]", ""
    ReplaceAllLiteral doc.content, "[UtläggAntal]", ""
    ReplaceAllLiteral doc.content, "[Utlägg2Antal]", ""
    ReplaceAllLiteral doc.content, "[Arvode]", ""
    ReplaceAllLiteral doc.content, "[Arvode2]", ""
    ReplaceAllLiteral doc.content, "[ArvodeAntal]", ""
    ReplaceAllLiteral doc.content, "[Arvode2Antal]", ""
    ReplaceAllLiteral doc.content, "[Tidspillan]", ""
    ReplaceAllLiteral doc.content, "[TidspillanAntal]", ""
    ReplaceAllLiteral doc.content, "[Tidspillan2]", ""
    ReplaceAllLiteral doc.content, "[Tidspillan2Antal]", ""
End Sub
