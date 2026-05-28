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

    KATS_ProcessTagsWithProcessors doc, BuildTagProcessorSpecs()

    doc.TrackRevisions = oldTrack
End Sub

Private Function BuildTagProcessorSpecs() As Variant
    BuildTagProcessorSpecs = Array( _
        Array("UTLAGGSSPECIFIKATION", "Process_UTLAGGSSPECIFIKATION"), _
        Array("ARGRUPPERTIDERDATUMANTALSUMMA", "Process_ARGRUPPERTIDERDATUMANTALSUMMA"), _
        Array("ARVODE", "Process_ARVODE"), _
        Array("ARVODE_TOTAL", "Process_ARVODE_TOTAL"), _
        Array("MOTTAGARE", "Process_MOTTAGARE"), _
        Array("SIGNATUR", "Process_SIGNATUR"), _
        Array("YTTRANDE_PARTER", "Process_YTTRANDE_PARTER"), _
        Array("YTTRANDE_SIGNATUR", "Process_YTTRANDE_SIGNATUR") _
    )
End Function

Private Sub KATS_ProcessTagsWithProcessors(ByVal doc As Document, ByVal processorSpecs As Variant)
    Dim i As Long

    For i = LBound(processorSpecs) To UBound(processorSpecs)
        ProcessTagEverywhere doc, CStr(processorSpecs(i)(0)), CStr(processorSpecs(i)(1))
    Next i
End Sub

Private Sub CleanupLeftoverPlaceholders(ByVal doc As Document)
    Dim placeholders As Variant
    placeholders = Array( _
        "[Utlägg2]", _
        "[Utlägg]", _
        "[UtläggAntal]", _
        "[Utlägg2Antal]", _
        "[Arvode]", _
        "[Arvode2]", _
        "[ArvodeAntal]", _
        "[Arvode2Antal]", _
        "[Tidspillan]", _
        "[TidspillanAntal]", _
        "[Tidspillan2]", _
        "[Tidspillan2Antal]" _
    )

    Dim i As Long
    For i = LBound(placeholders) To UBound(placeholders)
        ReplaceAllLiteral doc.content, CStr(placeholders(i)), ""
    Next i
End Sub
