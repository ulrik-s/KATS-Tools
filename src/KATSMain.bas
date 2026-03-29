Option Explicit

Public Sub RunKatsRibbon(control As IRibbonControl)
    KATS_ProcessAllTaggedBlocks
End Sub

Public Sub MailaTmpDokumentRibbon(control As IRibbonControl)
    MailaTmpDokument
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

    ' Enda stället där ordning + tag + processor definieras
    Dim pipeline As Variant
    pipeline = Array( _
        Array("UTLAGGSSPECIFIKATION", "Process_UTLAGGSSPECIFIKATION"), _
        Array("ARGRUPPERTIDERDATUMANTALSUMMA", "Process_ARGRUPPERTIDERDATUMANTALSUMMA"), _
        Array("ARVODE", "Process_ARVODE"), _
        Array("ARVODE_TOTAL", "Process_ARVODE_TOTAL"), _
        Array("MOTTAGARE", "Process_MOTTAGARE"), _
        Array("SIGNATUR", "Process_SIGNATUR"), _
        Array("YTTRANDE_SIGNATUR", "Process_YTTRANDE_SIGNATUR"), _
        Array("YTTRANDE_PARTER", "Process_YTTRANDE_PARTER") _
    )

    Dim i As Long
    For i = LBound(pipeline) To UBound(pipeline)
        ProcessTagEverywhere doc, CStr(pipeline(i)(0)), CStr(pipeline(i)(1))
    Next i

    doc.TrackRevisions = oldTrack
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

Private Sub ReplaceAllLiteral(ByVal rng As Range, ByVal findText As String, ByVal replaceText As String)
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


Private Function GetFileExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")

    If p > 0 Then
        GetFileExtension = Mid$(fileName, p)
    Else
        GetFileExtension = ".docx"
    End If
End Function

