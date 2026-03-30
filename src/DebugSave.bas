Option Explicit

Public Sub SaveDebugCopy(ByVal srcDoc As Document)
    On Error GoTo Fail

    Dim debugPath As String
    debugPath = GetDebugCopyPath()

    If Len(debugPath) = 0 Then Exit Sub

    Dim tmpDoc As Document
    Set tmpDoc = Documents.Add(Visible:=False)

    ' Kopiera hela dokumentinnehållet som snapshot
    tmpDoc.Range.FormattedText = srcDoc.Range.FormattedText

    ' Kopiera grundläggande sidinställningar
    CopyPageSetup srcDoc, tmpDoc

    ' Ta bort tidigare debugfil om den finns
    On Error Resume Next
    Kill debugPath
    On Error GoTo Fail

    ' Spara alltid som docx för att hålla det enkelt och stabilt
    tmpDoc.SaveAs2 fileName:=debugPath, FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    tmpDoc.Close SaveChanges:=False
    Exit Sub

Fail:
    On Error Resume Next
    If Not tmpDoc Is Nothing Then
        tmpDoc.Close SaveChanges:=False
    End If

    MsgBox "Kunde inte spara debugkopia." & vbCrLf & vbCrLf & _
           "Mål: " & debugPath & vbCrLf & vbCrLf & _
           "Fel: " & Err.Description, vbExclamation, "KATS-Tools"
End Sub

Public Function GetDebugCopyPath() As String
    Dim folderPath As String

#If Mac Then
    ' På Mac i Word pekar HOME normalt på Words container.
    ' Det är bra här eftersom vi vill undvika cross-app access.
    folderPath = Environ$("HOME") & "/Documents"
#Else
    folderPath = Environ$("USERPROFILE") & "\Documents"
#End If

    EnsureFolderExists folderPath
    GetDebugCopyPath = JoinPath(folderPath, "KATS-Debug-Last.docx")
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
    On Error GoTo 0
End Sub

Private Sub CopyPageSetup(ByVal srcDoc As Document, ByVal dstDoc As Document)
    With dstDoc.PageSetup
        .TopMargin = srcDoc.PageSetup.TopMargin
        .BottomMargin = srcDoc.PageSetup.BottomMargin
        .LeftMargin = srcDoc.PageSetup.LeftMargin
        .RightMargin = srcDoc.PageSetup.RightMargin
        .HeaderDistance = srcDoc.PageSetup.HeaderDistance
        .FooterDistance = srcDoc.PageSetup.FooterDistance
        .Orientation = srcDoc.PageSetup.Orientation
        .PageWidth = srcDoc.PageSetup.PageWidth
        .PageHeight = srcDoc.PageSetup.PageHeight
    End With
End Sub
