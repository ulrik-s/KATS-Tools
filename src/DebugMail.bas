Option Explicit

Public Sub DebugMailaTmpDokumentRibbon(control As IRibbonControl)
    MailaTmpDokument
End Sub

Public Sub MailaTmpDokument()
    Dim debugPath As String
    debugPath = GetDebugCopyPath()

    If Len(Dir$(debugPath)) = 0 Then
        MsgBox "Debugfilen hittades inte:" & vbCrLf & debugPath, vbExclamation, "KATS-Tools"
        Exit Sub
    End If

#If Mac Then
    MailaTmpDokumentMacOutlook debugPath
#Else
    MailaTmpDokumentWindowsOutlook debugPath
#End If
End Sub

Private Sub MailaTmpDokumentWindowsOutlook(ByVal filePath As String)
    On Error GoTo Fail

    Dim olApp As Object
    Dim olMail As Object

    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    With olMail
        .To = "ulrik.sjolin@gmail.com"
        .Subject = "KATS debugfil"
        .Body = "Hej," & vbCrLf & vbCrLf & _
                "Här kommer debugfilen från KATS-Tools." & vbCrLf & vbCrLf
        .Attachments.Add filePath
        .Display
    End With

    Exit Sub

Fail:
    MsgBox "Kunde inte skapa Outlook-mail." & vbCrLf & vbCrLf & _
           "Fel: " & Err.Description, vbExclamation, "KATS-Tools"
End Sub

Private Sub MailaTmpDokumentMacOutlook(ByVal filePath As String)
    On Error GoTo Fallback

    Dim payload As String
    payload = "ulrik.sjolin@gmail.com|||" & _
              "KATS debugfil|||" & _
              "Hej," & vbLf & vbLf & _
              "Här kommer debugfilen från KATS-Tools." & vbLf & vbLf & _
              "|||" & filePath

    Dim result As String
    result = AppleScriptTask("KATSMail.applescript", "createDebugMailInOutlook", payload)
    Exit Sub

Fallback:
    MailaTmpDokumentMacFallback filePath
End Sub

Private Sub MailaTmpDokumentMacFallback(ByVal filePath As String)
    On Error Resume Next

    ThisDocument.FollowHyperlink _
        Address:="mailto:ulrik.sjolin@gmail.com?subject=KATS%20debugfil&body=Hej,%0D%0A%0D%0AH%C3%A4r%20kommer%20debugfilen%20fr%C3%A5n%20KATS-Tools.%0D%0A"

#If Mac Then
    AppleScriptTask "KATSFileOps.applescript", "revealFileInFinder", filePath
#End If

    MsgBox "Kunde inte styra Outlook automatiskt på Mac." & vbCrLf & vbCrLf & _
           "Ett nytt mailutkast har öppnats." & vbCrLf & _
           "Dra in debugfilen i mailet från Finder-fönstret som öppnades.", _
           vbInformation, "KATS-Tools"
End Sub

