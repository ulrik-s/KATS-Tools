Option Explicit

Public Sub CheckForUpdate()
#If Mac Then
    CheckForUpdateMac
#Else
    CheckForUpdateWindows
#End If
End Sub

Private Sub CheckForUpdateMac()
    On Error GoTo Fail

    Dim result As String
    result = AppleScriptTask("KATSUpdater.applescript", "checkAndInstallUpdate", "")

    HandleMacUpdateResult result
    Exit Sub

Fail:
    MsgBox "Kunde inte starta uppdateringskontrollen på Mac." & vbCrLf & vbCrLf & _
           "Fel: " & Err.Description, vbExclamation, "KATS-Tools"
End Sub

Private Sub HandleMacUpdateResult(ByVal result As String)
    Dim t As String
    t = Trim$(result)

    If t = "" Then
        MsgBox "Uppdateraren gav inget svar.", vbExclamation, "KATS-Tools"
        Exit Sub
    End If

    If StrComp(t, "UPTODATE", vbTextCompare) = 0 Then
        MsgBox "Du kör redan senaste versionen.", vbInformation, "KATS-Tools"
        Exit Sub
    End If

    If Left$(t, 10) = "INSTALLED|" Then
        MsgBox "Uppdatering installerad till version " & Mid$(t, 11) & "." & vbCrLf & _
               "Starta om Word.", vbInformation, "KATS-Tools"
        Exit Sub
    End If

    If Left$(t, 7) = "FAILED|" Then
        MsgBox "Uppdateringen misslyckades." & vbCrLf & vbCrLf & _
               Mid$(t, 8), vbExclamation, "KATS-Tools"
        Exit Sub
    End If

    MsgBox t, vbInformation, "KATS-Tools"
End Sub

Private Sub CheckForUpdateWindows()
    On Error GoTo Fail

    Dim updaterPath As String
    updaterPath = JoinPath(ThisDocument.path, "KATSUpdater.bat")

    If Len(Dir$(updaterPath)) = 0 Then
        MsgBox "KATSUpdater.bat hittades inte:" & vbCrLf & updaterPath, vbExclamation, "KATS-Tools"
        Exit Sub
    End If

    Dim checkResult As String
    checkResult = RunUpdaterCheckWindows(updaterPath)

    checkResult = Replace(checkResult, vbCr, "")
    checkResult = Replace(checkResult, vbLf, "")
    checkResult = Trim$(checkResult)

    If checkResult = "" Then
        MsgBox "Uppdateraren gav inget svar.", vbExclamation, "KATS-Tools"
        Exit Sub
    End If

    If StrComp(checkResult, "UPTODATE", vbTextCompare) = 0 Then
        MsgBox "Du har redan senaste versionen.", vbInformation, "KATS-Tools"
        Exit Sub
    End If

    If Left$(checkResult, 7) = "FAILED|" Then
        MsgBox "Uppdateringskontrollen misslyckades." & vbCrLf & vbCrLf & _
               Mid$(checkResult, 8), vbExclamation, "KATS-Tools"
        Exit Sub
    End If

    If Left$(checkResult, 7) = "UPDATE|" Then
        Dim newVersion As String
        newVersion = Mid$(checkResult, 8)

        Shell QuoteArg(updaterPath), vbNormalFocus

        MsgBox "Ny version finns: " & newVersion & "." & vbCrLf & vbCrLf & _
               "Uppdateraren har startats. Folj instruktionerna och starta om Word nar uppdateringen ar klar.", _
               vbInformation, "KATS-Tools"
        Exit Sub
    End If

    MsgBox "Okant svar fran uppdateraren:" & vbCrLf & checkResult, vbExclamation, "KATS-Tools"
    Exit Sub

Fail:
    MsgBox "Kunde inte starta uppdateraren pa Windows." & vbCrLf & vbCrLf & _
           "Fel: " & Err.Description, vbExclamation, "KATS-Tools"
End Sub

Private Function RunUpdaterCheckWindows(ByVal updaterPath As String) As String
    Dim wsh As Object
    Dim execObj As Object
    Dim cmd As String
    Dim stdoutText As String
    Dim stderrText As String

    Set wsh = CreateObject("WScript.Shell")

    cmd = "cmd.exe /d /c call " & QuoteArg(updaterPath) & " --checkonly"

    Set execObj = wsh.Exec(cmd)

    Do While execObj.Status = 0
        DoEvents
    Loop

    stdoutText = execObj.StdOut.ReadAll
    stderrText = execObj.StdErr.ReadAll

    stdoutText = NormalizeUpdaterResult(stdoutText)
    stderrText = NormalizeUpdaterResult(stderrText)

    If Len(stdoutText) > 0 Then
        RunUpdaterCheckWindows = stdoutText
    ElseIf Len(stderrText) > 0 Then
        RunUpdaterCheckWindows = "FAILED|" & stderrText
    Else
        RunUpdaterCheckWindows = ""
    End If
End Function

Private Function NormalizeUpdaterResult(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    NormalizeUpdaterResult = Trim$(s)
End Function

Private Function QuoteArg(ByVal s As String) As String
    QuoteArg = """" & Replace(s, """", """""") & """"
End Function
