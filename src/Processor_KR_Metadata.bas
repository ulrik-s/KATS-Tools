Option Explicit

Private Type SignaturModel
    Namn As String
    Titel As String
    Postort As String
End Type

Private Type MottagareReadModel
    Raw As String
End Type

Private Type MottagareComputedModel
    FirstLine As String
    Postort As String
End Type

' ---- SIGNATUR ----
Public Sub Process_SIGNATUR(ByVal content As Range)
    Dim model As SignaturModel
    model = ReadSignatur()
    model = TransformSignatur(model)
    RenderSignatur content, model
End Sub

Private Function ReadSignatur() As SignaturModel
    ReadSignatur.Namn = GetFullName()
    ReadSignatur.Titel = GetTitle()
    ReadSignatur.Postort = GetCurrentPostort()
End Function

Private Function TransformSignatur(ByVal model As SignaturModel) As SignaturModel
    If Len(Trim$(model.Postort)) = 0 Then
        model.Postort = GetCity()
    End If
    TransformSignatur = model
End Function

Private Sub RenderSignatur(ByVal content As Range, ByVal model As SignaturModel)
    content.text = SwedishDateText(model.Postort) & vbCr & vbCr & model.Namn & vbCr & model.Titel
End Sub

' ---- MOTTAGARE ----
Public Sub Process_MOTTAGARE(ByVal content As Range)
    Dim t As Table
    Set t = ReadMottagareTable(content)
    If t Is Nothing Then Exit Sub

    Dim inputModel As MottagareReadModel
    inputModel = ReadMottagare(t)

    Dim computed As MottagareComputedModel
    computed = TransformMottagare(inputModel)
    If Len(computed.FirstLine) = 0 Then Exit Sub

    RenderMottagare t, computed
End Sub

Private Function ReadMottagareTable(ByVal content As Range) As Table
    Dim t As Table
    Set t = RequireSingleTable(content)
    If t Is Nothing Then Exit Function
    If t.rows.count <> 1 Or t.Columns.count < 2 Then Exit Function
    Set ReadMottagareTable = t
End Function

Private Function ReadMottagare(ByVal t As Table) As MottagareReadModel
    ReadMottagare.Raw = CellTextSafe(t, 1, 2)
    If Len(ReadMottagare.Raw) = 0 Then Exit Function

    ReadMottagare.Raw = Replace(ReadMottagare.Raw, vbCrLf, vbCr)
    ReadMottagare.Raw = Replace(ReadMottagare.Raw, Chr(11), vbCr)
    ReadMottagare.Raw = Replace(ReadMottagare.Raw, Chr(7), "")
End Function

Private Function TransformMottagare(ByVal inputModel As MottagareReadModel) As MottagareComputedModel
    Dim lines() As String
    Dim out As MottagareComputedModel
    On Error GoTo NoLines
    lines = Split(inputModel.Raw, vbCr)
    out.FirstLine = FirstNonEmptyLine(lines)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim postort As String
        If TryExtractPostort(lines(i), postort) Then
            out.Postort = postort
            Exit For
        End If
    Next i

    TransformMottagare = out
    Exit Function

NoLines:
    TransformMottagare.FirstLine = ""
End Function

Private Sub RenderMottagare(ByVal t As Table, ByVal computed As MottagareComputedModel)
    If Len(computed.Postort) > 0 Then
        SetCurrentPostort computed.Postort
    End If
    CellSetTextSafe t, 1, 2, computed.FirstLine & vbCr & "via e-post"
End Sub
