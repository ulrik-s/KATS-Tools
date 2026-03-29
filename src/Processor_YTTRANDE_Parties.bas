Option Explicit

Private Type YttrandeSignaturModel
    Namn As String
    Titel As String
    Postort As String
End Type

Private Type YttrandeParterReadModel
    Raw As String
End Type

Private Type YttrandeParterComputedModel
    LeftParty As String
    RightParty As String
    Names As Variant
    IsValid As Boolean
End Type

Public Function ExtractPersonNames(ByVal fullName As String, ByVal textBlock As String) As Variant
    Dim lines() As String
    Dim line As Variant
    Dim result() As String
    Dim count As Long

    count = -1

    ' Lägg alltid till första parametern först
    If Len(Trim$(fullName)) > 0 Then
        count = count + 1
        ReDim Preserve result(0 To count)
        result(count) = Trim$(fullName)
    End If

    textBlock = Replace(textBlock, vbCrLf, vbLf)
    textBlock = Replace(textBlock, vbCr, vbLf)
    lines = Split(textBlock, vbLf)

    For Each line In lines
        Dim s As String
        s = Trim$(CStr(line))

        If Len(s) > 0 Then
            Dim commaPos As Long
            commaPos = InStr(1, s, ",", vbBinaryCompare)

            If commaPos > 0 Then
                Dim afterComma As String
                afterComma = Trim$(Mid$(s, commaPos + 1))

                If StartsWithPersonnummer(afterComma) Then
                    Dim candidate As String
                    candidate = Trim$(Left$(s, commaPos - 1))

                    ' Ta bort eventuell etikett före namn, t.ex. "Motpart:" eller "Övriga:"
                    Dim colonPos As Long
                    colonPos = InStrRev(candidate, ":", -1, vbBinaryCompare)
                    If colonPos > 0 Then
                        candidate = Trim$(Mid$(candidate, colonPos + 1))
                    End If

                    If Len(candidate) > 0 Then
                        If Not NameExists(result, count, candidate) Then
                            count = count + 1
                            ReDim Preserve result(0 To count)
                            result(count) = candidate
                        End If
                    End If
                End If
            End If
        End If
    Next line

    If count = -1 Then
        ExtractPersonNames = Array()
    Else
        ExtractPersonNames = result
    End If
End Function

Private Function StartsWithPersonnummer(ByVal s As String) As Boolean
    If Len(s) < 11 Then Exit Function

    StartsWithPersonnummer = _
        IsDigit(Mid$(s, 1, 1)) And _
        IsDigit(Mid$(s, 2, 1)) And _
        IsDigit(Mid$(s, 3, 1)) And _
        IsDigit(Mid$(s, 4, 1)) And _
        IsDigit(Mid$(s, 5, 1)) And _
        IsDigit(Mid$(s, 6, 1)) And _
        Mid$(s, 7, 1) = "-" And _
        IsDigit(Mid$(s, 8, 1)) And _
        IsDigit(Mid$(s, 9, 1)) And _
        IsDigit(Mid$(s, 10, 1)) And _
        IsDigit(Mid$(s, 11, 1))
End Function

Private Function NormalizeName(ByVal s As String) As String
    s = Trim$(s)

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeName = LCase$(s)
End Function

Private Function NameExists(ByRef arr() As String, ByVal count As Long, ByVal candidate As String) As Boolean
    Dim i As Long

    If count < 0 Then Exit Function

    For i = 0 To count
        If NormalizeName(arr(i)) = NormalizeName(candidate) Then
            NameExists = True
            Exit Function
        End If
    Next i
End Function


Public Sub Process_YTTRANDE_SIGNATUR(ByVal content As Range)
    Dim model As YttrandeSignaturModel
    model = ReadYttrandeSignatur()
    model = TransformYttrandeSignatur(model)
    RenderYttrandeSignatur content, model
End Sub

Private Function ReadYttrandeSignatur() As YttrandeSignaturModel
    ReadYttrandeSignatur.Namn = GetFullName()
    ReadYttrandeSignatur.Titel = GetTitle()
    ReadYttrandeSignatur.Postort = GetCity()
End Function

Private Function TransformYttrandeSignatur(model As YttrandeSignaturModel) As YttrandeSignaturModel
    If Len(Trim$(model.Postort)) = 0 Then
        model.Postort = "Lund"
    End If
    TransformYttrandeSignatur = model
End Function

Private Sub RenderYttrandeSignatur(ByVal content As Range, model As YttrandeSignaturModel)
    content.text = SwedishDateText(model.Postort) & vbCr & vbCr & model.Namn & vbCr & model.Titel
End Sub

Public Sub Process_YTTRANDE_PARTER(ByVal content As Range)
    Dim inputModel As YttrandeParterReadModel
    inputModel = ReadYttrandeParter(content)

    Dim computed As YttrandeParterComputedModel
    computed = TransformYttrandeParter(inputModel)
    If Not computed.IsValid Then Exit Sub

    RenderYttrandeParter content, computed
End Sub

Private Function ReadYttrandeParter(ByVal content As Range) As YttrandeParterReadModel
    ReadYttrandeParter.Raw = content.text
    ReadYttrandeParter.Raw = Replace(ReadYttrandeParter.Raw, vbCrLf, vbCr)
    ReadYttrandeParter.Raw = Replace(ReadYttrandeParter.Raw, Chr(11), vbCr)
    ReadYttrandeParter.Raw = Replace(ReadYttrandeParter.Raw, Chr(7), "")
    ReadYttrandeParter.Raw = Trim$(ReadYttrandeParter.Raw)

    If Len(ReadYttrandeParter.Raw) = 0 Then Exit Function
End Function

Private Function TransformYttrandeParter(inputModel As YttrandeParterReadModel) As YttrandeParterComputedModel
    Dim lines() As String
    On Error GoTo ParseFailed
    lines = Split(inputModel.Raw, vbCr)

    Dim i As Long
    If UBound(lines) >= 1 Then
        TransformYttrandeParter.LeftParty = ExtractLeftParty(lines(1))
    End If

    For i = LBound(lines) To UBound(lines)
        If InStr(1, lines(i), "Motpart:", vbTextCompare) > 0 Then
            TransformYttrandeParter.RightParty = ExtractMotpartName(lines(i))
            Exit For
        End If
    Next i

    If Len(TransformYttrandeParter.LeftParty) = 0 Then Exit Function
    If Len(TransformYttrandeParter.RightParty) = 0 Then Exit Function

    TransformYttrandeParter.Names = ExtractPersonNames(TransformYttrandeParter.LeftParty, inputModel.Raw)
    TransformYttrandeParter.IsValid = True
    Exit Function

ParseFailed:
    TransformYttrandeParter.IsValid = False
End Function

Private Sub RenderYttrandeParter(ByVal content As Range, computed As YttrandeParterComputedModel)
    ' Byt fortfarande ut [KundNamn] med default-värdet för vänster part
    ReplaceKundNamnEverywhere ActiveDocument, computed.LeftParty
    ReplaceRangeWithPartyDropdowns content, computed.Names, computed.LeftParty, computed.RightParty
End Sub

Private Function FindLiteralRange(ByVal searchIn As Range, ByVal needle As String) As Range
    Dim r As Range
    Set r = searchIn.Duplicate

    With r.Find
        .ClearFormatting
        .text = needle
        .Forward = True
        .Wrap = wdFindStop
        .format = False
        .MatchWildcards = False
        .MatchCase = True

        If .Execute Then
            Set FindLiteralRange = r.Duplicate
        Else
            Set FindLiteralRange = Nothing
        End If
    End With
End Function

Private Sub ReplaceRangeWithPartyDropdowns(ByVal target As Range, ByVal names As Variant, ByVal defaultLeft As String, ByVal defaultRight As String)
    Dim doc As Document
    Set doc = target.Document

    ' Nollställ formatering på hela målområdet först
    target.Font.Underline = wdUnderlineNone

    ' Bygg först ett vanligt textblock med två markörer
    target.text = "LEFT_PARTY ./. RIGHT_PARTY"

    Dim workRng As Range
    Set workRng = target.Duplicate

    Dim ccLeft As ContentControl
    Dim ccRight As ContentControl
    Dim leftRng As Range
    Dim rightRng As Range

    ' Vänster markör -> dropdown
    Set leftRng = FindLiteralRange(workRng, "LEFT_PARTY")
    If leftRng Is Nothing Then Exit Sub

    Set ccLeft = doc.ContentControls.Add(wdContentControlDropdownList, leftRng)
    SetupPartyDropdown ccLeft, names, defaultLeft, "Yttrande vänster part", "YTTRANDE_PART_LEFT"

    ' Höger markör -> dropdown
    Set rightRng = FindLiteralRange(workRng, "RIGHT_PARTY")
    If rightRng Is Nothing Then Exit Sub

    Set ccRight = doc.ContentControls.Add(wdContentControlDropdownList, rightRng)
    SetupPartyDropdown ccRight, names, defaultRight, "Yttrande höger part", "YTTRANDE_PART_RIGHT"

    ' Formatering:
    ' endast vänster dropdown ska vara understruken
    ccLeft.Range.Font.Underline = wdUnderlineSingle
    ccRight.Range.Font.Underline = wdUnderlineNone

    ' Separatorn mellan dropdowns ska inte vara understruken
    Dim sepRng As Range
    Set sepRng = doc.Range(ccLeft.Range.End, ccRight.Range.start)
    sepRng.Font.Underline = wdUnderlineNone
End Sub

Private Sub SetupPartyDropdown(ByVal cc As ContentControl, ByVal names As Variant, ByVal defaultName As String, ByVal title As String, ByVal tag As String)
    Dim i As Long

    cc.title = title
    cc.tag = tag

    If cc.ShowingPlaceholderText Then
        cc.SetPlaceholderText , , "Välj part"
    End If

    ' Rensa eventuella gamla entries
    Do While cc.DropdownListEntries.count > 0
        cc.DropdownListEntries(1).Delete
    Loop

    For i = LBound(names) To UBound(names)
        If Len(Trim$(CStr(names(i)))) > 0 Then
            cc.DropdownListEntries.Add CStr(names(i))
        End If
    Next i

    SetDropdownDefaultByText cc, defaultName
End Sub

Private Sub SetDropdownDefaultByText(ByVal cc As ContentControl, ByVal wantedText As String)
    Dim i As Long

    For i = 1 To cc.DropdownListEntries.count
        If StrComp(Trim$(cc.DropdownListEntries(i).text), Trim$(wantedText), vbTextCompare) = 0 Then
            cc.DropdownListEntries(i).Select
            Exit Sub
        End If
    Next i
End Sub

Private Function ExtractLeftParty(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)

    p = InStr(1, s, "./.", vbTextCompare)
    If p > 0 Then
        ExtractLeftParty = Trim$(Left$(s, p - 1))
    Else
        ExtractLeftParty = Trim$(s)
    End If
End Function

Private Function ExtractMotpartName(ByVal s As String) As String
    Dim p As Long
    Dim rest As String
    Dim commaPos As Long

    p = InStr(1, s, "Motpart:", vbTextCompare)
    If p = 0 Then
        ExtractMotpartName = ""
        Exit Function
    End If

    rest = Trim$(Mid$(s, p + Len("Motpart:")))
    commaPos = InStr(1, rest, ",", vbBinaryCompare)

    If commaPos > 0 Then
        ExtractMotpartName = Trim$(Left$(rest, commaPos - 1))
    Else
        ExtractMotpartName = Trim$(rest)
    End If
End Function

Private Sub ReplaceKundNamnEverywhere(ByVal doc As Document, ByVal kundNamn As String)
    ReplaceLiteralInRange doc.content, "[KundNamn]", kundNamn

    Dim shp As Shape
    For Each shp In doc.shapes
        ReplaceInShapeRecursive shp, "[KundNamn]", kundNamn
    Next shp

    Dim sec As Section
    Dim hf As HeaderFooter

    For Each sec In doc.Sections
        For Each hf In sec.Headers
            ReplaceLiteralInRange hf.Range, "[KundNamn]", kundNamn
            For Each shp In hf.shapes
                ReplaceInShapeRecursive shp, "[KundNamn]", kundNamn
            Next shp
        Next hf

        For Each hf In sec.Footers
            ReplaceLiteralInRange hf.Range, "[KundNamn]", kundNamn
            For Each shp In hf.shapes
                ReplaceInShapeRecursive shp, "[KundNamn]", kundNamn
            Next shp
        Next hf
    Next sec
End Sub

Private Sub ReplaceInShapeRecursive(ByVal shp As Shape, ByVal findText As String, ByVal replaceText As String)
    On Error Resume Next

    If shp.TextFrame.HasText Then
        ReplaceLiteralInRange shp.TextFrame.TextRange, findText, replaceText
    End If

    If shp.Type = msoGroup Then
        Dim gi As Shape
        For Each gi In shp.GroupItems
            ReplaceInShapeRecursive gi, findText, replaceText
        Next gi
    End If

    On Error GoTo 0
End Sub

Private Sub ReplaceLiteralInRange(ByVal rng As Range, ByVal findText As String, ByVal replaceText As String)
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
