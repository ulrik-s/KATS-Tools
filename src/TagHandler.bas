Option Explicit

' ============================================================
' Runner: process one tag everywhere (body + shapes + headers/footers)
' ============================================================

Public Sub ProcessTagEverywhere(ByVal doc As Document, ByVal tag As String, ByVal procName As String)
    ProcessTagInRange doc.content, tag, procName
    ProcessShapesCollection doc.shapes, tag, procName

    Dim sec As Section
    For Each sec In doc.Sections
        ProcessHeaderFooter sec.Headers, tag, procName
        ProcessHeaderFooter sec.Footers, tag, procName
    Next sec
End Sub

Public Function ConsumeFirstTagRangeEverywhere(ByVal doc As Document, ByVal tag As String) As Range
    Set ConsumeFirstTagRangeEverywhere = ConsumeFirstTagRangeInRange(doc.content, tag)
    If Not ConsumeFirstTagRangeEverywhere Is Nothing Then Exit Function

    Set ConsumeFirstTagRangeEverywhere = ConsumeFirstTagRangeInShapes(doc.shapes, tag)
    If Not ConsumeFirstTagRangeEverywhere Is Nothing Then Exit Function

    Dim sec As Section
    For Each sec In doc.Sections
        Set ConsumeFirstTagRangeEverywhere = ConsumeFirstTagRangeInHeaderFooter(sec.Headers, tag)
        If Not ConsumeFirstTagRangeEverywhere Is Nothing Then Exit Function

        Set ConsumeFirstTagRangeEverywhere = ConsumeFirstTagRangeInHeaderFooter(sec.Footers, tag)
        If Not ConsumeFirstTagRangeEverywhere Is Nothing Then Exit Function
    Next sec
End Function

Private Sub ProcessHeaderFooter(ByVal hfs As HeadersFooters, ByVal tag As String, ByVal procName As String)
    Dim hf As HeaderFooter
    For Each hf In hfs
        On Error Resume Next
        ProcessTagInRange hf.Range, tag, procName
        ProcessShapesCollection hf.shapes, tag, procName
        On Error GoTo 0
    Next hf
End Sub

Private Sub ProcessShapesCollection(ByVal shapes As shapes, ByVal tag As String, ByVal procName As String)
    Dim shp As Shape
    For Each shp In shapes
        ProcessShapeRecursive shp, tag, procName
    Next shp
End Sub

Private Sub ProcessShapeRecursive(ByVal shp As Shape, ByVal tag As String, ByVal procName As String)
    If ShapeHasText(shp) Then
        ProcessTagInRange shp.TextFrame.TextRange, tag, procName
    End If

    On Error Resume Next
    If shp.Type = msoGroup Then
        Dim gi As Shape
        For Each gi In shp.GroupItems
            ProcessShapeRecursive gi, tag, procName
        Next gi
    End If
    On Error GoTo 0
End Sub

Private Function ShapeHasText(ByVal shp As Shape) As Boolean
    On Error GoTo Nope
    ShapeHasText = shp.TextFrame.HasText
    Exit Function
Nope:
    ShapeHasText = False
End Function

Private Function ConsumeFirstTagRangeInHeaderFooter(ByVal hfs As HeadersFooters, ByVal tag As String) As Range
    Dim hf As HeaderFooter
    For Each hf In hfs
        On Error Resume Next
        Set ConsumeFirstTagRangeInHeaderFooter = ConsumeFirstTagRangeInRange(hf.Range, tag)
        On Error GoTo 0
        If Not ConsumeFirstTagRangeInHeaderFooter Is Nothing Then Exit Function

        On Error Resume Next
        Set ConsumeFirstTagRangeInHeaderFooter = ConsumeFirstTagRangeInShapes(hf.shapes, tag)
        On Error GoTo 0
        If Not ConsumeFirstTagRangeInHeaderFooter Is Nothing Then Exit Function
    Next hf
End Function

Private Function ConsumeFirstTagRangeInShapes(ByVal shapes As shapes, ByVal tag As String) As Range
    Dim shp As Shape
    For Each shp In shapes
        Set ConsumeFirstTagRangeInShapes = ConsumeFirstTagRangeInShape(shp, tag)
        If Not ConsumeFirstTagRangeInShapes Is Nothing Then Exit Function
    Next shp
End Function

Private Function ConsumeFirstTagRangeInShape(ByVal shp As Shape, ByVal tag As String) As Range
    If ShapeHasText(shp) Then
        Set ConsumeFirstTagRangeInShape = ConsumeFirstTagRangeInRange(shp.TextFrame.TextRange, tag)
        If Not ConsumeFirstTagRangeInShape Is Nothing Then Exit Function
    End If

    On Error Resume Next
    If shp.Type = msoGroup Then
        Dim gi As Shape
        For Each gi In shp.GroupItems
            Set ConsumeFirstTagRangeInShape = ConsumeFirstTagRangeInShape(gi, tag)
            If Not ConsumeFirstTagRangeInShape Is Nothing Then Exit Function
        Next gi
    End If
    On Error GoTo 0
End Function

' ============================================================
' Core scanner for ONE tag in ONE range
'
' Requires exact markers:
'   [[KATS_<TAG>_START]]
'   [[KATS_<TAG>_END]]
'
' Deletes markers BEFORE calling processor (END first).
' No nested-tag support.
' ============================================================

Private Sub ProcessTagInRange(ByVal root As Range, ByVal tag As String, ByVal procName As String)
    Dim startText As String, endText As String
    startText = "[[KATS_" & tag & "_START]]"
    endText = "[[KATS_" & tag & "_END]]"

    Dim searchRng As Range
    Set searchRng = root.Duplicate
    searchRng.Collapse wdCollapseStart

    Do While FindNextText(searchRng, startText)

        Dim startMarker As Range
        Set startMarker = searchRng.Duplicate  ' Find range == match

        Dim endSearch As Range
        Set endSearch = root.Duplicate
        endSearch.SetRange startMarker.End, root.End
        endSearch.Collapse wdCollapseStart

        If Not FindNextText(endSearch, endText) Then
            ' Broken block: skip past START marker and continue
            Set searchRng = root.Duplicate
            searchRng.SetRange startMarker.End, root.End
            searchRng.Collapse wdCollapseStart
            GoTo ContinueLoop
        End If

        Dim endMarker As Range
        Set endMarker = endSearch.Duplicate

        ' Capture content length BEFORE deletions
        Dim contentLen As Long
        contentLen = endMarker.start - startMarker.End

        Dim startPos As Long
        startPos = startMarker.start

        ' Delete markers (END first!)
        endMarker.text = ""
        startMarker.text = ""

        ' Recreate stable content range
        Dim content As Range
        Set content = root.Duplicate
        content.SetRange startPos, startPos + contentLen

        ' Call processor chosen by pipeline (no select-case here)
        On Error GoTo ProcFail
        Application.Run procName, content
        On Error GoTo 0

        ' Continue scanning after content
        Set searchRng = root.Duplicate
        searchRng.SetRange content.End, root.End
        searchRng.Collapse wdCollapseStart

ContinueLoop:
    Loop
    Exit Sub

ProcFail:
    Dim n As Long, d As String
    n = Err.number: d = Err.Description
    On Error GoTo 0
    Err.Raise n, "ProcessTagInRange(" & procName & ")", d
End Sub

Private Function FindNextText(ByRef rng As Range, ByVal what As String) As Boolean
    With rng.Find
        .ClearFormatting
        .text = what
        .Forward = True
        .Wrap = wdFindStop
        .format = False
        .MatchWildcards = False
        .MatchCase = True
        FindNextText = .Execute
    End With
End Function

Private Function ConsumeFirstTagRangeInRange(ByVal root As Range, ByVal tag As String) As Range
    Dim startText As String, endText As String
    startText = "[[KATS_" & tag & "_START]]"
    endText = "[[KATS_" & tag & "_END]]"

    Dim searchRng As Range
    Set searchRng = root.Duplicate
    searchRng.Collapse wdCollapseStart

    If Not FindNextText(searchRng, startText) Then Exit Function

    Dim startMarker As Range
    Set startMarker = searchRng.Duplicate

    Dim endSearch As Range
    Set endSearch = root.Duplicate
    endSearch.SetRange startMarker.End, root.End
    endSearch.Collapse wdCollapseStart

    If Not FindNextText(endSearch, endText) Then Exit Function

    Dim endMarker As Range
    Set endMarker = endSearch.Duplicate

    Dim contentLen As Long
    contentLen = endMarker.start - startMarker.End

    Dim startPos As Long
    startPos = startMarker.start

    endMarker.text = ""
    startMarker.text = ""

    Dim content As Range
    Set content = root.Duplicate
    content.SetRange startPos, startPos + contentLen

    Set ConsumeFirstTagRangeInRange = content
End Function
