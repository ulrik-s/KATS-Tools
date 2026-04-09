Option Explicit

Public Type MoGUser
    UName As String
    ShortName As String
    FullName As String
    Mileage As Currency
    Title As String
    City As String
    MatchKeys As String   ' semicolon-separated aliases
End Type

Private users() As MoGUser
Private usersLoaded As Boolean

Private resolvedUserIdx As Integer
Private resolvedUserReady As Boolean

' ============================================================
' Public API
' ============================================================

Public Function GetFullName() As String
    EnsureLoaded
    GetFullName = users(GetUserIdx()).FullName
End Function

Public Function GetTitle() As String
    EnsureLoaded
    GetTitle = users(GetUserIdx()).Title
End Function

Public Function GetMileage() As Currency
    EnsureLoaded
    GetMileage = users(GetUserIdx()).Mileage
End Function

Public Function GetCity() As String
    EnsureLoaded
    GetCity = users(GetUserIdx()).City
End Function

Public Function GetCurrentResolvedUserName() As String
    EnsureLoaded
    GetCurrentResolvedUserName = users(GetUserIdx()).FullName
End Function

Public Function GetCurrentRawUserIdentity() As String
    GetCurrentRawUserIdentity = GetPrimaryCurrentIdentity()
End Function

' ============================================================
' Debug helpers
' ============================================================

Public Sub Debug_ShowResolvedUserReport()
    MsgBox Debug_GetResolvedUserReport(), vbInformation, "MoGUser Debug"
End Sub

Public Function Debug_GetResolvedUserReport() As String
    EnsureLoaded

    Dim idx As Integer
    idx = GetUserIdx()

    Debug_GetResolvedUserReport = _
        "Raw USERNAME: " & Environ$("USERNAME") & vbCrLf & _
        "Raw USER: " & Environ$("USER") & vbCrLf & _
        "Raw LOGNAME: " & Environ$("LOGNAME") & vbCrLf & vbCrLf & _
        "Primary identity: " & GetPrimaryCurrentIdentity() & vbCrLf & _
        "Candidate keys: " & GetCurrentIdentityCandidates() & vbCrLf & vbCrLf & _
        "Resolved index: " & CStr(idx) & vbCrLf & _
        "Resolved uname: " & users(idx).UName & vbCrLf & _
        "Resolved full name: " & users(idx).FullName & vbCrLf & _
        "Resolved match keys: " & BuildUserKeyList(users(idx))
End Function

Public Function Debug_DumpAllUsers() As String
    EnsureLoaded

    Dim i As Long
    Dim s As String

    For i = LBound(users) To UBound(users)
        s = s & "[" & CStr(i) & "] " & users(i).FullName & vbCrLf
        s = s & "  uname: " & users(i).UName & vbCrLf
        s = s & "  short: " & users(i).ShortName & vbCrLf
        s = s & "  city: " & users(i).City & vbCrLf
        s = s & "  keys: " & BuildUserKeyList(users(i)) & vbCrLf & vbCrLf
    Next i

    Debug_DumpAllUsers = s
End Function

Public Sub Debug_ShowAllUsers()
    MsgBox Debug_DumpAllUsers(), vbInformation, "MoGUser Database"
End Sub

' ============================================================
' Load / database
' ============================================================

Public Sub LoadMoGUsers()
    ReDim users(0 To 9)

    AddUser 0, _
        "default", _
        TxtNagon(), _
        TxtNagon() & " " & TxtOkand(), _
        37@, _
        TxtBitradandeJurist(), _
        "Lund", _
        "default;nagon;okand;unknown"

    AddUser 1, _
        "cecilia", _
        "Cecilia", _
        "Cecilia Moll", _
        9.5@, _
        "Advokat", _
        "Lund", _
        "ceciliamoll"

    AddUser 2, _
        "alma", _
        "Alma", _
        "Alma Diaz R" & ChrW$(228) & "m" & ChrW$(246), _
        37@, _
        TxtBitradandeJurist(), _
        "Lund", _
        "almadiazramo;almaramo"

    AddUser 3, _
        "ulrik", _
        "Ulrik", _
        "Ulrik Sj" & ChrW$(246) & "lin", _
        483.99@, _
        "Ers Kjeserliga " & ChrW$(214) & "verh" & ChrW$(246) & "ghet", _
        "Utopia", _
        "ulriksjolin;ulriksjoelin;ulriksjolin1"

    AddUser 4, _
        "mans", _
        "M" & ChrW$(229) & "ns", _
        "M" & ChrW$(229) & "ns Bergendorff", _
        37@, _
        "Advokat", _
        "Malm" & ChrW$(246), _
        "mansbergendorff"

    AddUser 5, _
        "azar", _
        "Azar", _
        "Azar Akbarian", _
        37@, _
        TxtBitradandeJurist(), _
        "Malm" & ChrW$(246), _
        "azarakbarian"

    AddUser 6, _
        "petra", _
        "Petra", _
        "Petra Ramberg Persson", _
        37@, _
        "Advokat", _
        "Lund", _
        "petrarambergpersson;petrapersson"

    AddUser 7, _
        "annette", _
        "Annette", _
        "Annette Lantz", _
        37@, _
        "Advokatsekreterare", _
        "Lund", _
        "annettelantz"

    AddUser 8, _
        "maria", _
        "Maria", _
        "Maria Grosskopf", _
        12@, _
        "Advokat", _
        "Lund", _
        "mariagrosskopf"

    AddUser 9, _
        "jakobenoksson", _
        "Jakob", _
        "Jakob Enoksson", _
        37@, _
        TxtBitradandeJurist(), _
        "Lund", _
        "jakobenoksson;jakobenokson"

    usersLoaded = True
    resolvedUserReady = False
End Sub

Private Sub AddUser( _
    ByVal idx As Long, _
    ByVal uname As String, _
    ByVal shortName As String, _
    ByVal fullName As String, _
    ByVal mileage As Currency, _
    ByVal title As String, _
    ByVal city As String, _
    ByVal matchKeys As String)

    users(idx).UName = uname
    users(idx).ShortName = shortName
    users(idx).FullName = fullName
    users(idx).Mileage = mileage
    users(idx).Title = title
    users(idx).City = city
    users(idx).MatchKeys = matchKeys
End Sub

Private Sub EnsureLoaded()
    If usersLoaded Then Exit Sub
    LoadMoGUsers
End Sub

' ============================================================
' Resolve current user
' ============================================================

Private Function GetUserIdx() As Integer
    EnsureLoaded

    If resolvedUserReady Then
        GetUserIdx = resolvedUserIdx
        Exit Function
    End If

    resolvedUserIdx = ResolveCurrentUserIdx()
    resolvedUserReady = True
    GetUserIdx = resolvedUserIdx
End Function

Private Function ResolveCurrentUserIdx() As Integer
    Dim candidates As String
    candidates = GetCurrentIdentityCandidates()

    Dim parts() As String
    Dim i As Long
    Dim key As String
    Dim j As Long

    If Len(candidates) > 0 Then
        parts = Split(candidates, ";")

        ' Pass 1: exact uname only
        For i = LBound(parts) To UBound(parts)
            key = Trim$(parts(i))
            If Len(key) > 0 Then
                For j = LBound(users) To UBound(users)
                    If NormalizeUserKey(users(j).UName) = key Then
                        ResolveCurrentUserIdx = j
                        Exit Function
                    End If
                Next j
            End If
        Next i

        ' Pass 2: all generated keys + aliases
        For i = LBound(parts) To UBound(parts)
            key = Trim$(parts(i))
            If Len(key) > 0 Then
                For j = LBound(users) To UBound(users)
                    If UserMatches(users(j), key) Then
                        ResolveCurrentUserIdx = j
                        Exit Function
                    End If
                Next j
            End If
        Next i
    End If

    ResolveCurrentUserIdx = 0
End Function

Private Function UserMatches(ByRef u As MoGUser, ByVal candidateKey As String) As Boolean
    If Len(candidateKey) = 0 Then Exit Function
    UserMatches = KeyListContains(BuildUserKeyList(u), candidateKey)
End Function

' ============================================================
' Current identity candidates
' ============================================================

Private Function GetPrimaryCurrentIdentity() As String
    Dim s As String

    s = Trim$(Environ$("USERNAME"))
    If Len(s) = 0 Then s = Trim$(Environ$("USER"))
    If Len(s) = 0 Then s = Trim$(Environ$("LOGNAME"))

    GetPrimaryCurrentIdentity = s
End Function

Private Function GetCurrentIdentityCandidates() As String
    Dim result As String

    AddIdentityCandidates result, Environ$("USERNAME")
    AddIdentityCandidates result, Environ$("USER")
    AddIdentityCandidates result, Environ$("LOGNAME")

    GetCurrentIdentityCandidates = result
End Function

Private Sub AddIdentityCandidates(ByRef keyList As String, ByVal rawValue As String)
    Dim s As String
    Dim p As Long

    s = Trim$(rawValue)
    If Len(s) = 0 Then Exit Sub

    AddUniqueKey keyList, NormalizeUserKey(s)

    ' DOMAIN\username
    p = InStrRev(s, "\")
    If p > 0 And p < Len(s) Then
        AddUniqueKey keyList, NormalizeUserKey(Mid$(s, p + 1))
    End If

    ' domain/username
    p = InStrRev(s, "/")
    If p > 0 And p < Len(s) Then
        AddUniqueKey keyList, NormalizeUserKey(Mid$(s, p + 1))
    End If

    ' email local-part
    p = InStr(1, s, "@")
    If p > 1 Then
        AddUniqueKey keyList, NormalizeUserKey(Left$(s, p - 1))
    End If
End Sub

' ============================================================
' User key generation
' ============================================================

Private Function BuildUserKeyList(ByRef u As MoGUser) As String
    Dim keys As String
    Dim firstName As String
    Dim lastName As String
    Dim collapsedFull As String

    AddUniqueKey keys, u.UName
    AddUniqueKey keys, u.ShortName
    AddUniqueKey keys, u.FullName

    firstName = GetFirstWord(u.FullName)
    lastName = GetLastWord(u.FullName)
    collapsedFull = CollapseWords(u.FullName)

    AddUniqueKey keys, firstName
    AddUniqueKey keys, firstName & lastName
    AddUniqueKey keys, collapsedFull

    AddAliasKeys keys, u.MatchKeys

    BuildUserKeyList = keys
End Function

Private Sub AddAliasKeys(ByRef keyList As String, ByVal aliases As String)
    Dim parts() As String
    Dim i As Long

    aliases = Trim$(aliases)
    If Len(aliases) = 0 Then Exit Sub

    parts = Split(aliases, ";")
    For i = LBound(parts) To UBound(parts)
        AddUniqueKey keyList, parts(i)
    Next i
End Sub

' ============================================================
' Key list helpers
' ============================================================

Private Sub AddUniqueKey(ByRef keyList As String, ByVal key As String)
    key = NormalizeUserKey(key)
    If Len(key) = 0 Then Exit Sub

    If KeyListContains(keyList, key) Then Exit Sub

    If Len(keyList) = 0 Then
        keyList = key
    Else
        keyList = keyList & ";" & key
    End If
End Sub

Private Function KeyListContains(ByVal keyList As String, ByVal key As String) As Boolean
    Dim parts() As String
    Dim i As Long

    key = NormalizeUserKey(key)
    If Len(key) = 0 Then Exit Function
    If Len(keyList) = 0 Then Exit Function

    parts = Split(keyList, ";")
    For i = LBound(parts) To UBound(parts)
        If StrComp(Trim$(parts(i)), key, vbTextCompare) = 0 Then
            KeyListContains = True
            Exit Function
        End If
    Next i
End Function

' ============================================================
' Name parsing helpers
' ============================================================

Private Function NormalizeSpaces(ByVal s As String) As String
    s = Trim$(s)
    s = Replace(s, vbTab, " ")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeSpaces = s
End Function

Private Function GetFirstWord(ByVal s As String) As String
    Dim cleaned As String
    Dim p As Long

    cleaned = NormalizeSpaces(s)
    If Len(cleaned) = 0 Then Exit Function

    p = InStr(1, cleaned, " ")
    If p = 0 Then
        GetFirstWord = cleaned
    Else
        GetFirstWord = Left$(cleaned, p - 1)
    End If
End Function

Private Function GetLastWord(ByVal s As String) As String
    Dim cleaned As String
    Dim p As Long

    cleaned = NormalizeSpaces(s)
    If Len(cleaned) = 0 Then Exit Function

    p = InStrRev(cleaned, " ")
    If p = 0 Then
        GetLastWord = cleaned
    Else
        GetLastWord = Mid$(cleaned, p + 1)
    End If
End Function

Private Function CollapseWords(ByVal s As String) As String
    CollapseWords = Replace(NormalizeSpaces(s), " ", "")
End Function

' ============================================================
' Normalization
' ============================================================

Private Function NormalizeUserKey(ByVal s As String) As String
    s = LCase$(Trim$(s))

    ' Swedish letters via ChrW$
    s = Replace(s, ChrW$(229), "a") ' å
    s = Replace(s, ChrW$(228), "a") ' ä
    s = Replace(s, ChrW$(246), "o") ' ö
    s = Replace(s, ChrW$(197), "a") ' Å
    s = Replace(s, ChrW$(196), "a") ' Ä
    s = Replace(s, ChrW$(214), "o") ' Ö

    ' Common Latin accents
    s = Replace(s, ChrW$(225), "a") ' á
    s = Replace(s, ChrW$(224), "a") ' à
    s = Replace(s, ChrW$(226), "a") ' â
    s = Replace(s, ChrW$(227), "a") ' ã

    s = Replace(s, ChrW$(233), "e") ' é
    s = Replace(s, ChrW$(232), "e") ' è
    s = Replace(s, ChrW$(234), "e") ' ê
    s = Replace(s, ChrW$(235), "e") ' ë

    s = Replace(s, ChrW$(237), "i") ' í
    s = Replace(s, ChrW$(236), "i") ' ì
    s = Replace(s, ChrW$(238), "i") ' î
    s = Replace(s, ChrW$(239), "i") ' ï

    s = Replace(s, ChrW$(243), "o") ' ó
    s = Replace(s, ChrW$(242), "o") ' ò
    s = Replace(s, ChrW$(244), "o") ' ô
    s = Replace(s, ChrW$(245), "o") ' õ

    s = Replace(s, ChrW$(250), "u") ' ú
    s = Replace(s, ChrW$(249), "u") ' ù
    s = Replace(s, ChrW$(251), "u") ' û
    s = Replace(s, ChrW$(252), "u") ' ü

    s = Replace(s, ChrW$(241), "n") ' ñ

    ' Punctuation / separators
    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")
    s = Replace(s, ".", "")
    s = Replace(s, ",", "")
    s = Replace(s, ";", "")
    s = Replace(s, ":", "")
    s = Replace(s, "/", "")
    s = Replace(s, "\", "")
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, "{", "")
    s = Replace(s, "}", "")
    s = Replace(s, "'", "")
    s = Replace(s, ChrW$(8217), "") ' typographic apostrophe
    s = Replace(s, """", "")
    s = Replace(s, "`", "")
    s = Replace(s, "~", "")
    s = Replace(s, "!", "")
    s = Replace(s, "?", "")
    s = Replace(s, "#", "")
    s = Replace(s, "$", "")
    s = Replace(s, "%", "")
    s = Replace(s, "&", "")
    s = Replace(s, "*", "")
    s = Replace(s, "+", "")
    s = Replace(s, "=", "")
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Replace(s, "@", "")

    NormalizeUserKey = s
End Function

' ============================================================
' Small text builders for Swedish characters
' ============================================================

Private Function TxtNagon() As String
    TxtNagon = "N" & ChrW$(229) & "gon"
End Function

Private Function TxtOkand() As String
    TxtOkand = "Ok" & ChrW$(228) & "nd"
End Function

Private Function TxtBitradandeJurist() As String
    TxtBitradandeJurist = "Bitr" & ChrW$(228) & "dande jurist"
End Function
