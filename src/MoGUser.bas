    Option Explicit

' Gör typen public så andra moduler kan använda den om de vill
Public Type MoGUser
    uname As String
    ShortName As String
    fullName As String
    Mileage As Currency
    title As String
    City As String
End Type

' Modul-global array
Private users(0 To 10) As MoGUser
Private usersLoaded As Boolean

Public Sub LoadMoGUsers()
    usersLoaded = True

    users(0).uname = "default"
    users(0).ShortName = "Någon"
    users(0).fullName = "Någon Okänd"
    users(0).Mileage = 37
    users(0).title = "Biträdande jurist"
    users(0).City = "Lund"

    users(1).uname = "cecilia"
    users(1).ShortName = "Cecilia"
    users(1).fullName = "Cecilia Moll"
    users(1).Mileage = 9.5
    users(1).title = "Advokat"
    users(1).City = "Lund"

    users(2).uname = "alma"
    users(2).ShortName = "Alma"
    users(2).fullName = "Alma Diaz Rämö"
    users(2).Mileage = 37
    users(2).title = "Biträdande jurist"
    users(2).City = "Lund"

    users(3).uname = "ulrik"
    users(3).ShortName = "Ulrik"
    users(3).fullName = "Ulrik Sjölin"
    users(3).Mileage = 483.99
    users(3).title = "Ers Kjeserliga Överhöghet"
    users(3).City = "Utopia"

    users(4).uname = "mans"
    users(4).ShortName = "Måns"
    users(4).fullName = "Måns Bergendorff"
    users(4).Mileage = 37
    users(4).title = "Advokat"
    users(4).City = "Malmö"

    users(5).uname = "azar"
    users(5).ShortName = "Azar"
    users(5).fullName = "Azar Akbarian"
    users(5).Mileage = 37
    users(5).title = "Biträdande jurist"
    users(5).City = "Malmö"

    users(6).uname = "petra"
    users(6).ShortName = "Petra"
    users(6).fullName = "Petra Ramberg Persson"
    users(6).Mileage = 37
    users(6).title = "Advokat"
    users(6).City = "Lund"

    users(7).uname = "annette"
    users(7).ShortName = "Annnette"
    users(7).fullName = "Annette Lantz"
    users(7).Mileage = 37
    users(7).title = "Advokatsekreterare"
    users(7).City = "Lund"

    users(8).uname = "maria"
    users(8).ShortName = "Maria"
    users(8).fullName = "Maria Grosskopf"
    users(8).Mileage = 12
    users(8).title = "Advokat"
    users(8).City = "Lund"

    users(9).uname = "jakobenoksson"
    users(9).ShortName = "Jakob"
    users(9).fullName = "Jakob Enoksson"
    users(9).Mileage = 37
    users(9).title = "Biträdande jurist"
    users(9).City = "Lund"
End Sub

Private Sub EnsureLoaded()
    If usersLoaded Then Exit Sub
    LoadMoGUsers
End Sub

Private Function NormalizeUserKey(ByVal s As String) As String
    s = LCase$(Trim$(s))

    s = Replace(s, "å", "a")
    s = Replace(s, "ä", "a")
    s = Replace(s, "ö", "o")
    s = Replace(s, "é", "e")
    s = Replace(s, "è", "e")
    s = Replace(s, "ê", "e")
    s = Replace(s, "ü", "u")

    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")
    s = Replace(s, ".", "")
    s = Replace(s, "'", "")
    s = Replace(s, ChrW$(8217), "") ' typografisk apostrof

    NormalizeUserKey = s
End Function

Private Function GetFirstName(ByVal fullName As String) As String
    Dim parts() As String
    fullName = Trim$(fullName)
    If Len(fullName) = 0 Then Exit Function

    parts = Split(fullName, " ")
    GetFirstName = parts(LBound(parts))
End Function

Private Function GetLastName(ByVal fullName As String) As String
    Dim parts() As String
    fullName = Trim$(fullName)
    If Len(fullName) = 0 Then Exit Function

    parts = Split(fullName, " ")
    GetLastName = parts(UBound(parts))
End Function

Private Function GetFullNameNoSpaces(ByVal fullName As String) As String
    GetFullNameNoSpaces = NormalizeUserKey(fullName)
End Function

Private Function GetFirstAndLastNameKey(ByVal fullName As String) As String
    Dim firstName As String
    Dim lastName As String

    firstName = GetFirstName(fullName)
    lastName = GetLastName(fullName)

    GetFirstAndLastNameKey = NormalizeUserKey(firstName & lastName)
End Function

Private Function UserMatches(ByVal idx As Integer, ByVal uname As String) As Boolean
    Dim key As String
    key = NormalizeUserKey(uname)

    If key = "" Then Exit Function

    ' 1. explicit username field
    If NormalizeUserKey(users(idx).uname) = key Then
        UserMatches = True
        Exit Function
    End If

    ' 2. short name alone
    If NormalizeUserKey(users(idx).ShortName) = key Then
        UserMatches = True
        Exit Function
    End If

    ' 3. full name collapsed
    If GetFullNameNoSpaces(users(idx).fullName) = key Then
        UserMatches = True
        Exit Function
    End If

    ' 4. first + last (important for Windows usernames)
    If GetFirstAndLastNameKey(users(idx).fullName) = key Then
        UserMatches = True
        Exit Function
    End If
End Function

Private Function GetCurrentUName() As String
    Dim u As String

    u = Environ$("USERNAME")   ' Windows
    If Len(u) = 0 Then u = Environ$("USER")      ' macOS
    If Len(u) = 0 Then u = Environ$("LOGNAME")   ' macOS fallback

    GetCurrentUName = Trim$(u)
End Function

Private Function GetUserIdx() As Integer
    EnsureLoaded

    Dim uname As String
    uname = GetCurrentUName()

    Dim i As Integer

    ' Pass 1: exact uname field first
    For i = LBound(users) To UBound(users)
        If NormalizeUserKey(users(i).uname) = NormalizeUserKey(uname) Then
            GetUserIdx = i
            Exit Function
        End If
    Next i

    ' Pass 2: broader matching across name fields
    For i = LBound(users) To UBound(users)
        If UserMatches(i, uname) Then
            GetUserIdx = i
            Exit Function
        End If
    Next i

    GetUserIdx = 0
End Function

Public Function GetFullName() As String
    GetFullName = users(GetUserIdx()).fullName
End Function

Public Function GetTitle() As String
    GetTitle = users(GetUserIdx()).title
End Function

Public Function GetMileage() As Double
    GetMileage = users(GetUserIdx()).Mileage
End Function

Public Function GetCity() As String
    GetCity = users(GetUserIdx()).City
End Function
