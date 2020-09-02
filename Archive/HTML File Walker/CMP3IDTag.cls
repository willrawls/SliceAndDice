VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMP3IDTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Filename As String
Public Complete As String
Public Title As String
Public Artist As String
Public Album As String
Public Year As Integer
Public Comment As String
Public IsValid As Boolean
Public HasData As Boolean

Public Enum enumGenre
    [Unknown] = 0
    [Classic Rock] = 1
    [Country] = 2
    [Dance] = 3
    [Disco] = 4
    [Funk] = 5
    [Grunge] = 6
    [Hip-Hop] = 7
    [Jazz] = 8
    [Metal] = 9
    [New Age] = 10
    [Oldies] = 11
    [Other] = 12
    [Pop] = 13
    [RandB] = 14
    [Rap] = 15
    [Reggae] = 16
    [Rock] = 17
    [Techno] = 18
    [Industrial] = 19
    [Alternative] = 20
    [Ska] = 21
    [Death Metal] = 22
    [Pranks] = 23
    [Soundtrack] = 24
    [Euro-Techno] = 25
    [Ambient] = 26
    [Trip-Hop] = 27
    [Vocal] = 28
    [Jazz+Funk] = 29
    [Fusion] = 30
    [Trance] = 31
    [Classical] = 32
    [Instrumental] = 33
    [Acid] = 34
    [House] = 35
    [Game] = 36
    [Sound Clip] = 37
    [Gospel] = 38
    [Noise] = 39
    [AlternRock] = 40
    [Bass] = 41
    [Soul] = 42
    [Punk] = 43
    [Space] = 44
    [Meditative] = 45
    [Instrumental Pop] = 46
    [Instrumental Rock] = 47
    [Ethnic] = 48
    [Gothic] = 49
    [Darkwave] = 50
    [Techno-Industrial] = 51
    [Electronic] = 52
    [Pop-Folk] = 53
    [Eurodance] = 54
    [Dream] = 55
    [Southern Rock] = 56
    [Comedy] = 57
    [Cult] = 58
    [Gangsta] = 59
    [Top 40] = 60
    [Christian Rap] = 61
    [Pop Funk] = 62
    [Jungle] = 63
    [Native American] = 64
    [Cabaret] = 65
    [New Wave] = 66
    [Psychadelic] = 67
    [Rave] = 68
    [Showtunes] = 69
    [Trailer] = 70
    [Lo Fi] = 71
    [Tribal] = 72
    [Acid Punk] = 73
    [Acid Jazz] = 74
    [Polka] = 75
    [Retro] = 76
    [Musical] = 77
    [Rock and Roll] = 78
    [Hard Rock] = 79
    [TooHigh] = 80
End Enum
Public Genre As enumGenre

Public Sub Read(ByVal sFilename As String)
    Dim fh As Integer
    Dim lFilename As Long

    fh = FreeFile
    Filename = sFilename
    lFilename = FileLen(sFilename)
    Open sFilename For Binary Access Read As #fh
         Seek #fh, lFilename - 127
         Complete = Input(128, fh)
         Year = Val(Trim$(Mid$(Complete, 94, 4)))
         If Left$(Complete, 3) = "TAG" Then
            IsValid = True
            Title = Trim$(Mid$(Complete, 4, 30))
            Artist = Trim$(Mid$(Complete, 34, 30))
            Album = Trim$(Mid$(Complete, 64, 30))
            Comment = Trim$(Mid$(Complete, 98, 30))
            HasData = (Len(Title) > 0 Or Len(Artist) > 0 Or Len(Album) > 0 Or Len(Comment) > 0)
            If Asc(Right$(Complete, 1)) < TooHigh Then
               Genre = Asc(Right$(Complete, 1))
            Else
               Genre = Unknown
            End If
         Else
            IsValid = False
            HasData = False
            Genre = Unknown
            Title = vbNullString
            Artist = vbNullString
            Album = vbNullString
            Comment = vbNullString
         End If
    Close #fh
End Sub

Public Sub Save()
    If Not IsValid Then Exit Sub

    Dim fh As Integer
    Dim lFilename As Long

    fh = FreeFile
    lFilename = FileLen(Filename)
    Open Filename For Binary Access Write As #fh
         Seek #fh, lFilename - 127
         Complete = "TAG" & Spaces(30 - Len(Title)) & Artist & Spaces(30 - Len(Artist)) & Album & Spaces(30 - Len(Album)) & Format(Year, "0000") & Comment & Spaces(30 - Len(Comment)) & Chr$(Genre)
           ' 0..2    == 'TAG' (3 Bytes)
           ' 3..32   == SongName (30 bytes)
           ' 33..62  == Artist (30 Bytes)
           ' 63..92  == Album Name (30 Bytes)
           ' 93..96  == Year (4 Bytes)
           ' 97..126 == Comment (30 Bytes)
           ' 127     == 1 Byte Song Genre Identifier
         If Len(Complete) = 128 Then
            Print #fh, Complete
         End If
    Close #fh
End Sub

Public Property Get sGenre() As String
    Select Case Genre
           Case [Unknown]: sGenre = "Unknown"
           Case [Classic Rock]: sGenre = "Classic Rock"
           Case [Country]: sGenre = "Country"
           Case [Dance]: sGenre = "Dance"
           Case [Disco]: sGenre = "Disco"
           Case [Funk]: sGenre = "Funk"
           Case [Grunge]: sGenre = "Grunge"
           Case [Hip-Hop]: sGenre = "Hip-Hop"
           Case [Jazz]: sGenre = "Jazz"
           Case [Metal]: sGenre = "Metal"
           Case [New Age]: sGenre = "New Age"
           Case [Oldies]: sGenre = "Oldies"
           Case [Other]: sGenre = "Other"
           Case [Pop]: sGenre = "Pop"
           Case [RandB]: sGenre = "R&B"
           Case [Rap]: sGenre = "Rap"
           Case [Reggae]: sGenre = "Reggae"
           Case [Rock]: sGenre = "Rock"
           Case [Techno]: sGenre = "Techno"
           Case [Industrial]: sGenre = "Industrial"
           Case [Alternative]: sGenre = "Alternative"
           Case [Ska]: sGenre = "Ska"
           Case [Death Metal]: sGenre = "Death Metal"
           Case [Pranks]: sGenre = "Pranks"
           Case [Soundtrack]: sGenre = "Soundtrack"
           Case [Euro-Techno]: sGenre = "Euro-Techno"
           Case [Ambient]: sGenre = "Ambient"
           Case [Trip-Hop]: sGenre = "Trip-Hop"
           Case [Vocal]: sGenre = "Vocal"
           Case [Jazz+Funk]: sGenre = "Jazz+Funk"
           Case [Fusion]: sGenre = "Fusion"
           Case [Trance]: sGenre = "Trance"
           Case [Classical]: sGenre = "Classical"
           Case [Instrumental]: sGenre = "Instrumental"
           Case [Acid]: sGenre = "Acid"
           Case [House]: sGenre = "House"
           Case [Game]: sGenre = "Game"
           Case [Sound Clip]: sGenre = "Sound Clip"
           Case [Gospel]: sGenre = "Gospel"
           Case [Noise]: sGenre = "Noise"
           Case [AlternRock]: sGenre = "AlternRock"
           Case [Bass]: sGenre = "Bass"
           Case [Soul]: sGenre = "Soul"
           Case [Punk]: sGenre = "Punk"
           Case [Space]: sGenre = "Space"
           Case [Meditative]: sGenre = "Meditative"
           Case [Instrumental Pop]: sGenre = "Instrumental Pop"
           Case [Instrumental Rock]: sGenre = "Instrumental Rock"
           Case [Ethnic]: sGenre = "Ethnic"
           Case [Gothic]: sGenre = "Gothic"
           Case [Darkwave]: sGenre = "Darkwave"
           Case [Techno-Industrial]: sGenre = "Techno-Industrial"
           Case [Electronic]: sGenre = "Electronic"
           Case [Pop-Folk]: sGenre = "Pop-Folk"
           Case [Eurodance]: sGenre = "Eurodance"
           Case [Dream]: sGenre = "Dream"
           Case [Southern Rock]: sGenre = "Southern Rock"
           Case [Comedy]: sGenre = "Comedy"
           Case [Cult]: sGenre = "Cult"
           Case [Gangsta]: sGenre = "Gangsta"
           Case [Top 40]: sGenre = "Top 40"
           Case [Christian Rap]: sGenre = "Christian Rap"
           Case [Pop Funk]: sGenre = "Pop/Funk"
           Case [Jungle]: sGenre = "Jungle"
           Case [Native American]: sGenre = "Native American"
           Case [Cabaret]: sGenre = "Cabaret"
           Case [New Wave]: sGenre = "New Wave"
           Case [Psychadelic]: sGenre = "Psychadelic"
           Case [Rave]: sGenre = "Rave"
           Case [Showtunes]: sGenre = "Showtunes"
           Case [Trailer]: sGenre = "Trailer"
           Case [Lo Fi]: sGenre = "Lo-Fi"
           Case [Tribal]: sGenre = "Tribal"
           Case [Acid Punk]: sGenre = "Acid Punk"
           Case [Acid Jazz]: sGenre = "Acid Jazz"
           Case [Polka]: sGenre = "Polka"
           Case [Retro]: sGenre = "Retro"
           Case [Musical]: sGenre = "Musical"
           Case [Rock and Roll]: sGenre = "Rock & Roll"
           Case [Hard Rock]: sGenre = "Hard Rock"
    End Select
End Property


Private Function Spaces(ByVal nNumber As Integer) As String
    Dim i As Integer
    For i = 1 To nNumber
        Spaces = Spaces & " "
    Next i
End Function


