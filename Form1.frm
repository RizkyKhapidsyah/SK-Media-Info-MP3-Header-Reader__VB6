VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MTech Media Info"
   ClientHeight    =   3165
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5445
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5445
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Text            =   "Filename: ""NO CURRENTLY OPEN FILE"""
      Top             =   2760
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Media File Information"
      Height          =   2655
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5325
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2295
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   2160
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.CommandButton cmdTagSave 
         Caption         =   "Save Tag"
         Height          =   330
         Left            =   135
         TabIndex        =   16
         Top             =   1275
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdTagClear 
         Caption         =   "Clear Tag"
         Height          =   330
         Left            =   135
         TabIndex        =   15
         Top             =   802
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdMP3Hdr 
         Caption         =   "MP3 INFO"
         Height          =   330
         Left            =   150
         TabIndex        =   14
         Top             =   330
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2295
         TabIndex        =   12
         Text            =   "Text6"
         Top             =   2175
         Width           =   2610
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2295
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1785
         Width           =   2610
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2295
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   1410
         Width           =   2610
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2295
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   1035
         Width           =   2610
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2295
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   645
         Width           =   2610
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2295
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   270
         Width           =   2610
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Playing Time:"
         Height          =   195
         Left            =   1185
         TabIndex        =   6
         Top             =   2235
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Filesize:"
         Height          =   195
         Left            =   1575
         TabIndex        =   5
         Top             =   1845
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mode:"
         Height          =   195
         Left            =   1680
         TabIndex        =   4
         Top             =   1455
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Khz:"
         Height          =   195
         Left            =   1815
         TabIndex        =   3
         Top             =   1095
         Width           =   315
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Kbps:"
         Height          =   195
         Left            =   1725
         TabIndex        =   2
         Top             =   705
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bits:"
         Height          =   195
         Left            =   1830
         TabIndex        =   1
         Top             =   330
         Width           =   300
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   375
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Media File"
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open File"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ThisFile As String
Dim FileType As String

Private Sub cmdMP3Hdr_Click()
    ReadMP3Header (ThisFile)
    With MP3HeaderInfo
    Dim Result
    Result = MsgBox("Bitrate: " & .Bitrate & Chr(13) & "Emphasis: " & .Emphasis _
        & Chr(13) & "Frequency: " & .Frequency & Chr(13) & "Layer: " & _
        .Layer & Chr(13) & "Mode: " & .Mode & Chr(13) & "Mpeg Version: " & .MpegVersion & _
        Chr(13) & "Playing Time: " & .FPlayTime, vbOKOnly, "MP3 Header Info")
    End With
End Sub

Private Sub cmdTagClear_Click()
    If RemoveMP3Tag(ThisFile) = True Then ClearTag
End Sub

Private Sub cmdTagSave_Click()
    If WriteMP3Tag(ThisFile, "TAG", Text1, Text2, Text3, Text5, Text4, Combo1.ListIndex - 1) = True Then
    Beep
    Exit Sub
    End If
    Dim Result
    Result = MsgBox("ERROR Writing File!", vbCritical, "ERROR")
End Sub

'Code
Private Sub Form_Load()
Me.Refresh
Dim Genre
Dim i
    
    sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
    ' Build the Genre array (VB6+ only)
    Genre = Split(sGenreMatrix, "|")

    ClearText
    CommonDialog1.InitDir = GetSetting(App.Title, "Settings", "LastDir", "C:\")
    Me.Left = GetSetting(App.Title, "Settings", "Left", Me.Left)
    Me.Top = GetSetting(App.Title, "Settings", "Top", Me.Top)
    For i = 0 To 100
    Combo1.AddItem Genre(i)
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.Title, "Settings", "Left", Me.Left
SaveSetting App.Title, "Settings", "Top", Me.Top
End
End Sub

Private Sub mnuAbout_Click()
Dim About
About = MsgBox("MicroTech Media Info" & Chr(13) & "1999 - Freeware" & Chr(13) & "by Shannon Harmon" & Chr(13), vbInformation, "Info")
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuOpen_Click()
Dim Response
On Error GoTo ErrCheck
    'Open file and check for file type extension correctness
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = &H1000 + &H4
    CommonDialog1.Filter = "All Supported Media Types|*.mp3;*.mp2;*.wav|MPEG Audio (*.mp3, *.mp2)|*.mp3;*.mp2|FileWav File (*.wav)|*.wav"
    CommonDialog1.ShowOpen
    SaveSetting App.Title, "Settings", "LastDir", CommonDialog1.FileName
    CommonDialog1.InitDir = CommonDialog1.FileName
    ThisFile = CommonDialog1.FileName
    FileType = UCase(Right(ThisFile, 3))
    If FileType = "WAV" Then GetWaveData: Exit Sub
    If FileType = "MP3" Or FileType = "MP2" Then GetMP3Data: Exit Sub
    
    'If file type was not correct then prompt
    Response = MsgBox("Select a valid media file!", vbOKCancel, "Error")
    If Response = 1 Then mnuOpen_Click 'User choose ok - open again

ErrCheck:
'Just using with the cancel error
End Sub

Private Sub GetWaveData()
    If wHeadInfo(ThisFile) Then
    With wInfo
        Text1 = .bits
        Text2 = Round(Left(.kbps / 1000, 4), 0)
        If Len(Text2) = 4 Then Text2 = Left(Text2, 2) & "H"
        Text3 = .freq
        If .channels = 1 Then Text4 = "MONO"
        If .channels = 2 Then Text4 = "STEREO"
        Text5 = Format(.wFilesize, "###,###,###,###") & " Bytes"
        Text6 = .wPlaytime
        Text7 = ThisFile
        Label1.Caption = "Bits:"
        Label2.Caption = "Kbps:"
        Label3.Caption = "Khz:"
        Label4.Caption = "Mode:"
        Label5.Caption = "Filesize:"
        Label6.Caption = "Playing Time:"
        cmdMP3Hdr.Visible = False
        cmdTagSave.Visible = False
        cmdTagClear.Visible = False
        Text1.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
        Text6.Enabled = False
        Text1.Enabled = False
        Combo1.Visible = False
Exit Sub
    End With
    End If
    Response = MsgBox("Select a valid wavfile!", vbOKCancel, "Error")
    If Response = 1 Then mnuOpen_Click 'User choose ok - open again
    
End Sub

Private Sub GetMP3Data()
    If GetMP3Tag(ThisFile) = True Then
    With MP3Info
        Text1 = RTrim(.sTitle)
        Text2 = RTrim(.sArtist)
        Text3 = RTrim(.sAlbum)
        Text4 = RTrim(.sComment)
        Text5 = RTrim(.sYear)
        Combo1.Text = RTrim(.sGenre)
        Text7 = ThisFile
        Label1.Caption = "Title:"
        Label2.Caption = "Artist:"
        Label3.Caption = "Album:"
        Label4.Caption = "Comment:"
        Label5.Caption = "Year:"
        cmdMP3Hdr.Visible = True
        cmdTagSave.Visible = True
        cmdTagClear.Visible = True
        Text1.Enabled = True
        Text2.Enabled = True
        Text3.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
        Text6.Enabled = True
        Combo1.Visible = True
        Exit Sub
    End With
    End If
    Response = MsgBox("Select a valid mpeg file!", vbOKCancel, "Error")
    If Response = 1 Then mnuOpen_Click 'User choose ok - open again
End Sub

Private Sub ClearText()
    ClearTag
    Text7 = "Filename: NO CURRENTLY OPEN FILE"
    Label1.Caption = "Bits:"
    Label2.Caption = "Kbps:"
    Label3.Caption = "Khz:"
    Label4.Caption = "Mode:"
    Label5.Caption = "Filesize:"
    Label6.Caption = "Playing Time:"
    cmdMP3Hdr.Visible = False
    cmdTagSave.Visible = False
    cmdTagClear.Visible = False
    Combo1.Visible = False
End Sub

Private Sub ClearTag()
    Text1 = "": Text2 = "": Text3 = "": Text4 = "": Text5 = "": Text6 = ""
End Sub

