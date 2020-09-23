VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPL 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   5685
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstMusik 
      Height          =   2565
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   12632256
      BackColor       =   16576
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Song"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Duration"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbOptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      Height          =   195
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3120
      Width           =   540
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1290
      Picture         =   "frmPl.frx":0000
      Top             =   3120
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   2625
      Left            =   90
      Top             =   390
      Width           =   5565
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frmPl.frx":014A
      Top             =   3120
      Width           =   240
   End
   Begin VB.Label lbAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      Height          =   195
      Left            =   480
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3090
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PlayList"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2467
      TabIndex        =   0
      Top             =   30
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   350
      Left            =   -30
      Top             =   -30
      Width           =   58005
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   390
      Top             =   3030
      Width           =   525
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   60
      Top             =   3030
      Width           =   345
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   1560
      Top             =   3030
      Width           =   825
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   1230
      Top             =   3030
      Width           =   345
   End
   Begin VB.Menu popAdd 
      Caption         =   "popAdd"
      Visible         =   0   'False
      Begin VB.Menu mnuaddFiles 
         Caption         =   "+Agregar archivo(s)"
      End
      Begin VB.Menu div0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaddDir 
         Caption         =   "+Agregar directorio"
      End
   End
   Begin VB.Menu popOptions 
      Caption         =   "popOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuSort 
         Caption         =   "Sort"
         Begin VB.Menu submnuAZ 
            Caption         =   "A-Z"
         End
         Begin VB.Menu submnuZA 
            Caption         =   "Z-A"
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "File"
         Begin VB.Menu submnuInfFile 
            Caption         =   "[¤] Info"
         End
         Begin VB.Menu div1 
            Caption         =   "-"
         End
         Begin VB.Menu submnudel 
            Caption         =   "[†] Del Entry"
         End
      End
   End
End
Attribute VB_Name = "frmPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type tSongData
    ArtistTitle As String * 100
    Seconds As Integer
    Path As String * 255
End Type

Private Sub Form_Load()
Dim tSD As tSongData, xItem As ListItem
    lbAdd.MouseIcon = LoadResPicture("HAND", vbResCursor)
    lbOptions.MouseIcon = LoadResPicture("HAND", vbResCursor)
    With lstMusik
        .ColumnHeaders(1).Width = .Width * 75 / 100
        .ColumnHeaders(2).Width = .Width * 25 / 100
    End With
    ''Read Playlist
    Dim I&, nFil&
    nFil = FreeFile
    Open "C:\proplayer.pls" For Random As nFil Len = Len(tSD)
        With lstMusik.ListItems
            Do While Not EOF(nFil)
                Get nFil, , tSD
                If tSD.Seconds > 0 Then 'No add corrupted files
                    Set xItem = lstMusik.ListItems.Add(, , Trim(tSD.ArtistTitle))
                    xItem.SubItems(1) = ConvertSeconds(tSD.Seconds)
                    xItem.TaG = tSD.Seconds
                    xItem.Key = tSD.Path
                End If
            Loop
        End With
    
    Close nFil
End Sub

Private Sub Form_Unload(Cancel As Integer)
'''Save Playlist
Dim I&, nFil&, tSD As tSongData
    nFil = FreeFile
    Open "C:\proplayer.pls" For Random As nFil Len = Len(tSD)
        With lstMusik.ListItems
            For I = 1 To .Count
                tSD.ArtistTitle = .Item(I)
                tSD.Seconds = CInt(.Item(I).TaG)
                tSD.Path = .Item(I).Key
                Put nFil, , tSD
            Next
        End With
    
    Close nFil
End Sub

Private Sub lbAdd_Click()
    PopupMenu popAdd, , Shape7.Left, Shape8.Top - 700
End Sub

Private Sub lbOptions_Click()
    PopupMenu popOptions, , Shape4.Left, Shape3.Top - 600
End Sub

Private Sub lstMusik_DblClick()
'
    SetSong lstMusik.SelectedItem.Index
    frmMain.PlayerActions ePlay
End Sub

Function SetSong(Pos&)
If lstMusik.ListItems.Count Then
    frmMain.LoadAudio lstMusik.ListItems(Pos).Key
    CurSong = Pos
    lstMusik.ListItems(Pos).Selected = True
End If
    
End Function

Private Sub mnuaddDir_Click()
On Error GoTo Cancel
Dim BrowFol As Shell
Dim SelCarpeta As Folder
    Set BrowFol = New Shell
    Set SelCarpeta = BrowFol.BrowseForFolder(Me.hwnd, "Seleccion la Carpeta de Musica", 0, "C:\")
    ReadDir (SelCarpeta.Items.Item.Path)
Cancel:
    Set BrowFol = Nothing
    Set SelCarpeta = Nothing
    
End Sub

Function ReadDir(Path$)
Dim sFile$, vFiles() As String, n&
    Path = AddSlash(Path)
    sFile = Dir(Path & "*.mp3")
    Do While sFile <> ""
        n = n + 1
        ReDim Preserve vFiles(1 To n)
        vFiles(n) = Path & sFile
        sFile = Dir
    Loop
    BloquearControl lstMusik.hwnd, True
    For n = 1 To UBound(vFiles)
        AddFile vFiles(n)
    Next
    BloquearControl lstMusik.hwnd, False
    lstMusik.Refresh
'
End Function

Private Sub mnuaddFiles_Click()
On Error Resume Next
    frmMain.CmD.ShowOpen
    If Err.Number <> 32755 Then AddFile frmMain.CmD.Filename
    Err.Clear
End Sub

Function AddFile(FullPath$)
Dim cIdTag As New cMpeg_Id3v1Tag
Dim newLstIt As ListItem
Dim sArtist$, sTitle$
On Error Resume Next

    cIdTag.ArchivoMp3 = FullPath
    Err.Clear
    With lstMusik
        sArtist = IIf((cIdTag.Artista = ""), "NoArtist", cIdTag.Artista)
        sTitle = IIf((cIdTag.Tema = ""), "NoTitle", cIdTag.Tema)
        Set newLstIt = .ListItems.Add(, FullPath, Trim(sArtist) & " - " & Trim(sTitle))
        If (Err.Number = 0) And (cIdTag.Segun2 > 0) Then 'No add corrupted files and duplicated
            newLstIt.SubItems(1) = ConvertSeconds(cIdTag.Segun2)
            newLstIt.TaG = cIdTag.Segun2
        End If
    End With
    
End Function

Private Sub submnuAZ_Click()
    BloquearControl lstMusik.hwnd, True
    lstMusik.SortOrder = lvwAscending
    lstMusik.Sorted = True
    BloquearControl lstMusik.hwnd, False
End Sub

Private Sub submnuInfFile_Click()
    If CurSong > 0 Then
        frmFileInfo.Show
        frmFileInfo.TaG = lstMusik.ListItems(CurSong).Key
    End If
End Sub

Private Sub submnuZA_Click()
    BloquearControl lstMusik.hwnd, True
    lstMusik.SortOrder = lvwDescending
    lstMusik.Sorted = True
    BloquearControl lstMusik.hwnd, False
End Sub
