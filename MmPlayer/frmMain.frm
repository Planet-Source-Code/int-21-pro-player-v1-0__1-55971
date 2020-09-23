VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   158
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   381
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ProPlayer.xVis xVis 
      Height          =   660
      Left            =   90
      TabIndex        =   17
      Top             =   390
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   1164
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   2850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":015A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Title 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   -150
      ScaleHeight     =   255
      ScaleWidth      =   6615
      TabIndex        =   4
      Top             =   0
      Width           =   6645
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[X]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5460
         TabIndex        =   16
         Top             =   30
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pro-Player"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2430
         TabIndex        =   15
         Top             =   30
         Width           =   900
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Pl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6090
      TabIndex        =   3
      Top             =   2250
      Width           =   765
   End
   Begin VB.PictureBox PicTimers 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   3750
      ScaleHeight     =   765
      ScaleWidth      =   1725
      TabIndex        =   2
      Top             =   330
      Width           =   1755
   End
   Begin VB.Timer tmPlayer 
      Interval        =   120
      Left            =   6990
      Top             =   1650
   End
   Begin ProPlayer.SliderBar Cue 
      Height          =   225
      Left            =   270
      TabIndex        =   1
      Top             =   1380
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   397
      Posicion        =   0
      ImgFondo        =   "frmMain.frx":02B4
      ImgCue          =   "frmMain.frx":02D0
      CueColor        =   32768
      BckColor        =   8421504
   End
   Begin MSComDlg.CommonDialog CmD 
      Left            =   7020
      Top             =   990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Archivos de Audio|*.mp3;*.wav"
   End
   Begin VB.PictureBox PicInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   1620
      ScaleHeight     =   765
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   330
      Width           =   2145
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H00C0C0C0&
      Height          =   795
      Left            =   60
      Top             =   330
      Width           =   1545
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   4650
      Picture         =   "frmMain.frx":02EC
      Top             =   1830
      Width           =   240
   End
   Begin VB.Label lbOpen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      Height          =   195
      Left            =   4980
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   1860
      Width           =   510
   End
   Begin VB.Label lbPlayList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Playlist"
      Height          =   195
      Left            =   4980
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4650
      Picture         =   "frmMain.frx":0436
      Top             =   1410
      Width           =   240
   End
   Begin VB.Label lbNext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      Height          =   195
      Left            =   3780
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   1890
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3390
      TabIndex        =   8
      Top             =   1800
      Width           =   285
   End
   Begin VB.Label lbBck 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      Height          =   195
      Left            =   2850
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1890
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2460
      TabIndex        =   9
      Top             =   1800
      Width           =   285
   End
   Begin VB.Shape Shape10 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   4920
      Top             =   1350
      Width           =   645
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   4590
      Top             =   1350
      Width           =   345
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   3690
      Top             =   1830
      Width           =   525
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   3360
      Top             =   1830
      Width           =   345
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   2760
      Top             =   1830
      Width           =   525
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   2430
      Top             =   1830
      Width           =   345
   End
   Begin VB.Label lbStop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      Height          =   195
      Left            =   1890
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1890
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1560
      TabIndex        =   7
      Top             =   1800
      Width           =   285
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   1860
      Top             =   1830
      Width           =   525
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   1530
      Top             =   1830
      Width           =   345
   End
   Begin VB.Label lbPlay 
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      Height          =   195
      Left            =   810
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Tag             =   "Paused"
      Top             =   1890
      Width           =   540
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   690
      Top             =   1830
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   390
      TabIndex        =   5
      Top             =   1830
      Width           =   285
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   390
      Top             =   1830
      Width           =   345
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   4590
      Top             =   1770
      Width           =   345
   End
   Begin VB.Shape Shape11 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   4920
      Top             =   1770
      Width           =   645
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Pro-Player v1.0
'This is my 1st post about this player...i going to post more new stuff later
' i just wanna ur feedback about my code
'this code use FilgraphManager Object, and wave Apis to read sound data
'sorry by my english..im not englishman
'Enjoy

Dim wmP As FilgraphManager, wmPos As IMediaPosition, wmAu As IBasicAudio
Dim PlOn As Boolean
Dim Playing As Boolean

Private Sub Form_Load()
    Show
    lbOpen.MouseIcon = LoadResPicture("HAND", vbResCursor)
    lbPlay.MouseIcon = LoadResPicture("HAND", vbResCursor)
    lbStop.MouseIcon = LoadResPicture("HAND", vbResCursor)
    lbBck.MouseIcon = LoadResPicture("HAND", vbResCursor)
    lbNext.MouseIcon = LoadResPicture("HAND", vbResCursor)
    lbPlayList.MouseIcon = LoadResPicture("HAND", vbResCursor)
    
    Call lbPlayList_Click
    xVis.Init
End Sub

Private Sub lbBck_Click()
    LoadSong eBack
End Sub

Private Sub lbNext_Click()
    LoadSong eNext
End Sub

Private Sub lbPlay_Click()
If Not wmP Is Nothing Then
    If lbPlay.TaG = "Playing" Then
        PlayerActions 2
    ElseIf lbPlay.TaG = "Paused" Then
        PlayerActions 1
    End If
Else
    Call lbOpen_Click
End If
End Sub

Private Sub Cue_CambioPosicion(NuevaPosicion As Long)
If Not wmP Is Nothing Then wmPos.CurrentPosition = NuevaPosicion
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set wmP = Nothing
    Set wmPos = Nothing
    Set wmAu = Nothing
    Unload frmPL
    xVis.Vis = False
    xVis.Finished
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub

Private Sub lbOpen_Click()
Dim cIdTag As New cMpeg_Id3v1Tag
Dim newLstIt As ListItem
On Error Resume Next
    CmD.ShowOpen
    If Err.Number <> 32755 Then
        
        Set cIdTag = LoadAudio(CmD.Filename)
        
        With frmPL.lstMusik
            Set newLstIt = .ListItems.Add(, , cIdTag.Artista & " - " & cIdTag.Tema)
            newLstIt.SubItems(1) = ConvertSeconds(cIdTag.Segun2)
            newLstIt.TaG = CmD.Filename
        End With
        
    End If
    Err.Clear
End Sub

Private Sub lbPlayList_Click()
    If Not PlOn Then
        frmPL.Show
        frmPL.Top = frmMain.Top + frmMain.Height
        frmPL.Left = frmMain.Left
        PlOn = True
    Else
        Unload frmPL
        PlOn = False
    End If
End Sub

Private Sub lbStop_Click()
    PlayerActions 3
End Sub

Private Sub Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        ReleaseCapture
        Call SendMessage(frmMain.hWnd, &HA1, 2, 0&)
        If PlOn Then
            frmPL.Top = frmMain.Top + frmMain.Height
            frmPL.Left = frmMain.Left
        End If
End If
End Sub

Private Sub tmPlayer_Timer()
Dim lState&, sTimers$
If Playing Then
    If Not wmP Is Nothing Then
        PicTimers.Cls
        sTimers = "Avance:" & ConvertSeconds(wmPos.CurrentPosition) & vbCrLf
        sTimers = sTimers & "Resta:" & ConvertSeconds(wmPos.Duration - wmPos.CurrentPosition) & vbCrLf
        sTimers = sTimers & "Total:" & ConvertSeconds(wmPos.Duration)
        PicTimers.Print sTimers
        Cue.Posicion = wmPos.CurrentPosition
        If wmPos.CurrentPosition >= wmPos.Duration Then
            If PlOn Then
                CurSong = CurSong + 1
                frmPL.SetSong CurSong
                PlayerActions ePlay
            End If
        End If
    End If
Else
    tmPlayer.Enabled = False
End If
End Sub

Function LoadAudio(sAudio As String) As cMpeg_Id3v1Tag
Dim cIdTag As New cMpeg_Id3v1Tag
Dim sInfo$
    
    cIdTag.ArchivoMp3 = sAudio
    
    Set wmP = New FilgraphManager
    wmP.RenderFile sAudio
    Set wmPos = wmP
    Set wmAu = wmP
    
    PicInfo.Cls
    sInfo = cIdTag.Artista & vbCrLf
    sInfo = sInfo & cIdTag.Tema & vbCrLf
    sInfo = sInfo & "Kbit:" & cIdTag.BitRate & vbCrLf
    sInfo = sInfo & "Duracion:" & ConvertSeconds(cIdTag.Segun2) & vbCrLf
    sInfo = sInfo & "Hz:" & cIdTag.Frecuencia & vbCrLf
    PicInfo.Print sInfo
    Cue.Maximo = wmPos.Duration
    Cue.Minimo = 0
    
    'Call lbPlay_Click
Set LoadAudio = cIdTag
End Function

Function LoadSong(Move As eDirection)
Dim TotSongs&
    TotSongs = frmPL.lstMusik.ListItems.Count
    
    If TotSongs Then
        If Move = eBack Then
            CurSong = CurSong - 1
        Else
            CurSong = CurSong + 1
        End If
        If CurSong > TotSongs Then CurSong = TotSongs
        If CurSong < 1 Then CurSong = 1
    End If
    frmPL.SetSong CurSong
    lbPlay_Click
End Function

Function PlayerActions(Action As eActions)
    If Not wmP Is Nothing Then
        Select Case Action
            Case ePlay
                lbPlay.TaG = "Playing"
                lbPlay.Caption = "Play"
                wmP.Run
                tmPlayer.Enabled = True
                Playing = True
                xVis.Vis = True
            Case ePaused
                lbPlay.TaG = "Paused"
                lbPlay.Caption = "Paused"
                wmP.Pause
                xVis.Vis = False
            Case eStop
                wmP.StopWhenReady
                Playing = False
                xVis.Vis = False
        End Select
    End If
End Function
