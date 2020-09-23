VERSION 5.00
Begin VB.Form frmFileInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Info"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
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
   ScaleHeight     =   5745
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Undo Changes"
      Height          =   345
      Left            =   3510
      TabIndex        =   18
      Top             =   5250
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1320
      TabIndex        =   17
      Top             =   5250
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   345
      Left            =   150
      TabIndex        =   16
      Top             =   5250
      Width           =   1095
   End
   Begin VB.Frame frmMPEG 
      Caption         =   "MPEG Info"
      Height          =   2355
      Left            =   120
      TabIndex        =   14
      Top             =   2790
      Width           =   4815
      Begin VB.Label lbMPEGInfo 
         Caption         =   "MPEG INFO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   4395
      End
   End
   Begin VB.Frame frmIdTag 
      Caption         =   "ID3 Tag"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   510
      Width           =   4785
      Begin VB.ComboBox cboGenre 
         Height          =   315
         Left            =   2220
         TabIndex        =   13
         Top             =   1350
         Width           =   1425
      End
      Begin VB.TextBox txtYear 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   11
         Top             =   1380
         Width           =   675
      End
      Begin VB.TextBox txtComment 
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Top             =   1740
         Width           =   2805
      End
      Begin VB.TextBox txtAlbum 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   1020
         Width           =   2805
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   660
         Width           =   2805
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   300
         Width           =   2805
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Genre"
         Height          =   195
         Left            =   1710
         TabIndex        =   12
         Top             =   1410
         Width           =   435
      End
      Begin VB.Label lbYear 
         Alignment       =   1  'Right Justify
         Caption         =   "Year"
         Height          =   195
         Left            =   450
         TabIndex        =   10
         Top             =   1380
         Width           =   330
      End
      Begin VB.Label lbComment 
         Alignment       =   1  'Right Justify
         Caption         =   "Comment"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   1770
         Width           =   705
      End
      Begin VB.Label lbAlbum 
         Alignment       =   1  'Right Justify
         Caption         =   "Album"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Artist"
         Height          =   195
         Left            =   390
         TabIndex        =   4
         Top             =   660
         Width           =   390
      End
      Begin VB.Label lbTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Title"
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   330
         Width           =   300
      End
   End
   Begin VB.Label lbPath 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Path"
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   4770
   End
End
Attribute VB_Name = "frmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cID3 As New cMpeg_Id3v1Tag

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    cID3.ArchivoMp3 = Me.TaG
    txtArtist = cID3.Artista
    txtTitle = cID3.Tema
    txtAlbum = cID3.Album
    txtYear = cID3.Year
    txtComment = cID3.Comments
    lbMPEGInfo = "Size:" & cID3.TamaMp3 & vbCrLf
    lbMPEGInfo = lbMPEGInfo & "Header:" & cID3.Header & vbCrLf
    lbMPEGInfo = lbMPEGInfo & "Length:" & cID3.Segun2 & vbCrLf
    lbMPEGInfo = lbMPEGInfo & cID3.VersionLayer & vbCrLf
    lbMPEGInfo = lbMPEGInfo & cID3.BitRate & "kbits," & cID3.Frames & " frames" & vbCrLf
    lbMPEGInfo = lbMPEGInfo & cID3.Frecuencia & "Hz " & cID3.Mode
    lbPath = Me.TaG
End Sub
