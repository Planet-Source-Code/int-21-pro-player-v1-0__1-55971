VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clId3Tag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type MP3Info
  BitRate As String
  Channels As String
  COPYRIGHT As String
  CRC As String
  EMPHASIS As String
  Freq As String
  LAYER As String
  LENGTH As String
  MPEG As String
  ORIGINAL As String
  SIZE As String
End Type

Private Type VBRinfo
  VBRrate As String
  VBRlength As String
End Type

Dim MP3Size&, MP3Length&
Dim MP3File$
Dim mTitulo$, mArtista$, mAlbum$, mYear%, mComentario$, mGenero%
Dim mSize&, mLength&, mLayer$, mBitRate$, mFreqChan$, mCRC$, mCopyright$, mEmphasis$, mOriginal$
Function LeerMp3Info(RutaMp3$)
Dim Res$
    MP3File = RutaMp3
    Res = Dir(MP3File)
    If Len(Res) > 0 Then
        GetTagInf
        GetMP3Inf
    End If
End Function

Private Sub GetTagInf()
On Local Error Resume Next
  Dim Buf As String * 128
  Dim tmpStr As String
  Dim I As Byte
  
  'Mp3File = Mp3Filename
  'Get the size of mp3 file(in bytes)
  MP3Size = FileLen(MP3File)
  
  'labLength = labLength & mp3Length & " seconds"
  
  'Open the file for binary access in order to get the ID3 Tag
  Open MP3File For Binary As #1
    'Get last 128 bytes of the file. The size of file is reduced by 127 bytes, because
    'the last byte in file is in fact the size of file
    Get #1, MP3Size - 127, Buf
    'Check if the file has a tag
    If format(Left(Buf, 3), "<") <> "tag" Then
        Close #1
        Exit Sub
    Else
      'If it has a tag the separate the info obtained in the buffer string
      mTitulo = AscChar(Trim(Mid(Buf, 4, 30)))
      mArtista = AscChar(Trim(Mid(Buf, 34, 30)))
      mAlbum = AscChar(Trim(Mid(Buf, 64, 30)))
      mYear = AscChar(Trim(Mid(Buf, 94, 4)))
      mComentario = AscChar((Trim(Mid(Buf, 98, 30))))
      mGenero = AscChar(Trim(Asc(Mid$(Buf, 128, 1))))
    End If
  Close #1
End Sub

Private Sub GetMP3Inf()
  Dim accMP3Info As MP3Info
  
  getMP3Info MP3File, accMP3Info
  
  mSize = accMP3Info.SIZE
  mLength = accMP3Info.LENGTH
  mLayer = accMP3Info.MPEG & " " & accMP3Info.LAYER
  mBitRate = accMP3Info.BitRate
  mFreqChan = accMP3Info.Freq & " " & accMP3Info.Channels
  mCRC = accMP3Info.CRC
  mCopyright = accMP3Info.COPYRIGHT
  mEmphasis = accMP3Info.EMPHASIS
  mOriginal = accMP3Info.ORIGINAL
End Sub

Public Property Get Titulo() As String
    Titulo = mTitulo
End Property

Public Property Get Artista() As String
    Artista = mArtista
End Property
Public Property Get Album() As String
    Album = mAlbum
End Property
Public Property Get Year() As String
    Year = mYear
End Property
Public Property Get Comentario() As String
    Comentario = mComentario
End Property
Public Property Get Genero() As String
    Genero = mGenero
End Property
Public Property Get SIZE() As Long
    SIZE = mSize
End Property

Public Property Get LENGTH() As Long
    LENGTH = mLength
End Property

Public Property Get LAYER() As String
    LAYER = mLayer
End Property

Public Property Get BitRate() As String
    BitRate = mBitRate
End Property

Public Property Get FreqCanales() As String
    FreqCanales = mFreqChan
End Property

Public Property Get CRC() As String
    CRC = mCRC
End Property

Public Property Get COPYRIGHT() As String
    COPYRIGHT = mCopyright
End Property

Public Property Get EMPHASIS() As String
    EMPHASIS = mEmphasis
End Property

Public Property Get ORIGINAL() As String
    ORIGINAL = mOriginal
End Property

Private Sub getMP3Info(ByVal lpMP3File As String, ByRef lpMP3Info As MP3Info)
  Dim Buf As String * 4096
  Dim infoStr As String * 3
  Dim lpVBRinfo As VBRinfo
  Dim tmpByte As Byte
  Dim tmpNum As Byte
  Dim I As Integer
  Dim designator As Byte
  Dim baseFreq As Single
  Dim vbrBytes As Long
  
  Open lpMP3File For Binary As #1
    Get #1, 1, Buf
  Close #1
  
  For I = 1 To 4092
    If Asc(Mid(Buf, I, 1)) = &HFF Then
      tmpByte = Asc(Mid(Buf, I + 1, 1))
      If Between(tmpByte, &HF2, &HF7) Or Between(tmpByte, &HFA, &HFF) Then
        Exit For
      End If
    End If
  Next I
  If I = 4093 Then
    'MsgBox "Not a MP3 file...", vbCritical, "Error..."
  Else
    infoStr = Mid(Buf, I + 1, 3)
    'Getting info from 2nd byte(MPEG,Layer type and CRC)
    tmpByte = Asc(Mid(infoStr, 1, 1))
    
    'Getting CRC info
    If ((tmpByte Mod 16) Mod 2) = 0 Then
      lpMP3Info.CRC = "Yes"
    Else
      lpMP3Info.CRC = "No"
    End If
    
    'Getting MPEG type info
    If Between(tmpByte, &HF2, &HF7) Then
      lpMP3Info.MPEG = "MPEG 2.0"
      designator = 1
    Else
      lpMP3Info.MPEG = "MPEG 1.0"
      designator = 2
    End If
    
    'Getting layer info
    If Between(tmpByte, &HF2, &HF3) Or Between(tmpByte, &HFA, &HFB) Then
      lpMP3Info.LAYER = "layer 3"
    Else
      If Between(tmpByte, &HF4, &HF5) Or Between(tmpByte, &HFC, &HFD) Then
        lpMP3Info.LAYER = "layer 2"
      Else
        lpMP3Info.LAYER = "layer 1"
      End If
    End If
    
    'Getting info from 3rd byte(Frequency, Bit-rate)
    tmpByte = Asc(Mid(infoStr, 2, 1))
    
    'Getting frequency info
    If Between(tmpByte Mod 16, &H0, &H3) Then
      baseFreq = 22.05
    Else
      If Between(tmpByte Mod 16, &H4, &H7) Then
        baseFreq = 24
      Else
        baseFreq = 16
      End If
    End If
    lpMP3Info.Freq = baseFreq * designator * 1000 & " Hz"
    
    'Getting Bit-rate
    tmpNum = tmpByte \ 16 Mod 16
    If designator = 1 Then
      If tmpNum < &H8 Then
        lpMP3Info.BitRate = tmpNum * 8
      Else
        lpMP3Info.BitRate = 64 + (tmpNum - 8) * 16
      End If
    Else
      If tmpNum <= &H5 Then
        lpMP3Info.BitRate = (tmpNum + 3) * 8
      Else
        If tmpNum <= &H9 Then
          lpMP3Info.BitRate = 64 + (tmpNum - 5) * 16
        Else
          If tmpNum <= &HD Then
            lpMP3Info.BitRate = 128 + (tmpNum - 9) * 32
          Else
            lpMP3Info.BitRate = 320
          End If
        End If
      End If
    End If
    MP3Length = FileLen(lpMP3File) \ (Val(lpMP3Info.BitRate) / 8) \ 1000
    If Mid(Buf, I + 36, 4) = "Xing" Then
      vbrBytes = Asc(Mid(Buf, I + 45, 1)) * &H10000
      vbrBytes = vbrBytes + (Asc(Mid(Buf, I + 46, 1)) * &H100&)
      vbrBytes = vbrBytes + Asc(Mid(Buf, I + 47, 1))
      GetVBRrate lpMP3File, vbrBytes, lpVBRinfo
      lpMP3Info.BitRate = lpVBRinfo.VBRrate
      lpMP3Info.LENGTH = lpVBRinfo.VBRlength
    Else
      lpMP3Info.BitRate = lpMP3Info.BitRate & "Kbit"
      lpMP3Info.LENGTH = MP3Length '& " seconds"
    End If
    
    'Getting info from 4th byte(Original, Emphasis, Copyright, Channels)
    tmpByte = Asc(Mid(infoStr, 3, 1))
    tmpNum = tmpByte Mod 16
    
    
    'Getting Copyright bit
    If tmpNum \ 8 = 1 Then
      lpMP3Info.COPYRIGHT = " Yes"
      tmpNum = tmpNum - 8
    Else
      lpMP3Info.COPYRIGHT = " No"
    End If
    
    'Getting Original bit
    If (tmpNum \ 4) Mod 2 Then
      lpMP3Info.ORIGINAL = " Yes"
      tmpNum = tmpNum - 4
    Else
      lpMP3Info.ORIGINAL = " No"
    End If
    
    'Getting Emphasis bit
    Select Case tmpNum
      Case 0
        lpMP3Info.EMPHASIS = " None"
      Case 1
        lpMP3Info.EMPHASIS = " 50/15 microsec"
      Case 2
        lpMP3Info.EMPHASIS = " invalid"
      Case 3
        lpMP3Info.EMPHASIS = " CITT j. 17"
    End Select
    
    'Getting channel info
    tmpNum = (tmpByte \ 16) \ 4
    Select Case tmpNum
      Case 0
        lpMP3Info.Channels = " Stereo"
      Case 1
        lpMP3Info.Channels = " Joint Stereo"
      Case 2
        lpMP3Info.Channels = " 2 Channel"
      Case 3
        lpMP3Info.Channels = " Mono"
    End Select
  End If
  lpMP3Info.SIZE = FileLen(lpMP3File) '& " bytes"
End Sub

Private Sub GetVBRrate(ByVal lpMP3File As String, ByVal byteRead As Long, ByRef lpVBRinfo As VBRinfo)
  Dim I As Long
  Dim ok As Boolean

  I = 0
  byteRead = byteRead - &H39
  Do
    If byteRead > 0 Then
      I = I + 1
      byteRead = byteRead - 38 - Deljivo(I)
    Else
      ok = True
    End If
  Loop Until ok
  lpVBRinfo.VBRlength = Trim(Str(I)) '& " seconds"
  lpVBRinfo.VBRrate = Trim(Str(Int(8 * FileLen(lpMP3File) / (1000 * I)))) & " Kbit (VBR)"
End Sub

Private Function Deljivo(ByVal Num As Long) As Byte
  If Num Mod 3 = 0 Then
    Deljivo = 1
  Else
    Deljivo = 0
  End If
End Function
