VERSION 5.00
Begin VB.UserControl xVis 
   BackColor       =   &H00000000&
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   ForeColor       =   &H00008000&
   ScaleHeight     =   54
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   208
   Begin VB.PictureBox Scope 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00008000&
      Height          =   810
      Left            =   0
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   208
      TabIndex        =   1
      Top             =   0
      Width           =   3120
   End
   Begin VB.PictureBox ScopeBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   1770
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "xVis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lpPrevWndProc&
Const GWL_WNDPROC = -4
Dim Working As Boolean
Dim VolumeL&, VolumeR&
Const Max_Volume = 128
Dim mStyle As Single
Dim mVis As Boolean
Dim ScopeHeight&, Divisor&
Private OutData(0 To NumSamples - 1) As Single
Enum eStyle
    Sine = 1
    Spectrum
    Spectrum2
End Enum
Dim TamaBuffer&
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Property Get Style() As eStyle
    Style = mStyle
End Property

Public Property Let Style(ByVal newStyle As eStyle)
    mStyle = newStyle
End Property

Private Sub Scope_DblClick()
If Working Then
    mStyle = mStyle + 1
    If mStyle > 3 Then mStyle = 1
    Finish
    Init
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Style", mStyle, 1)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mStyle = PropBag.ReadProperty("Style", 1)
End Sub

Private Sub UserControl_Terminate()
    Finish
End Sub
                                
                                '********** Functions*********
                                '******************************
Private Sub DrawSpectrum2()
Dim I&
    With ScopeBuff 'Save some time referencing it...

            Call FFTAudio(InData2, OutData)
            .CurrentX = 0
            .CurrentY = ScopeHeight
            .Cls
            For I = 0 To 255 Step 4
                .CurrentY = ScopeHeight
                .CurrentX = I
                
                'I average two elements here because it gives a smoother appearance.
                ScopeBuff.Line (I, ScopeHeight)-(I + 2, ScopeHeight - (Sqr(Abs(OutData(I * 2) \ Divisor)) + Sqr(Abs(OutData(I * 2 + 1) \ Divisor)))), , BF
            Next
            
            Scope.Picture = .Image 'Display the double-buffer
            DoEvents
            
    End With
    
End Sub
                                
Private Sub DrawSpectrum()
Dim I&
    With ScopeBuff 'Save some time referencing it...

            Call FFTAudio(InData2, OutData)
            .CurrentX = 0
            .CurrentY = ScopeHeight
            .Cls
            For I = 0 To 255
                .CurrentY = ScopeHeight
                .CurrentX = I
                
                'I average two elements here because it gives a smoother appearance.
                ScopeBuff.Line Step(0, 0)-(I, ScopeHeight - (Sqr(Abs(OutData(I * 2) \ Divisor)) + Sqr(Abs(OutData(I * 2 + 1) \ Divisor))))
            Next
            
            Scope.Picture = .Image 'Display the double-buffer
            
            DoEvents
            
    End With
    
End Sub
                         

Private Sub DrawSine()
    Static X As Long
    Scope.Cls
    Scope.CurrentX = -1
    Scope.CurrentY = Scope.ScaleHeight \ 2
    
    'Plot the data...
    For X = 0 To 255
        Scope.Line Step(0, 0)-(X, InData(X * 2)), vbRed
        'Use these to plot dots instead of lines...
        'UserControl.PSet (X, InData(X * 2))
    Next
    Scope.CurrentX = -1
    'UserControl.CurrentY = (UserControl.ScaleHeight \ 2) \ 2
    For X = 0 To 255
        Scope.Line Step(0, 0)-(X, InData(X * 2 + 1)), vbYellow  'For a good soundcard...
        'UserControl.PSet (X, InData(X * 2 + 1)) 'For a good soundcard...
    Next
    'UserControl.CurrentY = UserControl.Width
    'Scope(1).CurrentY = Scope(0).Width
End Sub

Function Visualizador()
     Static Wave As WaveHdr

    
    Do
    
'**********************************************************
    If (mStyle = 2) Or (mStyle = 3) Then
        Wave.lpData = VarPtr(InData2(0))
        Wave.dwBufferLength = 1024 'This is now 512 so there's still 256 samples per channel
    Else
        Wave.lpData = VarPtr(InData(0))
        Wave.dwBufferLength = 512 'This is now 512 so there's still 256 samples per channel
    End If
    
    Wave.dwFlags = 0
'**********************************************************
    
    
        Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
    
        Do
            'Nothing -- we're waiting for the audio driver to mark
            'this wave chunk as done.
        Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
        
        Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        
        If DevHandle = 0 Then
            'The device has closed...
            Exit Do
        End If
        
        'Call DrawData

        'Scope.Cls
    

        Select Case mStyle
            Case 1
                DrawSine
                'DrawLines
            Case 2
                DrawSpectrum2
            Case 3
                DrawSpectrum
        End Select
    
        DoEvents
    Loop While (DevHandle <> 0) And (Working = True)  'While the audio device is open

End Function

Sub DrawLines()
Dim I&, lMitad&
    '**************************************************************
        'Draw vertical lines at left
        lMitad = Scope.ScaleHeight / 2
        'Line
        Scope.Line (15, 0)-(15, Scope.ScaleHeight), vbWhite
        For I = 1 To 255 Step 20
            Scope.Line (15, I)-(5, I), vbWhite
        Next
        
        Scope.Line (15, lMitad - 25)-(3, lMitad - 25), vbWhite
        Scope.Line (15, lMitad - 7)-(0, lMitad - 7), vbWhite
        Scope.Line (15, lMitad + 9)-(3, lMitad + 9), vbWhite
        'UserControl.Line (0, lMitad + 32)-(15, lMitad + 32), vbWhite
        '**************************************************************
        
                '**************************************************************
        'Draw Horizontal lines at bottom
        lMitad = Scope.ScaleWidth / 2
        'Line
        Scope.Line (15, 230)-(lMitad * 2, 230), vbWhite
        For I = 15 To 255 Step 20
            Scope.Line (I, 230)-(I, 245), vbWhite
        Next
        
        Scope.Line (lMitad - 32, 255)-(lMitad - 32, 225), vbWhite
        Scope.Line (lMitad + 27, 255)-(lMitad + 27, 225), vbWhite
        '**************************************************************
End Sub

Sub Init()
    ScopeBuff.Width = Scope.ScaleWidth
    ScopeBuff.Height = Scope.ScaleHeight
    ScopeBuff.BackColor = Scope.BackColor
    ScopeBuff.ForeColor = Scope.ForeColor
    Scope.Picture = LoadPicture("")
    If (mStyle = 2) Or (mStyle = 3) Then
        Start 1, 16

    Else
        Start 2, 8
        Scope.ScaleWidth = 255
        Scope.ScaleHeight = 255
    End If
    DoReverse
    Divisor = ((10 - 6 + 1) / 10) * 5200
    
    ScopeHeight = Scope.Height
'''''    lpPrevWndProc = SetWindowLong(UserControl.hwnd, GWL_WNDPROC, AddressOf Finish)
    
End Sub
Sub Finished()
    Finish
End Sub
                                '******************************
                                '******************************

Public Property Get Vis() As Boolean
    Vis = Working
End Property

Public Property Let Vis(ByVal NewVis As Boolean)
    Working = NewVis
    If Working Then Visualizador
End Property
