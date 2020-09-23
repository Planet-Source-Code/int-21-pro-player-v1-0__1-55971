Attribute VB_Name = "ModWav"
'----------------------------------------------------------------------
' Deeth Stereo Oscilloscope v1.0
' A simple oscilloscope application -- now in <<stereo>>
'----------------------------------------------------------------------
' Opens a waveform audio device for 8-bit 11kHz input, and plots the
' waveform to a window.  Can only be resized to a certain minimum
' size defined by the Shape box.
'----------------------------------------------------------------------
' It would be good to make this use the same double-buffering
' scheme as the Spectrum Analyzer.
'----------------------------------------------------------------------
' Murphy McCauley (MurphyMc@Concentric.NET) 08/12/99
'----------------------------------------------------------------------

Option Explicit

Public DevHandle As Long
Public InData(0 To 511) As Byte
Public InData2(0 To 1023) As Integer
Public Inited As Boolean
Public MinHeight As Long, MinWidth As Long

Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    Reserved As Long
End Type

Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Public Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Public Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Public Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Public Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Public Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Public Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Public Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Public Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit

Public Const WAVE_FORMAT_PCM = 1

Public Const WHDR_DONE = &H1&              '/* done bit */
Public Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Public Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Public Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Public Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Public Const WIM_OPEN = &H3BE
Public Const WIM_CLOSE = &H3BF
Public Const WIM_DATA = &H3C0

Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long

Declare Function waveInGetNumDevs Lib "winmm" () As Long
Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Public Const AngleNumerator = 6.283185   ' 2 * Pi = 2 * 3.14159265358979
Public Const NumSamples = 1024
Public Const NumBits = 10
'Used to store pre-calculated values
Private ReversedBits(0 To NumSamples - 1) As Long

Function Start(pChannels As Single, pBits As Single)
    Static WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = pChannels 'Two channels -- left and right
        .SamplesPerSec = 11025 '11khz
        .BitsPerSample = pBits
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    waveInOpen DevHandle, 0, VarPtr(WaveFormat), 0, 0, 0
    
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Function
    End If
    
    Call waveInStart(DevHandle)
    
    Inited = True
    
End Function

Function Finish()
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
End Function

Sub DoReverse()
    'I pre-calculate all these values.  It's a lot faster to just read them from an
    'array than it is to calculate 1024 of them every time FFTAudio() gets called.
    Dim I As Long
    For I = LBound(ReversedBits) To UBound(ReversedBits)
        ReversedBits(I) = ReverseBits(I, NumBits)
    Next
End Sub

Function ReverseBits(ByVal Index As Long, NumBits As Byte) As Long
    Dim I As Byte, Rev As Long
    
    For I = 0 To NumBits - 1
        Rev = (Rev * 2) Or (Index And 1)
        Index = Index \ 2
    Next
    
    ReverseBits = Rev
End Function


Sub FFTAudio(RealIn() As Integer, RealOut() As Single)
    'In this case, NumSamples isn't included (since it's always the same),
    'and the imaginary components are left out since they have no meaning here.
    
    'I've used Singles instead of Doubles pretty much everywhere.  I think this
    'makes it faster, but due to type conversion, it actually might not.  I should
    'check, but I haven't.
    
    'The imaginary components have no meaning in this application.  I just left out
    'the parts of the calculation that need the imaginary input values (which is a
    'big speed improvement right there), but we still need the output array because
    'it's used in the calculation.  It's static so that it doesn't get reallocated.
    Static ImagOut(0 To NumSamples - 1) As Single
    
    'In fact... I declare everything as static!  They all get initialized elsewhere,
    'and Staticing them saves from wasting time reallocating and takes pressure off
    'the heap.
    Static I As Long, j As Long, k As Long, n As Long, BlockSize As Long, BlockEnd As Long
    Static DeltaAngle As Single, DeltaAr As Single
    Static Alpha As Single, Beta As Single
    Static TR As Single, TI As Single, AR As Single, AI As Single
    
    For I = 0 To (NumSamples - 1)
        j = ReversedBits(I) 'I saved time here by pre-calculating all these values
        RealOut(j) = RealIn(I)
        ImagOut(j) = 0 'Since this array is static, gotta make sure it's clear
    Next
    
    BlockEnd = 1
    BlockSize = 2
    
    Do While BlockSize <= NumSamples
        DeltaAngle = AngleNumerator / BlockSize
        Alpha = Sin(0.5 * DeltaAngle)
        Alpha = 2! * Alpha * Alpha
        Beta = Sin(DeltaAngle)
        
        I = 0
        Do While I < NumSamples
            AR = 1!
            AI = 0!
            
            j = I
            For n = 0 To BlockEnd - 1
                k = j + BlockEnd
                TR = AR * RealOut(k) - AI * ImagOut(k)
                TI = AI * RealOut(k) + AR * ImagOut(k)
                RealOut(k) = RealOut(j) - TR
                ImagOut(k) = ImagOut(j) - TI
                RealOut(j) = RealOut(j) + TR
                ImagOut(j) = ImagOut(j) + TI
                DeltaAr = Alpha * AR + Beta * AI
                AI = AI - (Alpha * AI - Beta * AR)
                AR = AR - DeltaAr
                j = j + 1
            Next
            
            I = I + BlockSize
        Loop
        
        BlockEnd = BlockSize
        BlockSize = BlockSize * 2
    Loop
End Sub
