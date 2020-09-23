Attribute VB_Name = "Library"
Option Explicit
Public Enum eDirection
    eBack = 1
    eNext
End Enum

Public Enum eActions
    ePlay = 1
    ePaused
    eStop
End Enum

Public CurSong&
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Function AscChar(ByVal MyMot As String) As String
Dim j As Integer
Dim k
Dim TmpMot As String

    For j = 1 To Len(MyMot)
        k = Mid(MyMot, j, 1)
        If Asc(k) >= 20 Then
            TmpMot = TmpMot & k
        End If
    Next j
    If TmpMot = VBA.Space(Len(TmpMot)) Then TmpMot = ""
    AscChar = TmpMot
End Function

Function ConvertSeconds(Seconds) As String    'As Date
Dim Tm   As String
    Tm = Format(Int(Seconds / 60), "00") & ":" & Format(Seconds Mod 60, "00")
    ConvertSeconds = Tm
End Function

Function AddSlash(Path)
Dim sTmp$
If Mid(Path, Len(Path), 1) = "\" Then sTmp = Path Else sTmp = Path & "\"
    AddSlash = sTmp
End Function

Public Sub BloquearControl(XhWnd As Long, cLock As Boolean)
    Dim I As Long
    
    If cLock Then
        ' This will lock the control
        LockWindowUpdate XhWnd
    Else
        ' This will unlock controls
        LockWindowUpdate 0
    End If
End Sub
