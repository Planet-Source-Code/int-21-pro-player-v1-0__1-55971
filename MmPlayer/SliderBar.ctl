VERSION 5.00
Begin VB.UserControl SliderBar 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   334
   ToolboxBitmap   =   "SliderBar.ctx":0000
   Begin VB.PictureBox Cue 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1530
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   0
      Top             =   -4
      Width           =   225
   End
End
Attribute VB_Name = "SliderBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_Min&, m_Max&, m_Pos&
Dim Arrastrando As Boolean
Public Event CambioPosicion(NuevaPosicion As Long)

Private Sub Cue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Arrastrando = True
End Sub

'****************************************************************
'AREA DE OBJETOS*********************************************
'****************************************************************
Private Sub Cue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If (Button = vbLeftButton) And Arrastrando Then 'Si se esta arrastrando con el boton izquierdo del Mouse
        Dim tPos&
        X = X / Screen.TwipsPerPixelX
        If 1 < (X + Cue.Left) Then 'Estamos bien del lado Izq
            
            If (Cue.Left < (UserControl.ScaleWidth - Cue.Width)) Or (X < 0) Then  ' Nos movemos entra Izq. y Der.
                tPos = X + Cue.Left
            Else 'estamos alejados del lado Der. Adjustamos
                tPos = (UserControl.ScaleWidth - Cue.Width)
            End If
            
        Else 'Nos estamos pasando del lado Izq..Adjustamos
            tPos = 0
        End If
        
        Cue.Move tPos
        m_Pos = ((tPos * m_Max) \ UserControl.ScaleWidth)
        
        
    End If
End Sub

Private Sub Cue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Arrastrando And (Button = vbLeftButton) Then
        RaiseEvent CambioPosicion(m_Pos)
        Arrastrando = False
    End If
End Sub

Private Sub UserControl_Initialize()
    m_Min = 1
    m_Max = 100
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Max = PropBag.ReadProperty("Maximo", 100)
    m_Min = PropBag.ReadProperty("Minimo", 1)
    m_Pos = PropBag.ReadProperty("Posicion", 50)
    UserControl.BackColor = PropBag.ReadProperty("BckColor", 0)
    Cue.BackColor = PropBag.ReadProperty("CueColor", vbGrayed)
    
    MoverPos m_Pos
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.Picture = PropBag.ReadProperty("ImgFondo", UserControl.Picture)
    Cue.Picture = PropBag.ReadProperty("ImgCue", Cue.Picture)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Maximo", m_Max, 100)
    Call PropBag.WriteProperty("Minimo", m_Min, 1)
    Call PropBag.WriteProperty("Posicion", m_Pos, 50)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ImgFondo", UserControl.Picture)
    Call PropBag.WriteProperty("ImgCue", Cue.Picture)
    Call PropBag.WriteProperty("CueColor", Cue.BackColor, vbGrayed)
    Call PropBag.WriteProperty("BckColor", UserControl.BackColor, 0)
End Sub
'****************************************************************
'AREA DE PROPIEDADES*********************************************
'****************************************************************
Public Property Get ImgFondo() As Picture
    Set ImgFondo = UserControl.Picture
End Property

Public Property Set ImgFondo(ByVal Img As Picture)
    Set UserControl.Picture = Img
    PropertyChanged "ImgFondo"
End Property

Public Property Get ImgCue() As Picture
    Set ImgCue = Cue.Picture
End Property


Public Property Set ImgCue(ByVal Img As Picture)
    Set Cue.Picture = Img
    PropertyChanged "ImgCue"
End Property

Public Property Get Maximo() As Long
    Maximo = m_Max
End Property

Public Property Let Maximo(ByVal Nuevo As Long)
    If (Nuevo > m_Min) And (Nuevo > 1) Then m_Max = Nuevo
    PropertyChanged "Maximo"
End Property

Public Property Get Minimo() As Long
    Minimo = m_Min
End Property

Public Property Let Minimo(ByVal Nuevo As Long)
    If (Nuevo > 1) And (Nuevo < m_Max) Then m_Min = Nuevo
    PropertyChanged "Minimo"
End Property

Public Property Get Posicion() As Long
    Posicion = m_Pos
End Property

Public Property Let Posicion(ByVal Nuevo As Long)
    If (Nuevo >= m_Min) And (Nuevo <= m_Max) Then
        If Not Arrastrando Then
            m_Pos = Nuevo
            MoverPos m_Pos
            PropertyChanged "Posicion"
        End If
    End If
    
End Property
'****************************************************************
'AREA DE SUB Y FUNCIONES*********************************************
'****************************************************************
Private Sub MoverPos(Valor As Long)
Dim tPos
    If Not Arrastrando Then
        tPos = (Valor * (UserControl.ScaleWidth - Cue.Width)) \ m_Max
        Cue.Move tPos
    End If
End Sub


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


Public Property Get BckColor() As OLE_COLOR
    BckColor = UserControl.BackColor
End Property

Public Property Let BckColor(ByVal vNewValue As OLE_COLOR)
    UserControl.BackColor = vNewValue
End Property

Public Property Get CueColor() As OLE_COLOR
    CueColor = Cue.BackColor
End Property

Public Property Let CueColor(ByVal vNewValue As OLE_COLOR)
    Cue.BackColor = vNewValue
End Property
