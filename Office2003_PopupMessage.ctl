VERSION 5.00
Begin VB.UserControl Office2003_PopupMessage 
   Appearance      =   0  'Flat
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3165
   ScaleWidth      =   5970
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   4740
      TabIndex        =   0
      Top             =   240
      Width           =   4740
      Begin VB.Image Pic_Image 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Lbl_Caption 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3855
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4920
      Top             =   480
   End
   Begin VB.Image Top_Silver 
      Height          =   120
      Left            =   0
      Picture         =   "Office2003_PopupMessage.ctx":0000
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Image Top_Olive 
      Height          =   120
      Left            =   0
      Picture         =   "Office2003_PopupMessage.ctx":1F22
      Top             =   2280
      Width           =   4935
   End
   Begin VB.Image Top_Blue 
      Height          =   120
      Left            =   0
      Picture         =   "Office2003_PopupMessage.ctx":3E44
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Image Pic_top 
      Height          =   120
      Left            =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "Office2003_PopupMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim xx, xxx, R1, R2, G1, G2, B1, B2, Rs, Gs, Bs, Rx, Gx, Bx
Dim LCol1, Border1, Border2

Public Enum AppearanceConst
    Blue = 0
    Silver = 1
    Olive = 2
End Enum
Private MyCaption As String
Private MyFont As Font
Private MyForeColor As OLE_COLOR
Private DefForeColor As OLE_COLOR
Private NewButtonIcon As Picture
Private MyAppearance As AppearanceConst
Private Const MyDefAppearance = Blue
Private Const DefCaption = "KDC"


Const m_def_TransparencyLevel = 0
Const m_def_TransparencyDirection = 0
Const m_def_Text = "Text"

Dim m_TransparencyLevel As Integer
Dim m_TransparencyDirection As Integer
Dim m_Caption As String


Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Function MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseOut(Button, Shift, X, Y)
End Function

Private Sub Lbl_Caption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub Lbl_Caption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub Lbl_Caption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub Pic_Middle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub Pic_Middle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub Pic_Middle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub Pic_top_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub Pic_top_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub Pic_top_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Appearance = PropBag.ReadProperty("Appearance", MyDefAppearance)
Set UserControl.Pic_Image.Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.Lbl_Caption.Caption = PropBag.ReadProperty("Caption", "Some text")
    m_TransparencyLevel = PropBag.ReadProperty("TransparencyLevel", m_def_TransparencyLevel)
    m_TransparencyDirection = PropBag.ReadProperty("TransparencyDirection", m_def_TransparencyDirection)
    MakeTransparent UserControl.Parent.hWnd, m_TransparencyLevel
    If Ambient.UserMode Then Timer1.Enabled = True
End Sub

Private Sub UserControl_Resize()
If UserControl.Width <> 0 Then
        Pic1.Width = UserControl.Width
        Pic1.Height = UserControl.Height - Pic_top.Height
        Pic_top.Top = 0
        Pic_top.Left = 0
        Pic_Image.Top = 75
        Pic_Image.Left = 100
        Lbl_Caption.Top = 220
        Lbl_Caption.Left = 720
        UserControl.Width = Pic_top.Width
        Call SetGradient
    End If
End Sub
Public Property Get Caption() As String
    Caption = UserControl.Lbl_Caption.Caption
End Property


Public Property Let Caption(ByVal newCaption As String)

    UserControl.Lbl_Caption.Caption = newCaption
    UserControl.Refresh
    PropertyChanged "Caption"
End Property
Public Property Get Picture() As Picture
Set Picture = Pic_Image.Picture

End Property

Public Property Set Picture(ByVal picNew As Picture)
Set UserControl.Pic_Image.Picture = picNew
PropertyChanged "Picture"
End Property


Private Sub UserControl_Terminate()
    DoEvents
End Sub


Private Sub UserControl_Initialize()

    
    Pic1.Left = 0
    Pic1.Top = Pic_top.Height
    UserControl.Height = Pic1.Height
    UserControl.Width = Pic1.Width
    Call UserControl_Resize
    
    
    m_TransparencyLevel = 0
    m_TransparencyDirection = 0
   'UserControl.Height = Pic_top.Height + Pic_Middle.Height + Pic_down.Height
'UserControl.Width = Pic_top.Width
End Sub

Private Sub UserControl_InitProperties()
    Appearance = Blue
    m_TransparencyLevel = m_def_TransparencyLevel
    m_TransparencyDirection = m_def_TransparencyDirection
    If Ambient.UserMode Then Timer1.Enabled = True
End Sub

Public Property Get Appearance() As AppearanceConst
    Appearance = MyAppearance
End Property
Public Property Let Appearance(ByVal vData As AppearanceConst)
    MyAppearance = vData
    Call SetGradient
    ForeColor = DefForeColor
PropertyChanged "ForeColor"
PropertyChanged "Appearance"
End Property


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Appearance", MyAppearance, MyDefAppearance)
 PropBag.WriteProperty "Picture", UserControl.Pic_Image.Picture, Nothing
   PropBag.WriteProperty "Caption", UserControl.Lbl_Caption.Caption
    Call PropBag.WriteProperty("TransparencyLevel", m_TransparencyLevel, m_def_TransparencyLevel)
    Call PropBag.WriteProperty("TransparencyDirection", m_TransparencyDirection, m_def_TransparencyDirection)
End Sub

Private Sub Timer1_Timer()

  If m_TransparencyDirection <> 0 Then
    If MakeTransparent(UserControl.Parent.hWnd, m_TransparencyLevel) = 1 Then
      If m_TransparencyDirection < 0 Then UserControl.Parent.Visible = False
    End If
    m_TransparencyLevel = m_TransparencyLevel + m_TransparencyDirection
    If m_TransparencyLevel < Abs(m_TransparencyDirection) Then
      m_TransparencyDirection = 0
      m_TransparencyLevel = 0
      MakeTransparent UserControl.Parent.hWnd, m_TransparencyLevel
      Unload UserControl.Parent
    End If
    If m_TransparencyLevel > (255 - Abs(m_TransparencyDirection)) Then
      m_TransparencyDirection = 0
      m_TransparencyLevel = 255
    End If
  End If
  
End Sub






Public Property Get TransparencyDirection() As Long
    TransparencyDirection = m_TransparencyDirection
End Property

Public Property Let TransparencyDirection(ByVal New_TransparencyDirection As Long)
    m_TransparencyDirection = New_TransparencyDirection
    PropertyChanged "TransparencyDirection"
End Property


Public Property Get TransparencyLevel() As Long
    TransparencyLevel = m_TransparencyLevel
End Property

Public Property Let TransparencyLevel(ByVal New_TransparencyLevel As Long)
    m_TransparencyLevel = New_TransparencyLevel
    PropertyChanged "TransparencyLevel"
End Property

Public Function MakeVisible() As Variant
    m_TransparencyLevel = 0
    m_TransparencyDirection = 4
    MakeTransparent UserControl.Parent.hWnd, m_TransparencyLevel
    UserControl.Parent.Visible = True
    UserControl.Parent.SetFocus
End Function

Public Function MakeInVisible() As Variant
    m_TransparencyLevel = 255
    m_TransparencyDirection = -4
    MakeTransparent UserControl.Parent.hWnd, m_TransparencyLevel
End Function

Function Povecaj(Height1 As Single)
UserControl.Lbl_Caption.Height = Height1 - 495
UserControl.Pic1.Height = Height1
UserControl.Height = UserControl.Pic_top.Height + Pic1.Height - 120
End Function

Private Sub SetGradient()
    Select Case MyAppearance
        Case Is = Silver
            R1 = &HE8: R2 = &HB4
            G1 = &HEA: G2 = &HB3
            B1 = &HF2: B2 = &HCD
            Pic_top.Picture = Top_Silver.Picture
            Border1 = RGB(75, 75, 111)
            Border2 = RGB(75, 75, 111)
        Case Is = Olive
            R1 = &HE8: R2 = &HC0
            G1 = &HEE: G2 = &HCE
            B1 = &HCD: B2 = &H9A
            Pic_top.Picture = Top_Olive.Picture
            Border1 = RGB(63, 93, 56)
            Border2 = RGB(63, 93, 56)
        Case Is = Blue
            R1 = &HD6: R2 = &HA8
            G1 = &HE7: G2 = &HC4
            B1 = &HFC: B2 = &HEE
            Pic_top.Picture = Top_Blue.Picture
            Border1 = RGB(0, 0, 128)
            Border2 = RGB(0, 0, 128)
    End Select

Rx = R1: Gx = G1: Bx = B1
Rs = (R1 - R2) / (Pic1.ScaleHeight - 1)
Gs = (G1 - G2) / (Pic1.ScaleHeight - 1)
Bs = (B1 - B2) / (Pic1.ScaleHeight - 1)
    For xx = 0 To Pic1.Height - 1
      Pic1.Line (0, xx)-(Pic1.Width, xx), RGB(Rx, Gx, Bx)
        Rx = Rx - Rs
        Gx = Gx - Gs
        Bx = Bx - Bs
    Next xx

    
Pic1.Line (0, 0)-(Pic1.Width - 1, Pic1.Height - 1), Border1, B
Pic1.Line (0, Pic1.Height - 10)-(Pic1.Width, Pic1.Height - 10), Border2
Pic1.Line (Pic1.Width - 10, 0)-(Pic1.Width - 10, Pic1.Height - 10), Border2

'Bord1 = Pic1.Point(0, 0)
'Bord2 = Pic1.Point(Pic1.Width - 10, Pic1.Height - 10)
End Sub



