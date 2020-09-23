VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin Project1.Office2003_PopupMessage Office2003_PopupMessage1 
      Height          =   1110
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _extentx        =   8705
      _extenty        =   1958
      picture         =   "Form1.frx":0000
      caption         =   "Some text"
      Begin Project1.Office2003_Button Office2003_Button2 
         Height          =   270
         Left            =   4370
         TabIndex        =   2
         Top             =   140
         Width           =   270
         _extentx        =   476
         _extenty        =   476
         closebutton     =   0   'False
         align           =   0
         caption         =   ""
         font            =   "Form1.frx":0CDA
         backcolor       =   -2147483633
         picture         =   "Form1.frx":0D06
      End
      Begin Project1.Office2003_Button Office2003_Button1 
         Height          =   270
         Left            =   4650
         TabIndex        =   1
         Top             =   135
         Width           =   270
         _extentx        =   476
         _extenty        =   476
         closebutton     =   -1  'True
         font            =   "Form1.frx":1D18
         picture         =   "Form1.frx":1D44
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal cKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long

Private m_objParent As Object
 Dim MoveScreen As Boolean
 Dim CurrX As Integer
 Dim CurrY As Integer
 Dim MousX As Integer
 Dim MousY As Integer





Private Sub Form_Load()
Me.Height = Me.Office2003_PopupMessage1.Height
Me.Width = Me.Office2003_PopupMessage1.Width
Me.Top = GetSetting("Office2003", "Position", "Top", 1000)
Me.Left = GetSetting("Office2003", "Position", "Left", 1000)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Office2003_PopupMessage1.TransparencyLevel > 0 Then
    Cancel = True
    Office2003_PopupMessage1.TransparencyDirection = -2
  End If
End Sub

Private Sub Office2003_Button1_Click()
 If Office2003_PopupMessage1.TransparencyLevel > 0 Then
    Office2003_PopupMessage1.TransparencyDirection = -5
  End If
End Sub

Private Sub Office2003_Button2_Click()
If Me.Height = 1110 Then
 Me.Office2003_PopupMessage1.Povecaj 1700
Me.Height = 1700
Else
 Me.Office2003_PopupMessage1.Povecaj 1110
Me.Height = 1110
End If
End Sub

Private Sub Office2003_PopupMessage1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MoveScreen = True
  MousX = X
  MousY = Y
End Sub
Private Sub Office2003_PopupMessage1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MoveScreen Then
  Call SaveSetting("Office2003", "Position", "Top", Me.Top)
  Call SaveSetting("Office2003", "Position", "Left", Me.Left)
  CurrX = Form1.Left - MousX + X
  CurrY = Form1.Top - MousY + Y
   Form1.Move CurrX, CurrY
 End If
End Sub
Private Sub Office2003_PopupMessage1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MoveScreen = False
End Sub

