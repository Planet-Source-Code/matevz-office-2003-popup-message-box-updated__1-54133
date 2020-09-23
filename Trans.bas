Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Public Function isTransparent(ByVal hWnd As Long) As Boolean
On Error Resume Next
Dim msg As Long
msg = GetWindowLong(hWnd, GWL_EXSTYLE)
If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
  isTransparent = True
Else
  isTransparent = False
End If
If Err Then
  isTransparent = False
End If
End Function

Public Function MakeTransparent(ByVal hWnd As Long, Perc As Integer) As Long
Dim msg As Long
On Error Resume Next
If Perc < 0 Or Perc > 255 Then
  MakeTransparent = 1
Else
  msg = GetWindowLong(hWnd, GWL_EXSTYLE)
  msg = msg Or WS_EX_LAYERED
  SetWindowLong hWnd, GWL_EXSTYLE, msg
  SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
  MakeTransparent = 0
End If
If Err Then
  MakeTransparent = 2
End If
End Function

Public Function MakeOpaque(ByVal hWnd As Long) As Long
Dim msg As Long
On Error Resume Next
msg = GetWindowLong(hWnd, GWL_EXSTYLE)
msg = msg And Not WS_EX_LAYERED
SetWindowLong hWnd, GWL_EXSTYLE, msg
SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
MakeOpaque = 0
If Err Then
  MakeOpaque = 2
End If
End Function


Public Function PopMessageBox(Caption As String, stretch As Boolean, Transparent As Long)
 With Form1
  .Office2003_PopupMessage1.Caption = Caption
  .Office2003_Button2.Visible = stretch
  .Office2003_PopupMessage1.TransparencyDirection = Transparent
  .Show
 End With
End Function

