VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4860
   LinkTopic       =   "Form2"
   ScaleHeight     =   3675
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Text            =   "True"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   360
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Text            =   "5"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Popup"
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Stretch button:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Caption:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Transparent Level:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
PopMessageBox Text1.Text, Combo2.Text, Combo1.Text
End Sub

Private Sub Form_Load()
Me.Text1.Text = "This is a sample of the Office 2003 MessageBox usage "
Me.Text1.Text = Me.Text1.Text & "It 's not finish yet it neet alot of programing. "
Me.Text1.Text = Me.Text1.Text & "If someone modify this, please let me know ...."
Me.Text1.Text = Me.Text1.Text & "                All your questions you can send me to mail: Matevz.cel@volja.net."
Me.Text1.Text = Me.Text1.Text & " Please vote for me ;)"

Dim i As Integer
For i = 1 To 50
Combo1.AddItem i
Next i

Combo2.AddItem False
Combo2.AddItem True
End Sub

