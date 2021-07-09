VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "記憶大考驗-兆炫"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3840
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3960
      Top             =   120
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "看答案"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "　　按Go鈕會顯示五個數字，且在3秒後消失，並於公佈答案約10秒後重新開始。"
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "遊戲說明："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim a, b As Integer
Randomize

For a = 1 To 5
  b = Int(Rnd() * 99) + 1
  Text1.Text = Text1.Text & b & Space(4)
Next a

Timer1.Enabled = True

End Sub

Private Sub Command2_Click()


Timer2.Enabled = True

Text2.Text = Text1.Text

End Sub



Private Sub Timer1_Timer()
Text1.Visible = False
Command1.Enabled = False
Command2.Enabled = True
Timer1.Enabled = False

End Sub


Private Sub Timer2_Timer()

Text1.Visible = True
Text1.Text = ""
Text2.Text = ""
Command1.Enabled = True
Command2.Enabled = False
Timer2.Enabled = False

End Sub

