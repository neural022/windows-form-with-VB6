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
   Begin VB.CommandButton Command3 
      Caption         =   "重新開始"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3840
      Top             =   240
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
   Begin VB.Label Label3 
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "按Go鈕會顯示五個數字，並在3秒後消失"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b(5), c As Integer

Private Sub Command1_Click()

Randomize

For a = 1 To 5
  b(a) = Int(Rnd() * 99) + 1
  Text1.Text = Text1.Text & b(a) & Space(4)
Next a

Timer1.Enabled = True


End Sub
Private Sub Command2_Click()

b1 = Val(Text2.Text)
b2 = Val(Text3.Text)
b3 = Val(Text4.Text)
b4 = Val(Text5.Text)
b5 = Val(Text6.Text)

c = 0

For a = 1 To 5
    
    If b1 = b(a) Then
    c = c + 1
    End If
    
    If b2 = b(a) Then
    c = c + 1
    End If
    
    If b3 = b(a) Then
    c = c + 1
    End If
    
    If b4 = b(a) Then
    c = c + 1
    End If
    
    If b5 = b(a) Then
    c = c + 1
    End If
    
Next a

Text1.Visible = True

Label3.Caption = "恭喜你共答對" & c & "題"

If c = 0 Then
Label3.Caption = "恭喜你全錯!"
End If

End Sub


Private Sub Command3_Click()

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Label3.Caption = ""
Text1.Visible = True
Command1.Enabled = True
Command2.Enabled = False

End Sub

Private Sub Timer1_Timer()

Text1.Visible = False
Command1.Enabled = False
Command2.Enabled = True
Timer1.Enabled = False

End Sub
