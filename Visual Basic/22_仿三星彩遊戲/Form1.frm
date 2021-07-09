VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "仿三星彩遊戲-兆炫"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   630
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開獎"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label6 
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label5 
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "開獎號碼："
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "請猜號碼："
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Randomize

Label4.Caption = Int(Rnd() * 10)
Label5.Caption = Int(Rnd() * 10)
Label6.Caption = Int(Rnd() * 10)

If Val(Text1.Text) = Label4.Caption And Val(Text2.Text) = Label5.Caption And Val(Text3.Text) = Label6.Caption Then
   Label3.ForeColor = &H0&
   Label3.Caption = "恭喜你猜中了!"
Else
   Label3.ForeColor = &HFF&
   Label3.Caption = "恭喜你猜錯了！"
End If

End Sub

