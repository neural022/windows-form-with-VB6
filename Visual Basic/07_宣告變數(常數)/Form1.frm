VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "兆炫-宣告常數"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4980
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "重新輸入"
      Height          =   1695
      Left            =   3600
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "圓周長"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "圓面積"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "公分"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "圓半徑"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const pi As Single = 3.14159

Private Sub Command1_Click()
Text1.Text = ""
Label2.Caption = ""
Label3.Caption = ""
End Sub

Private Sub Text1_Change()
Label2.Caption = Val(pi) * Val(Text1.Text) ^ 2 & "平方公分"
Label3.Caption = 2 * Val(pi) * Val(Text1.Text) & "公分"
End Sub
