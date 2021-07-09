VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "陳兆炫-字數統計"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "字數統計"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "清除重來"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入文字"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
Label2.Caption = ""
End Sub

Private Sub Command2_Click()
Dim abc As Integer
abc = Len(Text1.Text)
Label2.Caption = "共輸入" & abc & "字"
End Sub

