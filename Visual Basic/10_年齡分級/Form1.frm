VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "年齡分級-兆炫"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   3000
      TabIndex        =   9
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1560
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重新輸入"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確認"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "請輸入出生年月日"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "日"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "月"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label2 
      Height          =   975
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "年"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Integer, b As Integer, c As Integer, d As Integer
a = Text1.Text
b = Text2.Text
c = Text3.Text
d = Year(Date) - a


If d <= 13 Then
   Label2.Caption = "您的年齡為" & d & "屬於童年"
ElseIf d < 20 Then
   Label2.Caption = "您的年齡為" & d & "屬於少年"
ElseIf d < 30 Then
   Label2.Caption = "您的年齡為" & d & "屬於青年"
ElseIf d < 55 Then
   Label2.Caption = "您的年齡為" & d & "屬於壯年"
Else
   Label2.Caption = "您的年齡為" & d & "屬於老年"
End If

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Label2.Caption = ""

End Sub

Private Sub Command3_Click()
End
End Sub

