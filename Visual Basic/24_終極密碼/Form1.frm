VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "兆炫的終極密碼"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "開始遊戲"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "設定數字"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "遊戲說明：先設定數字，再換對手「開始遊戲」猜密碼。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1), b As Integer
Private Sub Command1_Click()

a(1) = InputBox("請輸入一個1~99的數", "猜數字遊戲-設定數字")

End Sub

Private Sub Command2_Click()

Dim c, d As Integer

c = 1
d = 99
b = InputBox("請輸入" & c & "~" & d & "的數", "終極密碼-猜數字中..")

If b = a(1) Then
   MsgBox "恭喜你猜對了", vbInformation, "終極密碼猜數字"
 Else
 
    Do
     If b > a(1) Then
       d = b
       b = InputBox("請輸入" & c & "~" & d & "的數", "終極密碼-猜數字中..")
      Else
        c = b
         b = InputBox("請輸入" & c & "~" & d & "的數", "終極密碼-猜數字中..")
      End If
    Loop Until b = a(1)
      MsgBox "恭喜你猜對了", vbInformation, "終極密碼猜數字"
    
End If

End Sub
