VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "電腦術語大考驗-兆炫"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   4950
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "解答"
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "命題"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   960
      Left            =   3600
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "題目："
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "請點選相對應的英文電腦語言："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer


Private Sub Command1_Click()

Dim i As Integer
Dim a(1 To 5) As String
Dim b(1 To 5) As String

a(1) = InputBox("請輸入對應中文電腦術語名稱", "命題")
b(1) = InputBox("請輸入對應的英文電腦術語名稱", "命題")
a(2) = InputBox("請輸入對應中文電腦術語名稱", "命題")
b(2) = InputBox("請輸入對應的英文電腦術語名稱", "命題")
a(3) = InputBox("請輸入對應中文電腦術語名稱", "命題")
b(3) = InputBox("請輸入對應的英文電腦術語名稱", "命題")
a(4) = InputBox("請輸入對應中文電腦術語名稱", "命題")
b(4) = InputBox("請輸入對應的英文電腦術語名稱", "命題")
a(5) = InputBox("請輸入對應中文電腦術語名稱", "命題")
b(5) = InputBox("請輸入對應的英文電腦術語名稱", "命題")

n = Int(Rnd() * 5) + 1
Label2.Caption = Label2.Caption + a(n)

For i = 1 To 5
   List1.AddItem b(i)
Next i

End Sub

Private Sub Command2_Click()

Label2.Caption = "題目："
Label3.Caption = ""
List1.Clear

End Sub

Private Sub Command3_Click()

If List1.ListIndex + 1 = n Then
  Label3.Caption = "答對了！"
Else
  Label3.Caption = "答錯了！"
End If

End Sub

