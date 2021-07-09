VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "計算面積-陳兆炫"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   6855
   StartUpPosition =   3  '系統預設值
   Begin VB.OptionButton Option6 
      Caption         =   "圓形"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.OptionButton Option5 
      Caption         =   "平行四邊形"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "梯形"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "三角形"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "正方形"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "長方形"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計算面積"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
a1 = InputBox("請輸入長(cm)", "長方形面積")
a2 = InputBox("請輸入寬(cm)", "長方形面積")
a = a1 * a2
MsgBox "長方形面積" & a & "平方公分"
End If

If Option2.Value = True Then
b1 = InputBox("請輸入邊長(cm)", "正方形面積")
b = b1 ^ 2
MsgBox "正方形面積" & b & "平方公分"
End If

If Option3.Value = True Then
c1 = InputBox("請輸入底(cm)", "三角形面積")
c2 = InputBox("請輸入高(cm)", "三角形面積")
c = c1 * c2 / 2
MsgBox "三角形面積" & c & "平方公分"
End If

If Option4.Value = True Then
d1 = InputBox("請輸入上底(cm)", "梯形面積")
d2 = InputBox("請輸入下底(cm)", "梯形面積")
d3 = InputBox("請輸入高(cm)", "梯形面積")
d = (Val(d1) + Val(d2)) * d3 / 2
MsgBox "梯形面積" & d & "平方公分"
End If

If Option5.Value = True Then
e1 = InputBox("請輸入底(cm)", "平行四邊形面積")
e2 = InputBox("請輸入高(cm)", "平行四邊形面積")
e = e1 * e2
MsgBox "平行四邊形面積", "平方公分"
End If

If Option6.Value = True Then
f1 = InputBox("請輸入半徑(cm)", "圓形面積")
f = f1 ^ 2 * 3.14
MsgBox "圓形面積" & f & "平方公分"
End If
 
End Sub

Private Sub Command2_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
End Sub

Private Sub Command3_Click()
End
End Sub
