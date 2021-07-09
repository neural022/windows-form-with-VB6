VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "繪製由*組成的圖案-兆炫"
   ClientHeight    =   2145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   6180
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame1 
      Caption         =   "點選想繪製的圖形"
      Height          =   1575
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "關閉"
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "平行四邊形"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "正三角形"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "矩形"
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, d, e, f As Integer
Private Sub Command1_Click()
End
End Sub

Private Sub Option1_Click()

Cls

Print
For a = 1 To 5

   Print Space(2);
   Print String(7, "*")
 
Next a

Print
 
End Sub

Private Sub Option2_Click()

Cls

Print

c = 1
While c < 11

      d = 12 - c
      Print Space(d);
      Print String(c, "*")
      c = c + 2
Wend

Print

End Sub

Private Sub Option3_Click()

Cls

Print
  
e = 0
  Do
   
   f = e + 2
   Print Space(f);
   Print String(7, "*")
   e = e + 1
   
  Loop While e < 5
  
Print

End Sub
