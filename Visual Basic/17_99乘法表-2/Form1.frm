VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "99乘法表橫向-兆炫"
   ClientHeight    =   2355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   11025
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "列印"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Integer
Private Sub Command1_Click()

For a = 1 To 9

   For b = 1 To 9
   
       Print "    ";
       
       Print a & "x" & b & "=" & a * b,
       
   Next b
  Print
  
Next a

End Sub
