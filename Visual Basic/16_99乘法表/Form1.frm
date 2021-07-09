VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "99乘法表直向-兆炫"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   3675
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "列印"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   5640
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Integer
Private Sub Command1_Click()

For a = 1 To 7 Step 3

   For b = 1 To 9
   
       Print "    ";
       
       Print a & "x" & b & "=" & a * b,
       Print a + 1 & "x" & b + 1 & "=" & (a + 1) * b,
       Print a + 2 & "x" & a + 2 & "=" & (a + 2) * b
       
       
   Next b
  Print
  
Next a

End Sub

