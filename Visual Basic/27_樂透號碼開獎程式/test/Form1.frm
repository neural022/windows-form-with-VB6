VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   14970
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   8280
      TabIndex        =   0
      Top             =   6840
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim z(1 To 42) As Integer
Private Sub Command1_Click()


For i = 1 To 42

Do

z(i) = Int(Rnd() * 42) + 1
k = 0
    For j = 1 To i - 1
    
        If z(j) = z(i) Then
        k = 1
        End If
    Next j
    
Loop Until k = 0

Next i




a = 1
While a < 42

For i = 1 To (42 - a)

If z(i) > z(i + 1) Then
 c = z(i)
 z(i) = z(i + 1)
 z(i + 1) = c
End If

Next i
a = a + 1
Wend

For x = 1 To 42

Print Chr(10) & z(x);

Next x

End Sub

