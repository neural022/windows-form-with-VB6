VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "樂透號碼開獎程式-兆炫"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "號碼排序"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開獎"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 42) As Integer
Private Sub Command1_Click()

Randomize

For i = 1 To 42

    Do

        a(i) = Int(Rnd() * 42) + 1
        k = 0
             For j = 1 To i - 1
    
               If a(j) = a(i) Then
                 k = 1
               End If
             Next j
    
    Loop Until k = 0

Next i

Label1.Caption = a(1) & Space(4) & a(2) & Space(4) & a(3) & Space(4) & a(4) & Space(4) & a(5) & Space(4) & a(6) & Space(4)

End Sub

Private Sub Command2_Click()

i = 1
While i < 6

    For j = 1 To (6 - i)

        If a(j) > a(j + 1) Then
         b = a(j)
         a(j) = a(j + 1)
        a(j + 1) = b
        End If

    Next j
    i = i + 1
Wend

Label2.Caption = a(1) & Space(4) & a(2) & Space(4) & a(3) & Space(4) & a(4) & Space(4) & a(5) & Space(4) & a(6) & Space(4)
End Sub
