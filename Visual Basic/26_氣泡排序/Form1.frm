VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "排序(大到小)"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "排序(小到大)"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 5), b As Integer
Private Sub Command1_Click()

a(1) = 20
a(2) = 5
a(3) = 30
a(4) = 40
a(5) = 15


i = 1
While i < 5

    For j = 1 To (5 - i)
    
      If a(j) > a(j + 1) Then
        b = a(j)
        a(j) = a(j + 1)
        a(j + 1) = b
      End If
    
    Next j
    
    
i = i + 1
Wend
Label1.Caption = a(1) & Space(4) & a(2) & Space(4) & a(3) & Space(4) & a(4) & Space(4) & a(5)


End Sub


Private Sub Command2_Click()
a(1) = 20
a(2) = 5
a(3) = 30
a(4) = 40
a(5) = 15

i = 1
While i < 5

    For j = 1 To (5 - i)
    
      If a(j) < a(j + 1) Then
       b = a(j)
       a(j) = a(j + 1)
       a(j + 1) = b
      End If
    
    Next j
    
    
i = i + 1
Wend
Label2.Caption = a(1) & Space(4) & a(2) & Space(4) & a(3) & Space(4) & a(4) & Space(4) & a(5)


End Sub
