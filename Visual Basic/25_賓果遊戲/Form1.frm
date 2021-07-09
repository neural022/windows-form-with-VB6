VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "兆炫的賓果遊戲"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6255
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "Per fect"
      Height          =   495
      Left            =   360
      TabIndex        =   28
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "開始遊戲-2"
      Height          =   615
      Left            =   1800
      TabIndex        =   27
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   24
      Left            =   4800
      TabIndex        =   26
      Top             =   3360
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   23
      Left            =   3720
      TabIndex        =   25
      Top             =   3360
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   22
      Left            =   2640
      TabIndex        =   24
      Top             =   3360
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   21
      Left            =   1560
      TabIndex        =   23
      Top             =   3360
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   20
      Left            =   480
      TabIndex        =   22
      Top             =   3360
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   19
      Left            =   4800
      TabIndex        =   21
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   18
      Left            =   3720
      TabIndex        =   20
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   17
      Left            =   2640
      TabIndex        =   19
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   16
      Left            =   1560
      TabIndex        =   18
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   15
      Left            =   480
      TabIndex        =   17
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   14
      Left            =   4800
      TabIndex        =   16
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   13
      Left            =   3720
      TabIndex        =   15
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   12
      Left            =   2640
      TabIndex        =   14
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   11
      Left            =   1560
      TabIndex        =   13
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   10
      Left            =   480
      TabIndex        =   12
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   9
      Left            =   4800
      TabIndex        =   11
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   8
      Left            =   3720
      TabIndex        =   10
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   7
      Left            =   2640
      TabIndex        =   9
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   6
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   5
      Left            =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   4
      Left            =   4800
      TabIndex        =   6
      Top             =   480
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   3
      Left            =   3720
      TabIndex        =   5
      Top             =   480
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   2
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開始遊戲-1"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1000
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   4320
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click(Index As Integer)

If Check1(0).Value = 1 And Check1(1).Value = 1 And Check1(2).Value = 1 And Check1(3).Value = 1 And Check1(4).Value = 1 Then
Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If

If Check1(5).Value = 1 And Check1(6).Value = 1 And Check1(7).Value = 1 And Check1(8).Value = 1 And Check1(9).Value = 1 Then
Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If

If Check1(10).Value = 1 And Check1(11).Value = 1 And Check1(12).Value = 1 And Check1(13).Value = 1 And Check1(14).Value = 1 Then
Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If

If Check1(15).Value = 1 And Check1(16).Value = 1 And Check1(17).Value = 1 And Check1(18).Value = 1 And Check1(19).Value = 1 Then
Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If

If Check1(20).Value = 1 And Check1(21).Value = 1 And Check1(22).Value = 1 And Check1(23).Value = 1 And Check1(24).Value = 1 Then
Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If




If Check1(0).Value = 1 And Check1(5).Value = 1 And Check1(10).Value = 1 And Check1(15).Value = 1 And Check1(20).Value = 1 Then
Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If

If Check1(1).Value = 1 And Check1(6).Value = 1 And Check1(11).Value = 1 And Check1(16).Value = 1 And Check1(21).Value = 1 Then
Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If

If Check1(2).Value = 1 And Check1(7).Value = 1 And Check1(12).Value = 1 And Check1(17).Value = 1 And Check1(22).Value = 1 Then
Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If

If Check1(3).Value = 1 And Check1(8).Value = 1 And Check1(13).Value = 1 And Check1(18).Value = 1 And Check1(23).Value = 1 Then
Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If

If Check1(4).Value = 1 And Check1(9).Value = 1 And Check1(14).Value = 1 And Check1(19).Value = 1 And Check1(24).Value = 1 Then
Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If




If Check1(0).Value = 1 And Check1(6).Value = 1 And Check1(12).Value = 1 And Check1(18).Value = 1 And Check1(24).Value = 1 Then
Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If

If Check1(20).Value = 1 And Check1(16).Value = 1 And Check1(12).Value = 1 And Check1(8).Value = 1 And Check1(4).Value = 1 Then

Label1.ForeColor = &HFF&
Label1.Caption = "賓果!"
End If

End Sub

Private Sub Command1_Click()

Check1(0).Value = 0
Check1(1).Value = 0
Check1(2).Value = 0
Check1(3).Value = 0
Check1(4).Value = 0
Check1(5).Value = 0
Check1(6).Value = 0
Check1(7).Value = 0
Check1(8).Value = 0
Check1(9).Value = 0
Check1(10).Value = 0
Check1(11).Value = 0
Check1(12).Value = 0
Check1(13).Value = 0
Check1(14).Value = 0
Check1(15).Value = 0
Check1(16).Value = 0
Check1(17).Value = 0
Check1(18).Value = 0
Check1(19).Value = 0
Check1(20).Value = 0
Check1(21).Value = 0
Check1(22).Value = 0
Check1(23).Value = 0
Check1(24).Value = 0
Label1.ForeColor = &HFF&
Label1.Caption = ""

Dim a(1 To 25) As Integer

Randomize
  
  
For i = 1 To 25

b:   a(i) = Int(Rnd() * 25) + 1
  Check1(i - 1).Caption = a(i)
  
  If i = 1 Then
  Check1(i - 1).Caption = a(i)
  Else
    For j = 1 To i - 1
  
      Do Until a(i) <> a(j)
       GoTo b
      Loop
      
    Next j
   End If
    Check1(i - 1).Caption = a(i)
    
Next i

End Sub

Private Sub Command2_Click()


Check1(0).Value = 0
Check1(1).Value = 0
Check1(2).Value = 0
Check1(3).Value = 0
Check1(4).Value = 0
Check1(5).Value = 0
Check1(6).Value = 0
Check1(7).Value = 0
Check1(8).Value = 0
Check1(9).Value = 0
Check1(10).Value = 0
Check1(11).Value = 0
Check1(12).Value = 0
Check1(13).Value = 0
Check1(14).Value = 0
Check1(15).Value = 0
Check1(16).Value = 0
Check1(17).Value = 0
Check1(18).Value = 0
Check1(19).Value = 0
Check1(20).Value = 0
Check1(21).Value = 0
Check1(22).Value = 0
Check1(23).Value = 0
Check1(24).Value = 0
Label1.ForeColor = &HFF&
Label1.Caption = ""



Dim a(1 To 25) As Integer

Randomize

For i = 1 To 25
  j = 1
  
  While j = 1

    a(i) = Int(Rnd() * 25) + 1
     j = 0
   
         For k = 1 To i - 1
       
          If a(k) = a(i) Then
           j = 1
           End If
         
         Next k

  Wend

Next i

Check1(0).Caption = a(1)
Check1(1).Caption = a(2)
Check1(2).Caption = a(3)
Check1(3).Caption = a(4)
Check1(4).Caption = a(5)
Check1(5).Caption = a(6)
Check1(6).Caption = a(7)
Check1(7).Caption = a(8)
Check1(8).Caption = a(9)
Check1(9).Caption = a(10)
Check1(10).Caption = a(11)
Check1(11).Caption = a(12)
Check1(12).Caption = a(13)
Check1(13).Caption = a(14)
Check1(14).Caption = a(15)
Check1(15).Caption = a(16)
Check1(16).Caption = a(17)
Check1(17).Caption = a(18)
Check1(18).Caption = a(19)
Check1(19).Caption = a(20)
Check1(20).Caption = a(21)
Check1(21).Caption = a(22)
Check1(22).Caption = a(23)
Check1(23).Caption = a(24)
Check1(24).Caption = a(25)
End Sub

Private Sub Command3_Click()

For x = 0 To 24
Check1(x).Value = 0
Next x
Label1.ForeColor = &HFF&
Label1.Caption = ""

Dim a(1 To 25) As Integer

Randomize

For i = 1 To 25

Do

a(i) = Int(Rnd() * 25) + 1

k = 0
    For j = 1 To i - 1
    
        If a(j) = a(i) Then
        k = 1
        End If
    Next j
    
Loop Until k = 0

Next i

For y = 0 To 24
Check1(y).Caption = a(y + 1)
Next y

End Sub
