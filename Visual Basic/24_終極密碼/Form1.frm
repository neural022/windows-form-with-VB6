VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�������׷��K�X"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "�}�l�C��"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�]�w�Ʀr"
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
      Caption         =   "�C�������G���]�w�Ʀr�A�A�����u�}�l�C���v�q�K�X�C"
      BeginProperty Font 
         Name            =   "�s�ө���"
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

a(1) = InputBox("�п�J�@��1~99����", "�q�Ʀr�C��-�]�w�Ʀr")

End Sub

Private Sub Command2_Click()

Dim c, d As Integer

c = 1
d = 99
b = InputBox("�п�J" & c & "~" & d & "����", "�׷��K�X-�q�Ʀr��..")

If b = a(1) Then
   MsgBox "���ߧA�q��F", vbInformation, "�׷��K�X�q�Ʀr"
 Else
 
    Do
     If b > a(1) Then
       d = b
       b = InputBox("�п�J" & c & "~" & d & "����", "�׷��K�X-�q�Ʀr��..")
      Else
        c = b
         b = InputBox("�п�J" & c & "~" & d & "����", "�׷��K�X-�q�Ʀr��..")
      End If
    Loop Until b = a(1)
      MsgBox "���ߧA�q��F", vbInformation, "�׷��K�X�q�Ʀr"
    
End If

End Sub
