VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���Z�P�_-����2"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4560
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�T�{"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "��J���Z"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim score As Integer
score = Val(Text1.Text)
If score < 0 Then

Label2.Caption = "�̧C��0���A�L�����ơA�Э��s��J"

ElseIf score < 60 Then
 Label2.Caption = "���Z���ή�"
ElseIf score < 70 Then
 Label2.Caption = "���Z����"
ElseIf score < 80 Then
 Label2.Caption = "���Z�A��"
ElseIf score < 90 Then
 Label2.Caption = "���Z�ҵ�"
ElseIf score <= 100 Then
Label2.Caption = "���Z�u��"

ElseIf score > 100 Then
Label2.Caption = "������100���A�L�����ơA�Э��s��J"

End If
End Sub

Private Sub Command2_Click()
End
End Sub
