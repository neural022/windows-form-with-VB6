VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���n�p�⾹2-������"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   5145
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Left            =   3960
      TabIndex        =   17
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�M��"
      Height          =   495
      Left            =   3960
      TabIndex        =   16
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�p��"
      Height          =   495
      Left            =   3960
      TabIndex        =   15
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.OptionButton Option6 
      Caption         =   "���"
      Height          =   250
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton Option5 
      Caption         =   "����|���"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   1355
   End
   Begin VB.OptionButton Option4 
      Caption         =   "���"
      Height          =   250
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "�T����"
      Height          =   250
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "�����"
      Height          =   250
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "�����"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "2�B�п�J�ƭ�"
      Height          =   495
      Left            =   600
      TabIndex        =   20
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "1�B�п�ܭn�p�⪺�ϧ�"
      Height          =   375
      Left            =   600
      TabIndex        =   19
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   720
      TabIndex        =   18
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
Label7.Caption = "����έ��n" & Val(Text1.Text) * Val(Text2.Text) & "���褽��"
End If

If Option2.Value = True Then
Label7.Caption = "����έ��n" & Val(Text1.Text) ^ 2 & "���褽��"
End If

If Option3.Value = True Then
Label7.Caption = "�T���έ��n" & Val(Text1.Text) * Val(Text2.Text) / 2 & "���褽��"
End If

If Option4.Value = True Then
Label7.Caption = "��έ��n" & (Val(Text1.Text) + Val(Text2.Text)) * Val(Text3.Text) / 2 & "���褽��"
End If

If Option5.Value = True Then
Label7.Caption = "����|��έ��n" & Val(Text1.Text) * Val(Text2.Text) & "���褽��"
End If

If Option6.Value = True Then
Label7.Caption = "��έ��n" & Val(Text1.Text) ^ 2 * 3.14 & "���褽��"
End If

End Sub

Private Sub Command2_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Option1_Click()
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

If Option1.Value = True Then
Label1.Caption = "��"
Label2.Caption = "����"
Label3.Caption = "�e"
Label4.Caption = "����"
End If

End Sub


Private Sub Option2_Click()
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

If Option2.Value = True Then
Label1.Caption = "���"
Label2.Caption = "����"
End If
End Sub

Private Sub Option3_Click()
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

If Option3.Value = True Then
Label1.Caption = "��"
Label2.Caption = "����"
Label3.Caption = "��"
Label4.Caption = "����"
End If
End Sub

Private Sub Option4_Click()
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

If Option4.Value = True Then
Label1.Caption = "�W��"
Label2.Caption = "����"
Label3.Caption = "�U��"
Label4.Caption = "����"
Label5.Caption = "��"
Label6.Caption = "����"
End If
End Sub

Private Sub Option5_Click()
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

If Option5.Value = True Then
Label1.Caption = "��"
Label2.Caption = "����"
Label3.Caption = "��"
Label4.Caption = "����"
End If
End Sub

Private Sub Option6_Click()
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

If Option6.Value = True Then
Label1.Caption = "�b�|"
Label2.Caption = "����"
End If
End Sub
