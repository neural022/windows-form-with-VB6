VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�p�⭱�n-������"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   6855
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.OptionButton Option6 
      Caption         =   "���"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.OptionButton Option5 
      Caption         =   "����|���"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "���"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "�T����"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "�����"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "�����"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�M��"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�p�⭱�n"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
a1 = InputBox("�п�J��(cm)", "����έ��n")
a2 = InputBox("�п�J�e(cm)", "����έ��n")
a = a1 * a2
MsgBox "����έ��n" & a & "���褽��"
End If

If Option2.Value = True Then
b1 = InputBox("�п�J���(cm)", "����έ��n")
b = b1 ^ 2
MsgBox "����έ��n" & b & "���褽��"
End If

If Option3.Value = True Then
c1 = InputBox("�п�J��(cm)", "�T���έ��n")
c2 = InputBox("�п�J��(cm)", "�T���έ��n")
c = c1 * c2 / 2
MsgBox "�T���έ��n" & c & "���褽��"
End If

If Option4.Value = True Then
d1 = InputBox("�п�J�W��(cm)", "��έ��n")
d2 = InputBox("�п�J�U��(cm)", "��έ��n")
d3 = InputBox("�п�J��(cm)", "��έ��n")
d = (Val(d1) + Val(d2)) * d3 / 2
MsgBox "��έ��n" & d & "���褽��"
End If

If Option5.Value = True Then
e1 = InputBox("�п�J��(cm)", "����|��έ��n")
e2 = InputBox("�п�J��(cm)", "����|��έ��n")
e = e1 * e2
MsgBox "����|��έ��n", "���褽��"
End If

If Option6.Value = True Then
f1 = InputBox("�п�J�b�|(cm)", "��έ��n")
f = f1 ^ 2 * 3.14
MsgBox "��έ��n" & f & "���褽��"
End If
 
End Sub

Private Sub Command2_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
End Sub

Private Sub Command3_Click()
End
End Sub
