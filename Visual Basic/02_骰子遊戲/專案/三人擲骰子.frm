VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�Y��l-������"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   5910
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command5 
      Caption         =   "����"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���s�}�l"
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�Y"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�Y"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Y"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  '�m�����
      Caption         =   "1�B�Ш̷Ӷ��Ǥ��O�Y��l"
      Height          =   975
      Left            =   1320
      TabIndex        =   11
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  '�m�����
      Caption         =   "��"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  '�m�����
      Caption         =   "�A"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      Caption         =   "��"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer
Dim c As Integer

Private Sub Command1_Click()
a = Int(Rnd() * 6) + 1
Label4.Caption = a

End Sub

Private Sub Command2_Click()
b = Int(Rnd() * 6) + 1
Label5.Caption = b


End Sub

Private Sub Command3_Click()
c = Int(Rnd() * 6) + 1
Label6.Caption = c


If a > b And a > c Then MsgBox "��Ĺ�F"
If b > c And b > a Then MsgBox "�AĹ�F"
If c > a And c > b Then MsgBox "��Ĺ�F"
If a = b And b = c Then MsgBox "����"

If a = b And a > c Then MsgBox "�ҩM�AĹ�F"
If b = c And b > a Then MsgBox "�A�M��Ĺ�F"
If c = a And c > b Then MsgBox "�ҩM��Ĺ�F"

End Sub

Private Sub Command4_Click()
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""

End Sub

Private Sub Command5_Click()
End

End Sub

