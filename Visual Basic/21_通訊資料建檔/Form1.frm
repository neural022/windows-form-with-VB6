VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�q�T��ƫ���-����"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   5265
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�x�s"
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   2640
      Width           =   855
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   3480
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   2280
      TabIndex        =   11
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1200
      TabIndex        =   10
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label8 
      Height          =   975
      Left            =   360
      TabIndex        =   15
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "��"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "��"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "�~"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "����q�ܡG"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "e-mail�G"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "�ͤ�G"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�m�W�G"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d(1 To 50, 1 To 4) As String
Dim n As Integer

Private Sub Command1_Click()
n = n + 1

d(n, 1) = Text1.Text
d(n, 2) = Combo1.Text + "/" + Combo2.Text + "/" + Combo3.Text
d(n, 3) = Text2.Text
d(n, 4) = Text3.Text

Label8.Caption = "��" & n & "����ƿ�J����"

End Sub

Private Sub Command2_Click()

Dim e As Integer

Text1.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Text2.Text = ""
Text3.Text = ""
Label8.Caption = ""

e = InputBox("�п�J���d�߲ĴX�����", "�d�߸�ƿ�J")

Label8.Caption = "�z�d�ߪ���" & e & "����Ƭ��G" & Chr(10) & "�m�W" & d(e, 1) & Chr(10) & "�ͤ�G" & d(e, 2) & Chr(10) & "e-mail�G" & d(e, 3) & Chr(10) & "����q�ܡG" & d(e, 4)

End Sub

Private Sub Form_Load()
Dim a, b, c As Integer

For a = 65 To 75
    Combo1.AddItem Str(a)
Next a

For b = 1 To 12
   Combo2.AddItem Str(b)
Next b

For c = 1 To 31
   Combo3.AddItem Str(c)
Next c

End Sub
