VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   3675
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command4 
      Caption         =   "�M��"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�̤p��"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�̤j��"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��J�ƭ�"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "���浲�G�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Private Sub Command1_Click()

Dim i As Integer
Dim a(1 To 5) As String

a(1) = InputBox("�п�J��1�Ӧ��Z", " ��J�ƭ�")
a(2) = InputBox("�п�J��2�Ӧ��Z", " ��J�ƭ�")
a(3) = InputBox("�п�J��3�Ӧ��Z", " ��J�ƭ�")
a(4) = InputBox("�п�J��4�Ӧ��Z", " ��J�ƭ�")
a(5) = InputBox("�п�J��5�Ӧ��Z", " ��J�ƭ�")

n = Int(Rnd(0 * 5) + 1)

For i = 1 To 5
   List1.AddItem a(n)
Next i

End Sub

Private Sub Command4_Click()
List1.Clear
End Sub
