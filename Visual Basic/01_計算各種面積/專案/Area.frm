VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���� �����n�p��-����"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   3645
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command3 
      Caption         =   "End"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�M��"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���n"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "����"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "����"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "�e"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "��"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label5.Caption = "����έ��n��" & Val(Text1.Text) * Val(Text2.Text) & "����"
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Label5.Caption = ""
End Sub

Private Sub Command3_Click()
End
End Sub
