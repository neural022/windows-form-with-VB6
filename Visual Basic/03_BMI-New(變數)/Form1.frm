VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "BMI�p��(If Else��)-����"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   3795
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���s��J"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�p��"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "�k"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "�k"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "�ʧO�G"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "�魫(����)�G"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "����(����)�G"
      Height          =   255
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
Private Sub Command1_Click()
Dim a As Single
a = Val(Text2.Text) / ((Val(Text1.Text) / 100) ^ 2)
If Option1.Value = True Then
   If a > 27.8 Then
     Label4.Caption = "�z��BMI��" & a & "�έD"
   Else
     Label4.Caption = "�z��BMI��" & a & "�@���ݩ�"
   End If
End If

If Option2.Value = True Then
   If a > 27.3 Then
     Label4.Caption = "�z��BMI��" & a & "�έD"
   Else
     Label4.Caption = "�z��BMI��" & a & "�@���ݩ�"
   End If
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Option1.Value = False
Option2.Value = False
Label4.Caption = ""
End Sub

Private Sub Command3_Click()
End
End Sub
