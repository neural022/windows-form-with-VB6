VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�P�y�d��-����"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   4140
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "���s�d��"
      Height          =   255
      Left            =   1320
      TabIndex        =   19
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "�Q�G�P�y"
      Height          =   2775
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   3495
      Begin VB.OptionButton Option12 
         Caption         =   "�����y"
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         Caption         =   "���~�y"
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option10 
         Caption         =   "���~�y"
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option9 
         Caption         =   "�g��y"
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option8 
         Caption         =   "���Ȯy"
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option7 
         Caption         =   "�ѯ��y"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         Caption         =   "�B�k�y"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "��l�y"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "���ɮy"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "���l�y"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "�����y"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�d�Ϯy"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�d��"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "��"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "��"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "�п�J�X�ͤ���G"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b As Integer

a = Val(Text1.Text)
b = Val(Text2.Text)

 If a > 12 Or a < 1 Then
 
            Label4.Caption = "�S������A�Э��s�d��"
            Option1.Value = False
            Option2.Value = False
            Option3.Value = False
            Option4.Value = False
            Option5.Value = False
            Option6.Value = False
            Option7.Value = False
            Option8.Value = False
            Option9.Value = False
            Option10.Value = False
            Option11.Value = False
            Option12.Value = False

 Else

            If b > 31 Or b < 1 Then
                Label4.Caption = a & "��S���Ӥ���A�Э��s�d��"
                Option1.Value = False
                Option2.Value = False
                Option3.Value = False
                Option4.Value = False
                Option5.Value = False
                Option6.Value = False
                Option7.Value = False
                Option8.Value = False
                Option9.Value = False
                Option10.Value = False
                Option11.Value = False
                Option12.Value = False

            Else

                Select Case a

                Case 4, 6, 9, 11
                    If b > 30 Then
                        Label4.Caption = a & "��S���Ӥ���A�Э��s�d��"
        
                        Option1.Value = False
                        Option2.Value = False
                        Option3.Value = False
                        Option4.Value = False
                        Option5.Value = False
                        Option6.Value = False
                        Option7.Value = False
                        Option8.Value = False
                        Option9.Value = False
                        Option10.Value = False
                        Option11.Value = False
                        Option12.Value = False
        
                    Else
        
                         Select Case a
        
                            Case 4
                                   If b >= 21 Then
                                     Option2.Value = True
                                   Else
                                     Option1.Value = True
                                   End If
                  
                            Case 6
                                   If b >= 21 Then
                                     Option4.Value = True
                                   Else
                                     Option3.Value = True
                                   End If
                 
                            Case 9
                                   If b >= 23 Then
                                     Option7.Value = True
                                   Else
                                     Option6.Value = True
                                   End If
                  
                            Case 11
                                   If b >= 22 Then
                                     Option9.Value = True
                                   Else
                                     Option8.Value = True
                                   End If
                  
                         End Select
                         
                    End If
        
                Case 1, 2, 3, 5, 7, 8, 10, 12
                   If a = 2 And b > 29 Then
                        Label4.Caption = a & "��S���Ӥ���A�Э��s�d��"
            
                        Option1.Value = False
                        Option2.Value = False
                        Option3.Value = False
                        Option4.Value = False
                        Option5.Value = False
                        Option6.Value = False
                        Option7.Value = False
                        Option8.Value = False
                        Option9.Value = False
                        Option10.Value = False
                        Option11.Value = False
                        Option12.Value = False

                   Else
          
          
                        Select Case a
     
                            Case 1
     
                                   If b >= 20 Then
                                     Option11.Value = True
                                   Else
                                     Option10.Value = True
                                   End If
     
                            Case 2
   
                                   If b >= 20 Then
                                     Option12.Value = True
                                   Else
                                     Option11.Value = True
                                   End If
     
                            Case 3
                    
                                   If b >= 20 Then
                                     Option1.Value = True
                                   Else
                                     Option12.Value = True
                                   End If
     
                            Case 5
                
                                   If b >= 21 Then
                                     Option3.Value = True
                                   Else
                                     Option2.Value = True
                                   End If
         
                            Case 7
                
                                   If b >= 23 Then
                                     Option5.Value = True
                                   Else
                                     Option4.Value = True
                                   End If
   
                            Case 8
                  
                                   If b >= 23 Then
                                     Option6.Value = True
                                   Else
                                     Option5.Value = True
                                   End If
   
                            Case 10
                 
                                   If b >= 23 Then
                                     Option8.Value = True
                                   Else
                                     Option7.Value = True
                                   End If
     
                            Case 12
                
                                   If b >= 22 Then
                                     Option10.Value = True
                                   Else
                                     Option9.Value = True
                                   End If
                        End Select
                     End If
                End Select
    
                If Option1.Value = True Then Label4.Caption = "�d�Ϯy��3/21~4/20"
                If Option2.Value = True Then Label4.Caption = "�����y��4/21~5/21"
                If Option3.Value = True Then Label4.Caption = "���l�y��5/21~6/21"
                If Option4.Value = True Then Label4.Caption = "���ɮy��6/22~7/22"
                If Option5.Value = True Then Label4.Caption = "��l�y��7/23~8/22"
                If Option6.Value = True Then Label4.Caption = "�B�k�y��8/23~9/22"
                If Option7.Value = True Then Label4.Caption = "�ѯ��y��9/23~10/22"
                If Option8.Value = True Then Label4.Caption = "���Ȯy��10/23~11/21"
                If Option9.Value = True Then Label4.Caption = "�g��y��11/22~12/21"
                If Option10.Value = True Then Label4.Caption = "�]�~�y��12/22~1/19"
                If Option11.Value = True Then Label4.Caption = "���~�y��1/20~2/19"
                If Option12.Value = True Then Label4.Caption = "�����y��2/20~3/20"
    
            End If
 End If

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Label4.Caption = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Option9.Value = False
Option10.Value = False
Option11.Value = False
Option12.Value = False
End Sub


Private Sub Option1_Click()
Label4.Caption = ""

If Option1.Value = True Then Label4.Caption = "�d�Ϯy��3/21~4/20"
End Sub

Private Sub Option2_Click()
Label4.Caption = ""

If Option2.Value = True Then Label4.Caption = "�����y��4/21~5/21"
End Sub

Private Sub Option3_Click()
Label4.Caption = ""

If Option3.Value = True Then Label4.Caption = "���l�y��5/21~6/21"
End Sub

Private Sub Option4_Click()
Label4.Caption = ""

If Option4.Value = True Then Label4.Caption = "���ɮy��6/22~7/22"
End Sub

Private Sub Option5_Click()
Label4.Caption = ""

If Option5.Value = True Then Label4.Caption = "��l�y��7/23~8/22"
End Sub

Private Sub Option6_Click()
Label4.Caption = ""

If Option6.Value = True Then Label4.Caption = "�B�k�y��8/23~9/22"
End Sub

Private Sub Option7_Click()
Label4.Caption = ""

If Option7.Value = True Then Label4.Caption = "�ѯ��y��9/23~10/22"
End Sub

Private Sub Option8_Click()
Label4.Caption = ""

If Option8.Value = True Then Label4.Caption = "���Ȯy��10/23~11/21"
End Sub

Private Sub Option9_Click()
Label4.Caption = ""

If Option9.Value = True Then Label4.Caption = "�g��y��11/22~12/21"
End Sub

Private Sub Option10_Click()
Label4.Caption = ""

If Option10.Value = True Then Label4.Caption = "�]�~�y��12/22~1/19"
End Sub

Private Sub Option11_Click()
Label4.Caption = ""

If Option11.Value = True Then Label4.Caption = "���~�y��1/20~2/19"
End Sub

Private Sub Option12_Click()
Label4.Caption = ""

If Option12.Value = True Then Label4.Caption = "�����y��2/20~3/20"
End Sub
