VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���B-����"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4755
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   4455
      Begin VB.Label Label3 
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�d��"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "�п�ܰW���G"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "�п�ܰ_���G"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�ۭq��Ʀ��^�ǭ�,�P�@���(Form)�U���i���ۦP�ۭq��ƦW��

Dim x, y, z As Integer
Dim a(1 To 28) As Single
Private Sub Command1_Click()
 Dim w, b As Single
    
 If Combo1.Text <> "" And Combo2.Text <> "" Then
    w = mileage_1(b)
    w = mileage_2(b)
    w = mileage_3(b)
    w = mileage_4(b)
    w = mileage_5(b)
    w = mileage_6(b)
    
Rem �ھڨ��{�P�_����
If w < 5 Then
    x = 20
    y = 16
    z = 8
ElseIf w < 8 Then
    x = 25
    y = 20
    z = 10
ElseIf w < 11 Then
    x = 30
    y = 24
    z = 12
ElseIf w < 14 Then
    x = 35
    y = 28
    z = 14
ElseIf w < 17 Then
    x = 40
    y = 32
    z = 16
ElseIf w < 20 Then
    x = 45
    y = 36
    z = 18
ElseIf w < 23 Then
    x = 50
    y = 40
    z = 20
ElseIf w < 27 Then
    x = 55
    y = 44
    z = 22
ElseIf w < 31 Then
    x = 60
    y = 48
    z = 24
Else
    x = 65
    y = 48
    z = 28

End If

Label3.Caption = "��{���G" & x & "��" & Chr(10) & Chr(10) & "�y�C�d/�@�d�q�G" & y & "��" & Chr(10) & Chr(10) & "�q�ѥd�B�R�ߥd�B�R�߳���d�G" & z & "��"

Else
    Combo1.Text = ""
    Combo2.Text = ""
    MsgBox "�п�J�_�W��", vbInformation, "���B�����d��"
End If


End Sub

Function mileage_1(b As Single) As Single

Rem �_���G1�H��
If Combo1.Text = "�H��" Then
    If Combo2.Text = "�H��" Then
        b = a(1)
    ElseIf Combo2.Text = "����L" Then
        b = a(1) + a(2)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(1) + a(2) + a(3)
    ElseIf Combo2.Text = "����" Then
        b = a(1) + a(2) + a(3) + a(4)
    ElseIf Combo2.Text = "���q" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "���w" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
        
Rem �_���G2����L
If Combo1.Text = "����L" Then
    If Combo2.Text = "�H��" Then
        b = a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(1)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(1) + a(3)
    ElseIf Combo2.Text = "����" Then
        b = a(1) + a(3) + a(4)
    ElseIf Combo2.Text = "���q" Then
        b = a(1) + a(3) + a(4) + a(5)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "���w" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
       
Rem �_���G3�˳�
If Combo1.Text = "�˳�" Then
    If Combo2.Text = "�H��" Then
        b = a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(1)
    ElseIf Combo2.Text = "����" Then
        b = a(1) + a(4)
    ElseIf Combo2.Text = "���q" Then
        b = a(1) + a(4) + a(5)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(1) + a(4) + a(5) + a(6)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "���w" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
       
Rem �_���G4����
If Combo1.Text = "����" Then
    If Combo2.Text = "�H��" Then
        b = a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(1)
    ElseIf Combo2.Text = "���q" Then
        b = a(1) + a(5)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(1) + a(5) + a(6)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(5) + a(6) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "���w" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
        
Rem �_���G5���q
If Combo1.Text = "���q" Then
    If Combo2.Text = "�H��" Then
        b = a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(1)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(1) + a(6)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(6) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(6) + a(7) + a(8)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "���w" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If

mileage_1 = b
    
End Function


Function mileage_2(b As Single) As Single

Rem �_���G6�_���^
If Combo1.Text = "�_���^" Then
    If Combo2.Text = "�H��" Then
        b = a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(1)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(7) + a(8)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(1) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "���w" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
        
Rem �_���G7�_��
If Combo1.Text = "�_��" Then
    If Combo2.Text = "�H��" Then
        b = a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1) + a(8)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(1) + a(8) + a(9)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(1) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "���w" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If

Rem �_���G8�_��
If Combo1.Text = "�_��" Then
    If Combo2.Text = "�H��" Then
        b = a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(1)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(1) + a(9)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(1) + a(9) + a(10)
    ElseIf Combo2.Text = "���w" Then
        b = a(1) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
        
        
Rem �_���G9ԧ����
If Combo1.Text = "ԧ����" Then
    If Combo2.Text = "�H��" Then
        b = a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(1)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(1) + a(10)
    ElseIf Combo2.Text = "���w" Then
        b = a(1) + a(10) + a(11)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
        
Rem �_���G10�۵P
If Combo1.Text = "�۵P" Then
    If Combo2.Text = "�H��" Then
        b = a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(1)
    ElseIf Combo2.Text = "���w" Then
        b = a(1) + a(11)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1) + a(11) + a(12)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If

mileage_2 = b

End Function


Function mileage_3(b As Single) As Single
        
Rem �_���G11���w
If Combo1.Text = "���w" Then
    If Combo2.Text = "�H��" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(1)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1) + a(12)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(12) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If



Rem �_���G12�ۤs
If Combo1.Text = "�ۤs" Then
    If Combo2.Text = "�H��" Then
        b = a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = (12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = (12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = (12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = (12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = (12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = (12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = (12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = (12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = (12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(1)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1) + a(13)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(13) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem �_���G13�h�L
If Combo1.Text = "�h�L" Then
    If Combo2.Text = "�H��" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(1)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1) + a(14)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(14) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem �_���G14�C��
If Combo1.Text = "�C��" Then
    If Combo2.Text = "�H��" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(14) + a(13) + (12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(14) + a(13) + (12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(1)
    ElseIf Combo2.Text = "��s" Then
        b = a(1) + a(15)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(15) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If



Rem �_���G15��s
If Combo1.Text = "��s" Then
    If Combo2.Text = "�H��" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(1)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1) + a(16)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(16) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If

mileage_3 = b

End Function


Function mileage_4(b As Single) As Single

Rem �_���G16���v���
If Combo1.Text = "���v���" Then
    If Combo2.Text = "�H��" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(1)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(17) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem �_���G17���s
If Combo1.Text = "���s" Then
    If Combo2.Text = "�H��" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(17) + a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(1)
    ElseIf Combo2.Text = "���s" Then
        b = a(1) + a(18)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(18) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem �_���G18���s
If Combo1.Text = "���s" Then
    If Combo2.Text = "�H��" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(18) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(18)
    ElseIf Combo2.Text = "���s" Then
        b = a(1)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1) + a(19)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(19) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem �_���G19�x�_����
If Combo1.Text = "�x�_����" Then
    If Combo2.Text = "�H��" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(19) + a(18)
    ElseIf Combo2.Text = "���s" Then
        b = a(19)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(1)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1) + a(20)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(20) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If

Rem �_���G20�x�j��|
If Combo1.Text = "�x�j��|" Then
    If Combo2.Text = "�H��" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "���s" Then
        b = a(20) + a(19)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(20)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(1)
    ElseIf Combo2.Text = "����������" Then
        b = a(1) + a(21)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(21) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


mileage_4 = b

End Function


Function mileage_5(b As Single) As Single

Rem �_���G21����������
If Combo1.Text = "����������" Then
    If Combo2.Text = "�H��" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "���s" Then
        b = a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(21) + a(20)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(21)
    ElseIf Combo2.Text = "����������" Then
        b = a(1)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1) + a(22)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(22) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If




Rem �_���G22�F��
If Combo1.Text = "�F��" Then
    If Combo2.Text = "�H��" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "���s" Then
        b = a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(22) + a(21)
    ElseIf Combo2.Text = "����������" Then
        b = a(22)
    ElseIf Combo2.Text = "�F��" Then
        b = a(1)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1) + a(23)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(23) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem �_���G23�j�w�˪L����
If Combo1.Text = "�j�w�˪L����" Then
    If Combo2.Text = "�H��" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "���s" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(23) + a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(23) + a(22) + a(21)
    ElseIf Combo2.Text = "����������" Then
        b = a(23) + a(22)
    ElseIf Combo2.Text = "�F��" Then
        b = a(23)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(1)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1) + a(24)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(24) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem �_���G24�j�w
If Combo1.Text = "�j�w" Then
    If Combo2.Text = "�H��" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "���s" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(24) + a(23) + a(22) + a(21)
    ElseIf Combo2.Text = "����������" Then
        b = a(24) + a(23) + a(22)
    ElseIf Combo2.Text = "�F��" Then
        b = a(24) + a(23)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(24)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(1)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1) + a(25)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(25) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(25) + a(26) + a(27)
    End If
End If


Rem �_���G25�H�q�w�M
If Combo1.Text = "�H�q�w�M" Then
    If Combo2.Text = "�H��" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "���s" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21)
    ElseIf Combo2.Text = "����������" Then
        b = a(25) + a(24) + a(23) + a(22)
    ElseIf Combo2.Text = "�F��" Then
        b = a(25) + a(24) + a(23)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(25) + a(24)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(25)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(1)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1) + a(26)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(26) + a(27)
    End If
End If



mileage_5 = b

End Function


Function mileage_6(b As Single) As Single
Rem �_���G26�x�_101/�@�T
If Combo1.Text = "�x�_101/�@�T" Then
    If Combo2.Text = "�H��" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + (12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "���s" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21)
    ElseIf Combo2.Text = "����������" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22)
    ElseIf Combo2.Text = "�F��" Then
        b = a(26) + a(25) + a(24) + a(23)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(26) + a(25) + a(24)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(26) + a(25)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(26)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(1)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1) + a(27)
    End If
End If


Rem �_���G27�H�s
If Combo1.Text = "�H�s" Then
    If Combo2.Text = "�H��" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "����L" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "�˳�" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "����" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "���q" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "�_���^" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "�_��" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "�_��" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "ԧ����" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "�۵P" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "���w" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "�ۤs" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "�h�L" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "�C��" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "��s" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "���v���" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "���s" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "���s" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "�x�_����" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "�x�j��|" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21)
    ElseIf Combo2.Text = "����������" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22)
    ElseIf Combo2.Text = "�F��" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23)
    ElseIf Combo2.Text = "�j�w�˪L����" Then
        b = a(27) + a(26) + a(25) + a(24)
    ElseIf Combo2.Text = "�j�w" Then
        b = a(27) + a(26) + a(25)
    ElseIf Combo2.Text = "�H�q�w�M" Then
        b = a(27) + a(26)
    ElseIf Combo2.Text = "�x�_101/�@�T" Then
        b = a(27)
    ElseIf Combo2.Text = "�H�s" Then
        b = a(1)
    End If
End If



mileage_6 = b

End Function





Private Sub Form_Load()

Rem �H���H�q�u�U���W��

Combo1.AddItem "�H��"
Combo1.AddItem "����L"
Combo1.AddItem "�˳�"

Combo1.AddItem "����"
Combo1.AddItem "���q"

Combo1.AddItem "�_���^"
Combo1.AddItem "�_��"
Combo1.AddItem "�_��"

Combo1.AddItem "ԧ����"
Combo1.AddItem "�۵P"
Combo1.AddItem "���w"

Combo1.AddItem "�ۤs"
Combo1.AddItem "�h�L"
Combo1.AddItem "�C��"

Combo1.AddItem "��s"
Combo1.AddItem "���v���"
Combo1.AddItem "���s"

Combo1.AddItem "���s"
Combo1.AddItem "�x�_����"
Combo1.AddItem "�x�j��|"
Combo1.AddItem "����������"

Combo1.AddItem "�F��"
Combo1.AddItem "�j�w�˪L����"
Combo1.AddItem "�j�w"
Combo1.AddItem "�H�q�w�M"
Combo1.AddItem "�x�_101/�@�T"

Combo1.AddItem "�H�s"





Combo2.AddItem "�H��"
Combo2.AddItem "����L"
Combo2.AddItem "�˳�"

Combo2.AddItem "����"
Combo2.AddItem "���q"

Combo2.AddItem "�_���^"
Combo2.AddItem "�_��"
Combo2.AddItem "�_��"

Combo2.AddItem "ԧ����"
Combo2.AddItem "�۵P"
Combo2.AddItem "���w"

Combo2.AddItem "�ۤs"
Combo2.AddItem "�h�L"
Combo2.AddItem "�C��"

Combo2.AddItem "��s"
Combo2.AddItem "���v���"
Combo2.AddItem "���s"

Combo2.AddItem "���s"
Combo2.AddItem "�x�_����"
Combo2.AddItem "�x�j��|"
Combo2.AddItem "����������"

Combo2.AddItem "�F��"
Combo2.AddItem "�j�w�˪L����"
Combo2.AddItem "�j�w"
Combo2.AddItem "�H�q�w�M"
Combo2.AddItem "�x�_101/�@�T"

Combo2.AddItem "�H�s"

Rem �H���H�q�u�C�@�����Z(���t�s�_��)

a(1) = 0
a(2) = 2.07
a(3) = 1.93

a(4) = 2.05
a(5) = 0.88

a(6) = 1.44
a(7) = 1.62
a(8) = 0.76

a(9) = 0.86
a(10) = 1.24
a(11) = 0.61

a(12) = 0.88
a(13) = 0.98
a(14) = 1.19

a(15) = 1.52
a(16) = 1.02
a(17) = 0.56

a(18) = 0.54
a(19) = 0.63
a(20) = 0.63
a(21) = 0.95

a(22) = 1.2
a(23) = 0.66
a(24) = 0.8
a(25) = 0.79
a(26) = 1.08

a(27) = 0.86


End Sub
