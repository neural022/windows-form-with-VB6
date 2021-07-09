VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "捷運-兆炫"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4755
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame1 
      Caption         =   "票價"
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
      Caption         =   "查詢"
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
      Caption         =   "請選擇訖站："
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "請選擇起站："
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
'自訂函數有回傳值,同一表單(Form)下不可有相同自訂函數名稱

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
    
Rem 根據里程判斷票價
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

Label3.Caption = "單程票：" & x & "元" & Chr(10) & Chr(10) & "悠遊卡/一卡通：" & y & "元" & Chr(10) & Chr(10) & "敬老卡、愛心卡、愛心陪伴卡：" & z & "元"

Else
    Combo1.Text = ""
    Combo2.Text = ""
    MsgBox "請輸入起訖站", vbInformation, "捷運票價查詢"
End If


End Sub

Function mileage_1(b As Single) As Single

Rem 起站：1淡水
If Combo1.Text = "淡水" Then
    If Combo2.Text = "淡水" Then
        b = a(1)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(1) + a(2)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(1) + a(2) + a(3)
    ElseIf Combo2.Text = "關渡" Then
        b = a(1) + a(2) + a(3) + a(4)
    ElseIf Combo2.Text = "忠義" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6)
    ElseIf Combo2.Text = "北投" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "石牌" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "明德" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
        
Rem 起站：2紅樹林
If Combo1.Text = "紅樹林" Then
    If Combo2.Text = "淡水" Then
        b = a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(1)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(1) + a(3)
    ElseIf Combo2.Text = "關渡" Then
        b = a(1) + a(3) + a(4)
    ElseIf Combo2.Text = "忠義" Then
        b = a(1) + a(3) + a(4) + a(5)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6)
    ElseIf Combo2.Text = "北投" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "石牌" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "明德" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
       
Rem 起站：3竹圍
If Combo1.Text = "竹圍" Then
    If Combo2.Text = "淡水" Then
        b = a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(1)
    ElseIf Combo2.Text = "關渡" Then
        b = a(1) + a(4)
    ElseIf Combo2.Text = "忠義" Then
        b = a(1) + a(4) + a(5)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(1) + a(4) + a(5) + a(6)
    ElseIf Combo2.Text = "北投" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "石牌" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "明德" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
       
Rem 起站：4關渡
If Combo1.Text = "關渡" Then
    If Combo2.Text = "淡水" Then
        b = a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(1)
    ElseIf Combo2.Text = "忠義" Then
        b = a(1) + a(5)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(1) + a(5) + a(6)
    ElseIf Combo2.Text = "北投" Then
        b = a(1) + a(5) + a(6) + a(7)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "石牌" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "明德" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
        
Rem 起站：5忠義
If Combo1.Text = "忠義" Then
    If Combo2.Text = "淡水" Then
        b = a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(1)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(1) + a(6)
    ElseIf Combo2.Text = "北投" Then
        b = a(1) + a(6) + a(7)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(1) + a(6) + a(7) + a(8)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "石牌" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "明德" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If

mileage_1 = b
    
End Function


Function mileage_2(b As Single) As Single

Rem 起站：6復興崗
If Combo1.Text = "復興崗" Then
    If Combo2.Text = "淡水" Then
        b = a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(1)
    ElseIf Combo2.Text = "北投" Then
        b = a(1) + a(7)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(1) + a(7) + a(8)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(1) + a(7) + a(8) + a(9)
    ElseIf Combo2.Text = "石牌" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "明德" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
        
Rem 起站：7北投
If Combo1.Text = "北投" Then
    If Combo2.Text = "淡水" Then
        b = a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(1)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(1) + a(8)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(1) + a(8) + a(9)
    ElseIf Combo2.Text = "石牌" Then
        b = a(1) + a(8) + a(9) + a(10)
    ElseIf Combo2.Text = "明德" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(8) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If

Rem 起站：8奇岩
If Combo1.Text = "奇岩" Then
    If Combo2.Text = "淡水" Then
        b = a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(1)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(1) + a(9)
    ElseIf Combo2.Text = "石牌" Then
        b = a(1) + a(9) + a(10)
    ElseIf Combo2.Text = "明德" Then
        b = a(1) + a(9) + a(10) + a(11)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(9) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
        
        
Rem 起站：9唭哩岸
If Combo1.Text = "唭哩岸" Then
    If Combo2.Text = "淡水" Then
        b = a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(1)
    ElseIf Combo2.Text = "石牌" Then
        b = a(1) + a(10)
    ElseIf Combo2.Text = "明德" Then
        b = a(1) + a(10) + a(11)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1) + a(10) + a(11) + a(12)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(10) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If
        
Rem 起站：10石牌
If Combo1.Text = "石牌" Then
    If Combo2.Text = "淡水" Then
        b = a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(1)
    ElseIf Combo2.Text = "明德" Then
        b = a(1) + a(11)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1) + a(11) + a(12)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(11) + a(12) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(11) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If

mileage_2 = b

End Function


Function mileage_3(b As Single) As Single
        
Rem 起站：11明德
If Combo1.Text = "明德" Then
    If Combo2.Text = "淡水" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(1)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1) + a(12)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(12) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(12) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(12) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If



Rem 起站：12芝山
If Combo1.Text = "芝山" Then
    If Combo2.Text = "淡水" Then
        b = a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = (12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = (12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = (12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = (12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = (12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = (12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = (12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = (12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = (12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(1)
    ElseIf Combo2.Text = "士林" Then
        b = a(1) + a(13)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(13) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(13) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(13) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem 起站：13士林
If Combo1.Text = "士林" Then
    If Combo2.Text = "淡水" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(1)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1) + a(14)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(14) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(14) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(14) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem 起站：14劍潭
If Combo1.Text = "劍潭" Then
    If Combo2.Text = "淡水" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(14) + a(13) + (12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(14) + a(13) + (12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(1)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1) + a(15)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(15) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(15) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(15) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If



Rem 起站：15圓山
If Combo1.Text = "圓山" Then
    If Combo2.Text = "淡水" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(1)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1) + a(16)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(16) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(16) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(16) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If

mileage_3 = b

End Function


Function mileage_4(b As Single) As Single

Rem 起站：16民權西路
If Combo1.Text = "民權西路" Then
    If Combo2.Text = "淡水" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(1)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1) + a(17)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(17) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(17) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(17) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem 起站：17雙連
If Combo1.Text = "雙連" Then
    If Combo2.Text = "淡水" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(17) + a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(17)
    ElseIf Combo2.Text = "雙連" Then
        b = a(1)
    ElseIf Combo2.Text = "中山" Then
        b = a(1) + a(18)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(18) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(18) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(18) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem 起站：18中山
If Combo1.Text = "中山" Then
    If Combo2.Text = "淡水" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(18) + a(17)
    ElseIf Combo2.Text = "雙連" Then
        b = a(18)
    ElseIf Combo2.Text = "中山" Then
        b = a(1)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1) + a(19)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(19) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(19) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(19) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem 起站：19台北車站
If Combo1.Text = "台北車站" Then
    If Combo2.Text = "淡水" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "雙連" Then
        b = a(19) + a(18)
    ElseIf Combo2.Text = "中山" Then
        b = a(19)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(1)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1) + a(20)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(20) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(20) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(20) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(20) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(20) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If

Rem 起站：20台大醫院
If Combo1.Text = "台大醫院" Then
    If Combo2.Text = "淡水" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "雙連" Then
        b = a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "中山" Then
        b = a(20) + a(19)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(20)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(1)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1) + a(21)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(21) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(21) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(21) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(21) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(21) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


mileage_4 = b

End Function


Function mileage_5(b As Single) As Single

Rem 起站：21中正紀念堂
If Combo1.Text = "中正紀念堂" Then
    If Combo2.Text = "淡水" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "雙連" Then
        b = a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "中山" Then
        b = a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(21) + a(20)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(21)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(1)
    ElseIf Combo2.Text = "東門" Then
        b = a(1) + a(22)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(22) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(22) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(22) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(22) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(22) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If




Rem 起站：22東門
If Combo1.Text = "東門" Then
    If Combo2.Text = "淡水" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "雙連" Then
        b = a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "中山" Then
        b = a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(22) + a(21)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(22)
    ElseIf Combo2.Text = "東門" Then
        b = a(1)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1) + a(23)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(23) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(23) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(23) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(23) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem 起站：23大安森林公園
If Combo1.Text = "大安森林公園" Then
    If Combo2.Text = "淡水" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "雙連" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "中山" Then
        b = a(23) + a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(23) + a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(23) + a(22) + a(21)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(23) + a(22)
    ElseIf Combo2.Text = "東門" Then
        b = a(23)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(1)
    ElseIf Combo2.Text = "大安" Then
        b = a(1) + a(24)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(24) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(24) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(24) + a(25) + a(26) + a(27)
    End If
End If


Rem 起站：24大安
If Combo1.Text = "大安" Then
    If Combo2.Text = "淡水" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "雙連" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "中山" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(24) + a(23) + a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(24) + a(23) + a(22) + a(21)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(24) + a(23) + a(22)
    ElseIf Combo2.Text = "東門" Then
        b = a(24) + a(23)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(24)
    ElseIf Combo2.Text = "大安" Then
        b = a(1)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1) + a(25)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(25) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(25) + a(26) + a(27)
    End If
End If


Rem 起站：25信義安和
If Combo1.Text = "信義安和" Then
    If Combo2.Text = "淡水" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "雙連" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "中山" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(25) + a(24) + a(23) + a(22) + a(21)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(25) + a(24) + a(23) + a(22)
    ElseIf Combo2.Text = "東門" Then
        b = a(25) + a(24) + a(23)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(25) + a(24)
    ElseIf Combo2.Text = "大安" Then
        b = a(25)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(1)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1) + a(26)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(26) + a(27)
    End If
End If



mileage_5 = b

End Function


Function mileage_6(b As Single) As Single
Rem 起站：26台北101/世貿
If Combo1.Text = "台北101/世貿" Then
    If Combo2.Text = "淡水" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + (12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "雙連" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "中山" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22) + a(21)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(26) + a(25) + a(24) + a(23) + a(22)
    ElseIf Combo2.Text = "東門" Then
        b = a(26) + a(25) + a(24) + a(23)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(26) + a(25) + a(24)
    ElseIf Combo2.Text = "大安" Then
        b = a(26) + a(25)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(26)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(1)
    ElseIf Combo2.Text = "象山" Then
        b = a(1) + a(27)
    End If
End If


Rem 起站：27象山
If Combo1.Text = "象山" Then
    If Combo2.Text = "淡水" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3) + a(2)
    ElseIf Combo2.Text = "紅樹林" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4) + a(3)
    ElseIf Combo2.Text = "竹圍" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5) + a(4)
    ElseIf Combo2.Text = "關渡" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6) + a(5)
    ElseIf Combo2.Text = "忠義" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7) + a(6)
    ElseIf Combo2.Text = "復興崗" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8) + a(7)
    ElseIf Combo2.Text = "北投" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9) + a(8)
    ElseIf Combo2.Text = "奇岩" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10) + a(9)
    ElseIf Combo2.Text = "唭哩岸" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11) + a(10)
    ElseIf Combo2.Text = "石牌" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12) + a(11)
    ElseIf Combo2.Text = "明德" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13) + a(12)
    ElseIf Combo2.Text = "芝山" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14) + a(13)
    ElseIf Combo2.Text = "士林" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15) + a(14)
    ElseIf Combo2.Text = "劍潭" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16) + a(15)
    ElseIf Combo2.Text = "圓山" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17) + a(16)
    ElseIf Combo2.Text = "民權西路" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18) + a(17)
    ElseIf Combo2.Text = "雙連" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19) + a(18)
    ElseIf Combo2.Text = "中山" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20) + a(19)
    ElseIf Combo2.Text = "台北車站" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21) + a(20)
    ElseIf Combo2.Text = "台大醫院" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22) + a(21)
    ElseIf Combo2.Text = "中正紀念堂" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23) + a(22)
    ElseIf Combo2.Text = "東門" Then
        b = a(27) + a(26) + a(25) + a(24) + a(23)
    ElseIf Combo2.Text = "大安森林公園" Then
        b = a(27) + a(26) + a(25) + a(24)
    ElseIf Combo2.Text = "大安" Then
        b = a(27) + a(26) + a(25)
    ElseIf Combo2.Text = "信義安和" Then
        b = a(27) + a(26)
    ElseIf Combo2.Text = "台北101/世貿" Then
        b = a(27)
    ElseIf Combo2.Text = "象山" Then
        b = a(1)
    End If
End If



mileage_6 = b

End Function





Private Sub Form_Load()

Rem 淡水信義線各站名稱

Combo1.AddItem "淡水"
Combo1.AddItem "紅樹林"
Combo1.AddItem "竹圍"

Combo1.AddItem "關渡"
Combo1.AddItem "忠義"

Combo1.AddItem "復興崗"
Combo1.AddItem "北投"
Combo1.AddItem "奇岩"

Combo1.AddItem "唭哩岸"
Combo1.AddItem "石牌"
Combo1.AddItem "明德"

Combo1.AddItem "芝山"
Combo1.AddItem "士林"
Combo1.AddItem "劍潭"

Combo1.AddItem "圓山"
Combo1.AddItem "民權西路"
Combo1.AddItem "雙連"

Combo1.AddItem "中山"
Combo1.AddItem "台北車站"
Combo1.AddItem "台大醫院"
Combo1.AddItem "中正紀念堂"

Combo1.AddItem "東門"
Combo1.AddItem "大安森林公園"
Combo1.AddItem "大安"
Combo1.AddItem "信義安和"
Combo1.AddItem "台北101/世貿"

Combo1.AddItem "象山"





Combo2.AddItem "淡水"
Combo2.AddItem "紅樹林"
Combo2.AddItem "竹圍"

Combo2.AddItem "關渡"
Combo2.AddItem "忠義"

Combo2.AddItem "復興崗"
Combo2.AddItem "北投"
Combo2.AddItem "奇岩"

Combo2.AddItem "唭哩岸"
Combo2.AddItem "石牌"
Combo2.AddItem "明德"

Combo2.AddItem "芝山"
Combo2.AddItem "士林"
Combo2.AddItem "劍潭"

Combo2.AddItem "圓山"
Combo2.AddItem "民權西路"
Combo2.AddItem "雙連"

Combo2.AddItem "中山"
Combo2.AddItem "台北車站"
Combo2.AddItem "台大醫院"
Combo2.AddItem "中正紀念堂"

Combo2.AddItem "東門"
Combo2.AddItem "大安森林公園"
Combo2.AddItem "大安"
Combo2.AddItem "信義安和"
Combo2.AddItem "台北101/世貿"

Combo2.AddItem "象山"

Rem 淡水信義線每一站間距(不含新北投)

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
