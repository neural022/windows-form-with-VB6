VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "NBA球隊查詢-兆炫"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   4650
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "重新查詢"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查詢"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Logo"
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      Begin VB.OptionButton Option6 
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   2280
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   2280
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   1275
         Left            =   120
         Picture         =   "Form1.frx":0000
         Top             =   2760
         Width           =   4110
      End
      Begin VB.Image Image1 
         Height          =   1215
         Left            =   240
         Picture         =   "Form1.frx":3E44
         Top             =   720
         Width           =   3810
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   5640
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入球隊名稱:(ex：xxxx隊)"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim NBA, a As String
NBA = Text1.Text

Select Case NBA

  Case Is = "洛杉磯湖人隊"
        Option1.Value = True
  Case Is = "芝加哥公牛隊"
        Option2.Value = True
  Case Is = "邁阿密熱火隊"
        Option3.Value = True
  Case Is = "洛杉磯快艇隊"
        Option4.Value = True
  Case Is = "克里夫蘭騎士隊"
        Option5.Value = True
  Case Is = "金州勇士隊"
        Option6.Value = True
        
End Select

If Option1.Value = True Then Label3.Caption = "此Logo為洛杉磯湖人隊"
If Option2.Value = True Then Label3.Caption = "此Logo為芝加哥公牛隊"
If Option3.Value = True Then Label3.Caption = "此Logo為邁阿密熱火隊"
If Option4.Value = True Then Label3.Caption = "此Logo為洛杉磯快艇隊"
If Option5.Value = True Then Label3.Caption = "此Logo為克里夫蘭騎士隊"
If Option6.Value = True Then Label3.Caption = "此Logo為金州勇士隊"

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Label3.Caption = ""
End Sub

Private Sub Option1_Click()
Label3.Caption = ""

If Option1.Value = True Then Label3.Caption = "此Logo為洛杉磯湖人隊"

End Sub

Private Sub Option2_Click()
Label3.Caption = ""

If Option2.Value = True Then Label3.Caption = "此Logo為芝加哥公牛隊"
End Sub

Private Sub Option3_Click()
Label3.Caption = ""

If Option3.Value = True Then Label3.Caption = "此Logo為邁阿密熱火隊"
End Sub

Private Sub Option4_Click()
Label3.Caption = ""

If Option4.Value = True Then Label3.Caption = "此Logo為洛杉磯快艇隊"
End Sub

Private Sub Option5_Click()
Label3.Caption = ""

If Option5.Value = True Then Label3.Caption = "此Logo為克里夫蘭騎士隊"
End Sub

Private Sub Option6_Click()
Label3.Caption = ""

If Option6.Value = True Then Label3.Caption = "此Logo為金州勇士隊"
End Sub
