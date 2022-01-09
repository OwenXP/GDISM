VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GDISM"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows 기본값
   Begin VB.OptionButton Option5 
      Caption         =   "이미지 언마운트"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.OptionButton Option4 
      Caption         =   "이미지 마운트"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   3000
      Width           =   2175
   End
   Begin VB.OptionButton Option3 
      Caption         =   "wim에디션 확인하기"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "이미지 적용"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "이미지 생성"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "관리자 권한으로 실행시켜 주세요."
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()

End Sub

Private Sub Command1_Click()

If Option1 = True Then

Form2.Show


ElseIf Option2 = True Then

Form4.Show

ElseIf Option3 = True Then

Form6.Show

ElseIf Option4 = True Then

Form7.Show

ElseIf Option5 = True Then

Form8.Show

Else

MsgBox "아무 옵션도 선택하지 않았습니다.", vbInformation, "정보"

End If


End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)



End Sub

