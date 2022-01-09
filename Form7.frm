VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7815
   LinkTopic       =   "Form7"
   ScaleHeight     =   5790
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows 기본값
   Begin VB.OptionButton Option2 
      Caption         =   "읽기/쓰기"
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
      Left            =   3720
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "읽기전용"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
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
      Left            =   5880
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   3840
      Width           =   3495
   End
   Begin VB.TextBox Text2 
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
      TabIndex        =   3
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox Text1 
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
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "마운트할 에디션 번호를 입력하십시오."
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "마운트할 디랙터리를 입력하십시오."
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "마운트될 이미지 파일 경로를 입력하십시오."
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1 = True Then
Shell ("cmd /k dism /mount-image /imagefile:" + Text1.Text + " /index:" + Text3.Text + " /mountdir:" + Text2.Text + " /readonly"), vbNormalNoFocus
ElseIf Option2 = True Then
Shell ("cmd /k dism /mount-image /imagefile:" + Text1.Text + " /index:" + Text3.Text + " /mountdir:" + Text2.Text), vbNormalNoFocus
Else
MsgBox "읽기전용, 읽기/쓰기중 하나를 선택하십시오.", vbInformation, "정보"


End If




End Sub

