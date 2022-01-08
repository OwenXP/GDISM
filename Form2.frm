VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "¿É¼Ç"
   ClientHeight    =   5295
   ClientLeft      =   10170
   ClientTop       =   4005
   ClientWidth     =   7005
   LinkTopic       =   "Form2"
   ScaleHeight     =   5295
   ScaleWidth      =   7005
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      DataField       =   "G"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "È®ÀÎ"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "C: , D: µî"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "ÀÌ¸§ ÀÔ·Â"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "~\(ÀÌ¸§).wim"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Ä¸ÃÄµÉ µð·ºÅÍ¸®¸¦ ÀÔ·ÂÇÏ½Ê½Ã¿À."
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "ÀúÀåµÉ ÀÌ¹ÌÁö ÆÄÀÏ °æ·Î¸¦ ÀÔ·ÂÇÏ½Ê½Ã¿À."
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Form3.Label8.Caption = "dism /capture-image /imagefile:" + Text1.Text + " /capturedir:" + Text2.Text + " /name:" + Text3.Text



Form3.Show

End Sub
