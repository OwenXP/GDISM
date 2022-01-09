VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "È®ÀÎÃ¢"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4695
   LinkTopic       =   "Form3"
   ScaleHeight     =   3675
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3255
      Begin VB.Label Label8 
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "¸í·É¾î"
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
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "È®ÀÎ"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3240
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

a = Label8.Caption

Shell ("cmd /k " + a), vbNormalNoFocus


End Sub

Private Sub Label5_Click()

End Sub

