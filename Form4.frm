VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7620
   LinkTopic       =   "Form4"
   ScaleHeight     =   5445
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "Ȯ��"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "1, 2, 3 �̷� ������ �Է��Ͻʽÿ�. ������� ����� Ȯ�ο��� Ȯ���Ͻʽÿ�."
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3840
      TabIndex        =   9
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "C: , D: ��"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   8
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "~\install.wim(.esd)"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "���� ������� ��ġ�Ұ��� �����Ͻʽÿ�."
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "���ٰ� ��ġ�Ұ��� �����Ͻʽÿ�."
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "��ġ �̹��� ������ ���"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   18
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
      Width           =   3255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Form5.Label1.Caption = "dism /apply-image /imagefile:" + Text1.Text + " /applydir:" + Text2.Text + " /index:" + Text3.Text

Form5.Show

End Sub
