VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GDISM"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.OptionButton Option3 
      Caption         =   "wim����� Ȯ���ϱ�"
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
      Left            =   600
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "�̹��� ����"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "�̹��� ����"
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
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ȯ��"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "������ �������� ������� �ּ���."
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

Else

MsgBox "�ƹ� �ɼǵ� �������� �ʾҽ��ϴ�.", vbInformation, "����"

End If


End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)



End Sub

