VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ӣ��ɼ�����������Լ�棩"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   11955
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command7 
      Caption         =   "�������棨��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      TabIndex        =   16
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�������ûɶ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9960
      TabIndex        =   15
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "miku��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   12
      Top             =   2880
      Width           =   6855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��һ�°�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   11
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��һ�°�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   1
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "�л�ģʽ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "���ĵ÷�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   10
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "ǰ����Ŀ�۵ķ���������1.25��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "���Ӣ��ɼ�������1.25�󣬱�����λС����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "���Ӣ��ɼ������ճɼ���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Ӣ��ɼ�����������Լ�棩"
      BeginProperty Font 
         Name            =   "��������"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Text2.Text = Format(Val(Text1.Text) / 1.25, "#.##")
End Sub

Private Sub Command3_Click()
Text4.Text = Val(Text2.Text) - 80 + Val(Text3.Text)
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text3.Text = ""
Text2.Text = ""
Text4.Text = ""
End Sub

Private Sub Command5_Click()
MsgBox ("��û����(����գ�)�g")
End Sub

Private Sub Command6_Click()
MsgBox ("�������ûɶ��")
End Sub

Private Sub Command7_Click()
Form1.Hide
Form3.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
