VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ӣ��ɼ����������������棩"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14790
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   16750.37
   ScaleMode       =   0  'User
   ScaleWidth      =   44078.09
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command6 
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
      Left            =   12240
      TabIndex        =   17
      Top             =   1440
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   4695
      Left            =   2160
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   4395
      TabIndex        =   16
      Top             =   3840
      Width           =   4455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�����棨ȷ�ţ�"
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
      Left            =   9840
      TabIndex        =   9
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�ؿ���"
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
      Left            =   720
      TabIndex        =   8
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
      Left            =   8160
      TabIndex        =   7
      Top             =   1920
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
      Left            =   6720
      TabIndex        =   6
      Top             =   1800
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
      Left            =   5040
      TabIndex        =   5
      Top             =   1800
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
      Left            =   3840
      TabIndex        =   4
      Top             =   1920
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
      Left            =   2280
      TabIndex        =   3
      Top             =   1800
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
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ˣ��ˣ��ˣ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7080
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Left            =   9480
      Picture         =   "Form3.frx":812B
      ScaleHeight     =   5595
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   2760
      Width           =   4815
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
      Left            =   6720
      TabIndex        =   15
      Top             =   1200
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
      Left            =   5040
      TabIndex        =   14
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "�������ɼ�������1.25�󣬱�����λС����"
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
      Left            =   2280
      TabIndex        =   13
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "�������ɼ������ճɼ���"
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
      Left            =   480
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "����ɼ����������������棩Yinggelishi Score jisuanqi"
      BeginProperty Font 
         Name            =   "��������"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   11
      Top             =   240
      Width           =   4575
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
      Left            =   10200
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
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
Form3.Hide
Form1.Show
End Sub

Private Sub Command6_Click()
MsgBox ("��û����(����գ�)�g")
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

