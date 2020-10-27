VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " →_→"
   ClientHeight    =   2040
   ClientLeft      =   8055
   ClientTop       =   4530
   ClientWidth     =   2625
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   2625
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "结束"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "宽（格）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "雷数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "长（格）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private st(0 To 2) As String
Private a(0 To 2) As Integer
Public f As Boolean

Private Sub Command1_Click()
 L = Val(Form1.Text(0).Text)
 H = Val(Form1.Text(1).Text)
 b = Val(Form1.Text(2).Text)
 Call F2
 Unload Me
 Form2.Visible = True
End Sub

Private Sub Command2_Click()
 End
End Sub

Private Sub Form_Load()
 f = True

End Sub

Private Sub Text_Change(Index As Integer) '此过程用以限制textbox的输入
 If f = False Then
  f = True
  Exit Sub
 End If
 Dim i, c As Integer
 Dim S As String
 For i = 0 To 2
  S = Text(i).Text
  If Text(i).Text <> "" Then
   For c = 1 To Len(S)
    If Asc(Mid(S, c, 1)) > 57 Or Asc(Mid(S, c, 1)) < 48 Then
     f = False
     a(i) = Text(i).SelStart - 1                                  '将光标复位
     Text(i).Text = st(i)
     Text(i).SelStart = a(i)
    End If
   Next
   End If
  st(i) = Text(i).Text
 Next
End Sub

