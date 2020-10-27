VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   3105
   ClientLeft      =   8370
   ClientTop       =   3900
   ClientWidth     =   1740
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   1740
   Begin VB.CommandButton Command5 
      Caption         =   "中级"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "高级"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "自定义"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "初级"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 H = 10
 L = 10
 b = 10
 Call F2
 Unload Me
 Form2.Visible = True
End Sub

Private Sub Command3_Click()
 Unload Me
 Form1.Visible = True
End Sub

Private Sub Command4_Click()
 H = 16
 L = 30
 b = 99
 Call F2
 Unload Me
 Form2.Visible = True
End Sub

Private Sub Command5_Click()
 H = 16
 L = 16
 b = 40
 Call F2
 Unload Me
 Form2.Visible = True
End Sub

