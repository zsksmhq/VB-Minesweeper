VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2700
   ClientLeft      =   9645
   ClientTop       =   5490
   ClientWidth     =   4320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4320
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   1200
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  0"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体-方正超大字符集"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Win As Boolean
Public timer As Boolean
Private Mice As Integer

Private Sub Label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Mouse = 0 Then
 Mouse = Button
 If timer = False Then
   Form2.Label1.Caption = "  1"
   Timer1.Enabled = True
   timer = True
  End If
  If Win = False Then
   If Button = 2 And Label(Index).Caption = "" And Label(Index).Appearance = 0 Then
     asdf = -1
     Call Bili
     Label(Index).Caption = "●"
    ElseIf Button = 2 And Label(Index).Caption = "●" Then
     asdf = 1
     Call Bili
     Label(Index).Caption = ""
    ElseIf Button = 1 And Label(Index).Caption = "" Then
     S = Index
     Call Search
     If Block(Sa, Sb).Bomb = False Then
       Call OpenB
      Else
       Form2.Label(Block(Sa, Sb).Name).BackColor = &HFF&
       MsgBox "Bomb~~~(￣_,￣ ) "
       Win = True
       Timer1.Enabled = False
      End If
   End If
  End If
Else
 If Mouse <> Button And Label(Index).BackColor = &HFFC0C0 Then
  Mice = Index
  Call OpenA
 End If
End If
End Sub

Public Sub OpenB()
Dim i, ii, n, nn As Integer
Dim saa, sbb As Integer
dd:
 DoEvents
 If Block(Sa, Sb).kai = False Then
  bn = bn + 1
  Block(Sa, Sb).kai = True
  Form2.Label(Block(Sa, Sb).Name).Caption = Block(Sa, Sb).BombNum
  Form2.Label(Block(Sa, Sb).Name).Appearance = 1
    Form2.Label(Block(Sa, Sb).Name).BackColor = &HFFC0C0
  Block(Sa, Sb).Sign = 2
 End If
 If Block(Sa, Sb).BombNum = 0 Then
  Label(Block(Sa, Sb).Name).Caption = ""
  For i = -1 To 1
   For ii = -1 To 1
    Block(Sa + i, Sb + ii).Sign = True
   Next ii
  Next i
  End If
 For n = 1 To H
  For nn = 1 To L
   If Block(nn, n).Sign = True And Block(nn, n).kai = False Then
    Sa = nn
    Sb = n
    GoTo dd
   End If
  Next
 Next
 If bn = H * L - b Then
  MsgBox "*★,°*:.☆\(￣▽￣)/$:*.°★* "
  Win = True
  Timer1.Enabled = False
 End If
End Sub

Private Sub Label_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Mouse = 0
End Sub


Private Sub Timer1_Timer()
 If Val(Label1.Caption) > 98 Then Label1.FontSize = 12
 If Val(Label1.Caption) < 9 Then
  Label1.Caption = " " & Str(Val(Label1.Caption) + 1)
 Else
  Label1.Caption = Str(Val(Label1.Caption) + 1)
 End If
End Sub

Public Sub OpenA()
 Dim n, nn, Jishu, Jishu1 As Integer
 S = Mice
 Jishu = Val(Label(Mice).Caption)
 Call Search
 For n = -1 To 1
  For nn = -1 To 1
   If Label(Block(Sa + n, Sb + nn).Name).Caption = "●" Then Jishu1 = Jishu1 + 1
  Next
 Next
 If Jishu = Jishu1 Then
  For n = -1 To 1
   For nn = -1 To 1
   If Label(Block(Sa + n, Sb + nn).Name).Caption <> "●" Then Block(Sa + n, Sb + nn).Sign = True
   Next
  Next
  Call OpenB
 End If
End Sub
