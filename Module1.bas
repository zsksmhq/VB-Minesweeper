Attribute VB_Name = "Module1"
Option Explicit
Public bn As Integer
Public a() As Integer
Public Block() As CAREA
Public H, L, b As Integer
Public S, Sa, Sb As Integer
Public asdf As Integer
Public Mouse As Integer
Public Type CAREA
  Bomb As Boolean
  Name As Integer
  number As Integer
  kai As Boolean
  Sign As Boolean
  BombNum As Integer
End Type
 Public Sub Bili()
 Dim sum As Integer
 sum = asdf + Form2.Label2.Caption
 If sum < 100 Then Form2.Label2.FontSize = 16
 If sum < 10 Then
  Form2.Label2.Caption = " " & Str(sum)
 Else
  Form2.Label2.Caption = Str(sum)
 End If
End Sub

Public Sub F2()
 Dim n, nn, k As Integer
 k = 0
 ReDim Preserve Block(0 To L + 1, 0 To H + 1)
 Form2.Height = H * 375 + 870
 Form2.Width = L * 375 + 45
 Form2.Top = (11520 - Form2.Height) / 2
 Form2.Left = (20175 - Form2.Width) / 2
 For n = 0 To H + 1
  For nn = 0 To L + 1
   If n <> 0 And nn <> 0 And n <> H + 1 And nn <> L + 1 Then
    k = k + 1
    Block(nn, n).Name = k
   Else
    Block(nn, n).Name = 0
   End If
  Next
 Next
 k = 0
 For n = 1 To H
  For nn = 1 To L
   k = k + 1
   Load Form2.Label(k)
        Form2.Label(k).Left = (nn - 1) * 375
        Form2.Label(k).Top = (n - 1) * 375 + 495
        Form2.Label(k).Visible = True
  Next
 Next
 Form2.Label2.Left = Form2.Width - 860
 Form2.Label2.Caption = b
 Call Bili
 Call CalBomb
End Sub

Public Sub Search()
 Dim n, nn As Integer
  For n = 1 To H
   For nn = 1 To L
    If Block(nn, n).Name = S Then
     Sa = nn
     Sb = n
     Exit Sub
    End If
   Next
  Next
End Sub

Public Sub CalBomb()
Dim i As Integer, k, n, nn As Integer
ReDim Preserve a(0 To b)
For i = 0 To b - 1
dd:
DoEvents
Randomize
a(i) = Int(Rnd * L * H)
If i > 0 Then
  For k = 0 To i - 1
     If a(i) = a(k) Then GoTo dd
  Next k
End If
Next i
For i = 0 To b - 1
S = a(i)
Call Search
 Block(Sa, Sb).Bomb = True
 For n = -1 To 1
  For nn = -1 To 1
   Block(Sa + n, Sb + nn).BombNum = Block(Sa + n, Sb + nn).BombNum + 1
  Next
 Next
Next

End Sub
