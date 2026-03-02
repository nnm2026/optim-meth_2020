VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "‘ибоначчи"
   ClientHeight    =   10815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10845
   OleObjectBlob   =   "fibonachi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
a = Val(TextBox1.Text): b = Val(TextBox2.Text)
toch = Val(TextBox3.Text): n = Val(TextBox4.Text)
x1 = 1: x2 = 1
For i = 2 To n
  x3 = x1 + x2
  x1 = x2: x2 = x3
  If i = n - 1 Then fn_1 = x3 'n-1 -е число ‘пбоначчи
Next
fn = x3 'n-е число ‘пбоначчи
x2 = a + ((b - a) * fn_1 + toch * (-1) ^ n) / fn
For i = 1 To n
  x4 = a - x2 + b
  If (x4 < x2) And (f(x4) > f(x2)) Then a = x4
  If (x4 < x2) And (f(x4) < f(x2)) Then
                                    b = x2: x2 = x4
    End If
  If (x4 > x2) And (f(x4) < f(x2)) Then
                                    a = x2: x2 = x4
    End If
  If (x4 > x2) And (f(x4) > f(x2)) Then b = x4
Next
Label5.Caption = "f= " + Str(f(x4)) + "   x4=" + Str(x4)
End Sub
Function f(x) As Single
 f = x * x / 2 - Sin(x)
End Function

