VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Корень f(x)=0 методом половинного деления"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6150
   OleObjectBlob   =   "polov_delen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
a = Val(TextBox1.Text): b = Val(TextBox2.Text)
toch = Val(TextBox3.Text)
yc = 1
Do While Abs(yc) > toch
     k = k + 1
     xc = (a + b) / 2
     yc = f(xc)
     If f(a) * yc > 0 Then a = xc Else b = xc
Loop
Label4.Caption = Label4.Caption + "Y = " + Str(yc) + "  x = " + Str(xc) + "  k = " + Str(k)
End Sub
Function f(n) As Single
f = n * n / 2 - Sin(n)
End Function
