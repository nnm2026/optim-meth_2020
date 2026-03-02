VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Sekush"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7020
   OleObjectBlob   =   "sekush.frx":0000
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
x = a - f(a) * (b - a) / (f(b) - f(a))
TextBox4.Text = TextBox4.Text + "f(a) = " + Str(f(a)) + "  f(b) = " + Str(f(b)) + Chr(13)
TextBox4.Text = TextBox4.Text + "x = " + Str(x) + "  f(x) = " + Str(f(x)) + Chr(13)
Do While Abs(f(x)) > toch
     k = k + 1
     x = a - f(a) * (b - a) / (f(b) - f(a))
     TextBox4.Text = TextBox4.Text + "f(a)=f(" + Str(a) + ")=" + Str(f(a)) + "  f(b) = f(" + Str(b) + ")=" + Str(f(b)) + Chr(13)
     TextBox4.Text = TextBox4.Text + "x = " + Str(x) + "  f(x) = " + Str(f(x)) + Chr(13)
     b = a
     a = x
Loop
Label4.Caption = Label4.Caption + "Y = " + Str(f(x)) + "  x = " + Str(x) + "  k = " + Str(k)
End Sub
Function f(n) As Single
f = n - Cos(n)
End Function
