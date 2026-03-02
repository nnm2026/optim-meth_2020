VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "ravnomer_metod"
   ClientHeight    =   12015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   OleObjectBlob   =   "ravnomer_metod.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
x = Val(TextBox1.Text)
dx = Val(TextBox2.Text)
f1 = f(x)
If TextBox3.Text <> "max" And TextBox3.Text <> "min" Then
    MsgBox "Следует ввести max либо min", vbOKOnly + vbInformation, "Внимание !"
End If
If TextBox3.Text = "max" Then f1 = -f1
f2 = f1
Do
     f1 = f2: n = n + 1: x = x + dx
     f2 = f(x)
     If TextBox3.Text = "max" Then f2 = -f2
     TextBox4.Text = TextBox4.Text & "f2= " & f2 & " n= " & n & " dx= " & dx & Chr(13)
Loop Until f1 <= f2
x = x - dx: f2 = f(x) 'откат к предыдущему значению функции
YY = "Ymin = ": If TextBox3.Text = "max" Then YY = "Ymax = "
Label4.Caption = YY + Str(f2) + " при х =  " + Str(x) + " при n =  " + Str(n)
End Sub
Function f(k) As Single
f = k * k / 2 - Sin(k)
End Function
