VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "NyutonMetod"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10815
   OleObjectBlob   =   "NyutonMetod.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
toch = Val(TextBox1.Text)
Z = Val(InputBox("Начальное значение х", "x = ", "0.1"))
Do
n = n + 1
x = Z
f = x - Cos(x) 'первая производная
D = 1 + Sin(x) 'вторая производная
Z = x - f / D
Loop Until Abs(Z - x) < toch
x = Z
ff = x * x / 2 - Sin(x) 'расчёт значения функции
If D > 0 Then CommandButton1.Caption = "min = " & ff & "    при х = " & x & " и при n = " & n
If D < 0 Then CommandButton1.Caption = "max = " & ff & "    при х = " & x
End Sub
