VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Золотое сечение"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10845
   OleObjectBlob   =   "Gold.frx":0000
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
t = 1.618033989: t2 = 1 / t: t1 = 1 - t2
Do While (b - a) > toch
    n = n + 1
    x1 = a + t1 * (b - a)
    x2 = a + t2 * (b - a)
    If f(x1) < f(x2) Then
                        b = x2
                      Else: a = x1
    End If
Loop
Label5.Caption = "f= " + Str(f(x1)) + "  x1=" + Str(x1) + "  n=" + Str(n)
End Sub
Function f(x) As Single
 f = x * x / 2 - Sin(x)
End Function

