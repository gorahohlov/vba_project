VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Поиск кодов ТНВЭД по частичному совпадению артикулов, аргументы:"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6885
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SpinButton1_Change()
    TextBox1.Text = SpinButton1.Value
End Sub

Private Sub SpinButton2_Change()
    TextBox2.Text = SpinButton2.Value
End Sub

Private Sub SpinButton3_Change()
    TextBox3.Text = SpinButton3.Value
End Sub

Private Sub SpinButton4_Change()
    TextBox4.Text = SpinButton4.Value
End Sub

Private Sub TextBox1_Change()
    
    On Error Resume Next
    SpinButton1.Value = TextBox1.Text
'    If TextBox1.Text = "" Then
    If Err.Number <> 0 Or TextBox1.Text = "" Then
        MsgBox "Значение не может быть пустым. Введите натуральное число."
        TextBox1.SetFocus
        TextBox1.Text = 9
        Exit Sub
    End If
    On Error GoTo 0
    
'    Do While TextBox1.Text = ""
'        MsgBox "Значение не может быть пустым. Введите натуральное число."
'    Loop
'    SpinButton1.Value = TextBox1.Text
End Sub

Private Sub TextBox2_Change()
    SpinButton2.Value = TextBox2.Text
End Sub

Private Sub TextBox3_Change()
    SpinButton3.Value = TextBox3.Text
End Sub

Private Sub TextBox4_Change()
    SpinButton4.Value = TextBox4.Text
End Sub

Private Sub UserForm_Click()

End Sub
