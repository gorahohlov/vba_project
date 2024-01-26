VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Поиск кодов ТНВЭД по частичному совпадению артикулов, аргументы:"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   OleObjectBlob   =   "UserForm1_2024-01-22.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub RefEdit1_Change()
    On Error Resume Next
    processing_row_num = Range(Me.RefEdit1.Text).Row
    article_col_num = Range(Me.RefEdit1.Text).Column
    If Err.Number <> 0 Then
        MsgBox "Невалидный или пустой диапазон. Введите ссылку на диапазон."
        RefEdit1.SetFocus
        RefEdit1.SelStart = 0
        RefEdit1.SelLength = Len(RefEdit1.Text)
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Sub RefEdit2_Change()
    On Error Resume Next
    Set vlookup_table_rng = Range(Me.RefEdit2.Text)
    If Err.Number <> 0 Then
        MsgBox "Невалидный или пустой диапазон. Введите ссылку на диапазон."
        RefEdit2.SetFocus
        RefEdit2.SelStart = 0
        RefEdit2.SelLength = Len(RefEdit2.Text)
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()
    Me.RefEdit1 = Cells( _
                        processing_row_num, _
                        article_col_num _
                       ).Address(external:=True)
    Me.RefEdit2 = vlookup_table_rng.Address(external:=True)
End Sub

'Public Sub UserForm_Initialize( _
'                                proc_row_num, _
'                                arti_col_num, _
'                                tabl_rng _
'                               )
'
'    Me.RefEdit1 = Cells( _
'                        proc_row_num, _
'                        arti_col_num _
'                       ).Address(external:=True)
'    Me.RefEdit2 = tabl_rng.Address(external:=True)
'
'End Sub

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
    If SpinButton4.Value > SpinButton3.Value Then
        MsgBox "Значение не может быть больше " & TextBox3.Text
        SpinButton4.SetFocus
        SpinButton4.Value = SpinButton3.Value
    End If
    TextBox4.Text = SpinButton4.Value
End Sub

Private Sub TextBox1_Change()
    On Error Resume Next
    SpinButton1.Value = Val(TextBox1.Text)
    TextBox1.Text = SpinButton1.Value
    If Err.Number <> 0 Or TextBox1.Text = "" Then
        MsgBox "Значение не может быть пустым." & vbCrLf _
        & "Введите натуральное число не более 16384."
        TextBox1.SetFocus
        TextBox1.Text = 9
        TextBox1.SelStart = 0
        TextBox1.SelLength = 1
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Sub TextBox2_Change()
    On Error Resume Next
    SpinButton2.Value = Val(TextBox2.Text)
    TextBox2.Text = SpinButton2.Value
    If Err.Number <> 0 Then
        MsgBox "Значение не может быть пустым." & vbCrLf _
                    & "Введите натуральное число не более 16384."
        TextBox2.SetFocus
        TextBox2.Text = 1
        TextBox2.SelStart = 0
        TextBox2.SelLength = 1
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Sub TextBox3_Change()
    On Error Resume Next
    SpinButton3.Value = Val(TextBox3.Text)
    TextBox3.Text = SpinButton3.Value
    If Err.Number <> 0 Then
        MsgBox "Значение не может быть пустым." & vbCrLf & _
               "Значение может быть только числом." & vbCrLf & _
                vbCrLf & _
               "Введите число максимального количества начальных " & _
                vbCrLf & _
               "символов артикула, по которым будут искаться совпадения."
        TextBox3.SetFocus
        TextBox3.Text = 12
        TextBox3.SelStart = 0
        TextBox3.SelLength = 2
        Exit Sub
    End If
'    If SpinButton3.Value < SpinButton4.Value Then
'        SpinButton4.Value = Val(TextBox3.Text)
'        TextBox4.Text = TextBox3.Text
'    End If
    On Error GoTo 0
End Sub

Private Sub TextBox4_Change()
    On Error Resume Next
    SpinButton4.Value = Val(TextBox4.Text)
    TextBox4.Text = SpinButton4.Value
    If Err.Number <> 0 Or Val(TextBox4.Text) > _
                                Val(TextBox3.Text) Then
        MsgBox "Значение не может быть пустым." & vbCrLf & _
               "Значение может быть только числом." & vbCrLf & _
               "Значение не может быть больше " & TextBox3.Text & _
                vbCrLf & vbCrLf & _
               "Введите число минимального количества начальных" & _
                vbCrLf & _
               "символов артикула, по которым будут искаться совпадения."
        TextBox4.SetFocus
        TextBox4.Text = 9
        TextBox4.SelStart = 0
        TextBox4.SelLength = 2
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Sub CommandButton1_Click()

    On Error Resume Next
    
    vlookup_arg3 = Val(Me.TextBox1.Text)
    vlookup_arg4 = Val(Me.TextBox2.Text)
    upper_interval = Val(Me.TextBox3.Text)
    lower_interval = Val(Me.TextBox4.Text)
    processing_row_num = Range(Me.RefEdit1.Text).Row
    article_col_num = Range(Me.RefEdit1.Text).Column
    Set vlookup_table_rng = Range(Me.RefEdit2.Text) '.Address(external:=True)
    
    If Err.Number <> 0 Then
        MsgBox "Невалидный или пустой диапазон. Введите ссылку на диапазон."
        
        If RefEdit1.Text = "" Then
            RefEdit1.SetFocus
            RefEdit1.SelStart = 0
            RefEdit1.SelLength = Len(RefEdit1.Text)
        ElseIf RefEdit2.Text = "" Then
            RefEdit2.SetFocus
            RefEdit2.SelStart = 0
            RefEdit2.SelLength = Len(RefEdit2.Text)
        Else
            RefEdit2.SetFocus
            RefEdit2.SelStart = 0
            RefEdit2.SelLength = Len(RefEdit2.Text)
        End If
        
        Exit Sub
    End If
    
    If Val(TextBox3.Text) < Val(TextBox4.Text) Then
        MsgBox "Значение не может быть меньше " & TextBox4.Text & _
                vbCrLf & vbCrLf & _
               "Введите число максимального количества начальных" & _
                vbCrLf & _
               "символов артикула, по которым будут искаться совпадения."
        TextBox3.Text = TextBox4.Text
        TextBox3.SetFocus
        TextBox3.SelStart = 0
        TextBox3.SelLength = Len(TextBox3.Text)
        Exit Sub
    End If
    
    cancel_flag = False
    Unload Me
    On Error GoTo 0
End Sub

Private Sub CommandButton2_Click()
    cancel_flag = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        cancel_flag = True
        Unload Me
        Cancel = False
    End If
End Sub
