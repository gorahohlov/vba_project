VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Поиск кодов ТНВЭД по частичному совпадению артикулов, аргументы:"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7965
   OleObjectBlob   =   "UserForm1_2024-02-12.frx":0000
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

    If sel_col_num = 1 Then
        Me.Label8.Enabled = False
        Me.Label9.Enabled = False
        Me.TextBox5.Enabled = False
        Me.TextBox6.Enabled = False
        Me.SpinButton5.Enabled = False
        Me.SpinButton6.Enabled = False
    ElseIf sel_col_num = 2 Then
        Me.Label9.Enabled = False
        Me.TextBox6.Enabled = False
        Me.SpinButton6.Enabled = False
    End If

End Sub

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

Private Sub SpinButton5_Change()
    TextBox5.Text = SpinButton5.Value
End Sub

Private Sub SpinButton6_Change()
    TextBox6.Text = SpinButton6.Value
End Sub

Private Sub TextBox1_Change()
    On Error Resume Next
    SpinButton1.Value = Val(TextBox1.Text)
    TextBox1.Text = SpinButton1.Value
    If Err.Number <> 0 Or _
       TextBox1.Text = "" Or _
       Val(TextBox1.Text) > Range(Me.RefEdit2.Text).Columns.Count Then
        MsgBox "Значение не может быть пустым или превышать ширину " & _
               "области Источника данных." & vbCrLf & vbCrLf & _
               "Введите натуральное число не более " & _
                Range(Me.RefEdit2.Text).Columns.Count & "."
        TextBox1.SetFocus
        TextBox1.Text = Range(Me.RefEdit2.Text).Columns.Count
        TextBox1.SelStart = 0
        TextBox1.SelLength = Len(TextBox1.Text)
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Sub TextBox5_Change()
    On Error Resume Next
    SpinButton5.Value = Val(TextBox5.Text)
    TextBox5.Text = SpinButton5.Value
    If Err.Number <> 0 Or _
       TextBox5.Text = "" Or _
       Val(TextBox5.Text) > Range(Me.RefEdit2.Text).Columns.Count Then
        MsgBox "Значение не может быть пустым или превышать ширину " & _
               "области Источника данных." & vbCrLf & vbCrLf & _
               "Введите натуральное число не более " & _
                Range(Me.RefEdit2.Text).Columns.Count & "."
        TextBox5.SetFocus
        TextBox5.Text = Range(Me.RefEdit2.Text).Columns.Count
        TextBox5.SelStart = 0
        TextBox5.SelLength = Len(TextBox5.Text)
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Sub TextBox6_Change()
    On Error Resume Next
    SpinButton6.Value = Val(TextBox6.Text)
    TextBox6.Text = SpinButton6.Value
    If Err.Number <> 0 Or _
       TextBox6.Text = "" Or _
       Val(TextBox6.Text) > Range(Me.RefEdit2.Text).Columns.Count Then
        MsgBox "Значение не может быть пустым или превышать ширину " & _
               "области Источника данных." & vbCrLf & vbCrLf & _
               "Введите натуральное число не более " & _
                Range(Me.RefEdit2.Text).Columns.Count & "."
        TextBox6.SetFocus
        TextBox6.Text = Range(Me.RefEdit2.Text).Columns.Count
        TextBox6.SelStart = 0
        TextBox6.SelLength = Len(TextBox6.Text)
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Sub TextBox2_Change()
    On Error Resume Next
    SpinButton2.Value = Val(TextBox2.Text)
    TextBox2.Text = SpinButton2.Value
    If Err.Number <> 0 Or _
       TextBox2.Text = "" Or _
       Val(TextBox2.Text) > Range(Me.RefEdit2.Text).Columns.Count Then
        MsgBox "Значение не может быть пустым или превышать ширину " & _
               "области Источника данных." & vbCrLf & vbCrLf & _
               "Введите натуральное число не более " & _
                Range(Me.RefEdit2.Text).Columns.Count & "."
        TextBox2.SetFocus
        TextBox2.Text = Range(Me.RefEdit2.Text).Columns.Count
        TextBox2.SelStart = 0
        TextBox2.SelLength = Len(TextBox2.Text)
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
    vlookup_arg4_ver2 = Val(Me.TextBox5.Text)
    vlookup_arg4_ver3 = Val(Me.TextBox6.Text)
    upper_interval = Val(Me.TextBox3.Text)
    lower_interval = Val(Me.TextBox4.Text)
    processing_row_num = Range(Me.RefEdit1.Text).Row
    article_col_num = Range(Me.RefEdit1.Text).Column
    Set vlookup_table_rng = Range(Me.RefEdit2.Text)
    
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
    
    table_area_width = vlookup_table_rng.Columns.Count
    If vlookup_arg3 > table_area_width Then
        MsgBox "Значение не может быть больше ширины в области:" & _
                vbCrLf & _
               "<Источник данных (диапазон): [" & _
                Range(RefEdit2.Text).Address(external:=False) & "]>." & _
                vbCrLf & vbCrLf & _
               "Введите натуральное число от 1 до " & table_area_width & "."
        TextBox1.Text = table_area_width
        TextBox1.SetFocus
        TextBox1.SelStart = 0
        TextBox1.SelLength = Len(TextBox1.Text)
        Exit Sub
    ElseIf vlookup_arg4 > table_area_width Then
        MsgBox "Значение не может быть больше ширины в области:" & _
                vbCrLf & _
               "<Источник данных (диапазон): [" & _
                Range(RefEdit2.Text).Address(external:=False) & "]>." & _
                vbCrLf & vbCrLf & _
               "Введите натуральное число от 1 до " & table_area_width & "."
        TextBox2.Text = table_area_width
        TextBox2.SetFocus
        TextBox2.SelStart = 0
        TextBox2.SelLength = Len(TextBox2.Text)
        Exit Sub
    ElseIf vlookup_arg4_ver2 > table_area_width And sel_col_num > 1 Then
        MsgBox "Значение не может быть больше ширины в области:" & _
                vbCrLf & _
               "<Источник данных (диапазон): [" & _
                Range(RefEdit2.Text).Address(external:=False) & "]>." & _
                vbCrLf & vbCrLf & _
               "Введите натуральное число от 1 до " & table_area_width & "."
        TextBox5.Text = table_area_width
        TextBox5.SetFocus
        TextBox5.SelStart = 0
        TextBox5.SelLength = Len(TextBox5.Text)
        Exit Sub
    ElseIf vlookup_arg4_ver3 > table_area_width And sel_col_num > 2 Then
        MsgBox "Значение не может быть больше ширины в области:" & _
                vbCrLf & _
               "<Источник данных (диапазон): [" & _
                Range(RefEdit2.Text).Address(external:=False) & "]>." & _
                vbCrLf & vbCrLf & _
               "Введите натуральное число от 1 до " & table_area_width & "."
        TextBox6.Text = table_area_width
        TextBox6.SetFocus
        TextBox6.SelStart = 0
        TextBox6.SelLength = Len(TextBox6.Text)
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
