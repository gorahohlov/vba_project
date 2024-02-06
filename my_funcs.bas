Attribute VB_Name = "Module1"

Function VLookUp2( _
                  search_value As Variant, _
                  table_rng As Range, _
                  search_col_num As Integer, _
                  result_col_num As Integer, _
                  match_num As Integer _
                 )

'search_value: ������� ��������;
'table_rng: �������, �������� �����, � ������� ������ ���������� �
'                   ����������;
'search_col_num: ����� �������, � ��������� ����� [table_rng],
'                   � ������� ������ ����������;
'result_col_num: ����� �������,  � ��������� ����� [table_rng],
'                   �� ������� ����������� ������� ������;
'match_num: ����� ���������� �������� (���� ���������� �������������);

    Dim FLG As Boolean
    Dim i As Integer
    Dim iCount As Integer

    FLG = False

    For i = 1 To table_rng.Rows.Count

        If table_rng.Cells(i, search_col_num) = search_value Then
            iCount = iCount + 1
        End If

        If iCount = match_num Then
            VLookUp2 = table_rng.Cells(i, result_col_num)
            FLG = True
            Exit For
        End If

    Next i

    If FLG = False Then
        VLookUp2 = CVErr(xlErrNA)
    End If

End Function

Function VLookUp3( _
                  search_value As Variant, _
                  table_rng As Range, _
                  search_col_num As Integer, _
                  result_col_num As Integer, _
                  match_num As Integer _
                 )

'search_value: ������� ��������;
'table_rng: �������, �������� �����, � ������� ������ ���������� �
'                   ����������;
'search_col_num: ����� �������, � ��������� ����� [table_rng],
'                   � ������� ������ ����������;
'result_col_num: ����� �������, � ��������� ����� [table_rng],
'                   �� ������� ����������� ������� ������;
'match_num: ����� ���������� �������� (���� ���������� �������������);

    Dim FLG As Boolean
    Dim i As Integer
    Dim iCount As Integer

    FLG = False

    For i = 1 To table_rng.Rows.Count

        If search_value Like table_rng.Cells(i, search_col_num) Then
            iCount = iCount + 1
        End If

        If iCount = match_num Then
            VLookUp3 = table_rng.Cells(i, result_col_num)
            FLG = True
            Exit For
        End If

    Next i

    If FLG = False Then
        VLookUp3 = CVErr(xlErrNA)
    End If

End Function

Function VLookUp4( _
                  search_value As Variant, _
                  table_rng As Range, _
                  search_col_num As Integer, _
                  result_col_num As Integer, _
                  Optional symbols_num = 0 _
                 )

'search_value: ������� ��������;
'table_rng: �������, �������� �����, � ������� ������ ���������� �
'                   ����������;
'search_col_num: ����� �������, � ��������� ����� [table_rng],
'                   � ������� ������ ����������;
'result_col_num: ����� �������, � ��������� ����� [table_rng],
'                   �� ������� ����������� ������� ������;
'symbols_num: ���������� ������ ����� �������� ��������� ��������
'                   (��������), �� ������� ����� �������� ����������.

    Dim FLG As Boolean
    Dim i As Integer

    FLG = False

    For i = 1 To table_rng.Rows.Count

        If symbols_num = 0 Then
            If table_rng.Cells(i, search_col_num) = search_value Then
                VLookUp4 = table_rng.Cells(i, result_col_num)
                FLG = True
                Exit For
            End If
        Else
            If Left(table_rng.Cells(i, search_col_num), symbols_num) = _
                                    Left(search_value, symbols_num) Then
                VLookUp4 = table_rng.Cells(i, result_col_num)
                FLG = True
                Exit For
            End If
        End If

    Next i

    If FLG = False Then
        VLookUp4 = CVErr(xlErrNA)
    End If

End Function

Function custom_toll( _
                     custom_sum As Variant, _
                     Optional currency_rate As Single = 1, _
                     Optional msg_flag As Boolean = False _
                    ) As Variant

'���������� �������.
'��� ������� ������������ �����  ���������� ������ � ����������� ��
'����� ���������� ��������� ������������ �� ������ (������ ��������),
'������ �������� - �������������� - ��� ����, �� ������� ����������!
'������ �������� �������, ����� �������� ����� ���������� ���������
'(� ������) ��� ������� ������. ���� ��������������� ��� ������ ��������
'(������������) ��� ���������� ��������� � ������, ����� ������ ��������
'(����) ����� �� ��������� ��� ��������� 1 (�������) - ����� ��������
'����� �� ���������, ���� ������ �������� ��������.

'��� ������ �������.
'���� ������� ��������� �� ����� � ������ (��������, ���� ���
'����� - �������) - ����� ������ ���������� ������� ����, �� �������
'����� ����������! ������ ��������, ����� �������� ����� � ������.
'����� ���������� ���������� ��������� � ������ "�����������" ��
'����������� "IF ElseIF Else EndIF" - ������� ������ ����� ����������
'������ � ������ � ����������� �� ���������� ���������. ���� ������
'�������� ������ � ������� �� 1 (�������) - ���� �� ����������� "IF
' ElseIF Else EndIF" ������� �� ���� ���� � ��������� �����������
'� ��� �� ������, � ������� ��������������� ������������ ����������
'��������� (������) ��������.

'�������� msg_flag (������ ��������) �������� ����� �� ����� ����������
'��������� � ����������� �� ������, ���� ��� �������� (���� ������
'������� ����������� ������ ������� ��������� �� ��������).
'������ ������ ����� � ��� �������� ����� ��� ���������; �� �����
'����� � ����, ��� ���� �� ����� ����� ����� � ���� �������� � ��
'�����-�� ������� ��������� ������ � ������ (������� ������ � ���������
'������� � �.�.) �� �������� ����� ����������� ����������� ����
'��������� ��� ������! ����������� ����� ��� �����.
'�� ��������� ������� ������������ ��� ��������� �� ������� ����, ����
'��� ��������� (�.�. ���� �������� "msg_string" �� ��������� ��������).
'������� �� ��������� ������ ���������� ��������� ���������
'��������������� ������ � ������ ����������� ������.

'����������� ������.
'������� ������������ ��������, ����� ������, �� ������� ���������
'�������, �������� �������� ������ ��������������� � ������� ����,
'��� ����������� �������� (True, False) ��� �������� �������� �
'������������� ���������, � ����� ���� ������ �������� ��������� ������.
'� ����� �������, ����� ������, �� ������� ��������� �������,
'������������� ��� ����, ������ ��� ���� ��� �������� �������������
'��������, ��������� �������� - ������ �������� ������: "#����!",
'"#�����!", "#���?", "#���/0!".
'����� � ������ ����� � ����� ������� ����������
'���������-�������������� "� �������� ������ ���������� � �������".

'���������� (��������).
'����� ���������� ������ ��������� � ���� ������� �� ������������
'��������� - �.�. � ����������� �� ����� ���������� ��������� - ���
'����� �����, �� ������� ����� ����������� 30000 ���. ������� �������
'� ������� ������ �� �������������� ����� ��������, ����� �����
'���������� ������ ������� �� ���� �����.


    Dim custom_sum_ru As Variant
    Dim bool_1 As Boolean, _
        bool_2 As Boolean
    Dim msg_string As String

    On Error Resume Next

    If TypeName(custom_sum) = "Range" Then
        bool_1 = TypeName(custom_sum.Value) = "Boolean"
        bool_2 = TypeName(custom_sum.Value) = "Error"
    Else
        bool_1 = TypeName(custom_sum) = "Boolean"
        bool_2 = TypeName(custom_sum) = "Error"
    End If

    custom_sum_ru = custom_sum * currency_rate
    
    If IsDate(custom_sum) Then
        msg_string = "������� ������� �������� ��� ������; " _
                      & vbCrLf & _
                     "�������� ��������� �� ����!"
        custom_toll = CVErr(xlErrValue)
    ElseIf bool_1 Then
        msg_string = "������� ������� �������� ��� ������; " _
                      & vbCrLf & _
                     "�������� ��������� �� ���������� ��������!"
        custom_toll = CVErr(xlErrValue)
    ElseIf Application.WorksheetFunction.IsText(custom_sum) Then
        msg_string = "������� ������� �������� ��� ������; " _
                      & vbCrLf & _
                     "�������� ��������� �� ��������� �������� (�����)!"
        custom_toll = CVErr(xlErrValue)
    ElseIf bool_2 Then
        msg_string = "������� �������� ����������� ��� ��������� ���:" _
                      & vbCrLf & _
                     "�������� ����������� �������� ��� ������ �� ������;" _
                      & vbCrLf & _
                      "������ ����������. ��������� ��������� ������."
        custom_toll = CVErr(xlErrName)
    ElseIf custom_sum < 0 Or currency_rate < 0 Then
        msg_string = "���������� ��������� ��� ���� ������ �� ����� " & _
                     "���� ������������� ������. " & vbCrLf & _
                     "��������� ���������� ������� ���������."
        custom_toll = CVErr(xlErrNum)
    ElseIf currency_rate = 0 Then
        msg_string = "� ������� ����������� ������� ������� �� ����." _
                      & vbCrLf & _
                      "��������� ��������� � ������ ���������� � �������!"
        custom_toll = CVErr(xlErrDiv0)
    ElseIf custom_sum_ru >= 0 And _
           IsNumeric(custom_sum_ru) And _
           custom_sum_ru <> "" Then
            If custom_sum_ru >= 0 And custom_sum_ru <= 200000 Then
                custom_toll = 775 / currency_rate
            ElseIf custom_sum_ru > 200000 And custom_sum_ru <= 450000 Then
                custom_toll = 1550 / currency_rate
            ElseIf custom_sum_ru > 450000 And custom_sum_ru <= 1200000 Then
                custom_toll = 3100 / currency_rate
            ElseIf custom_sum_ru > 1200000 And custom_sum_ru <= 2700000 Then
                custom_toll = 8530 / currency_rate
            ElseIf custom_sum_ru > 2700000 And custom_sum_ru <= 4200000 Then
                custom_toll = 12000 / currency_rate
            ElseIf custom_sum_ru > 4200000 And custom_sum_ru <= 5500000 Then
                custom_toll = 15500 / currency_rate
            ElseIf custom_sum_ru > 5500000 And custom_sum_ru <= 7000000 Then
                custom_toll = 20000 / currency_rate
            ElseIf custom_sum_ru > 7000000 And custom_sum_ru <= 8000000 Then
                custom_toll = 23000 / currency_rate
            ElseIf custom_sum_ru > 8000000 And custom_sum_ru <= 9000000 Then
                custom_toll = 25000 / currency_rate
            ElseIf custom_sum_ru > 9000000 And custom_sum_ru <= 10000000 Then
                custom_toll = 27000 / currency_rate
            ElseIf custom_sum_ru > 10000000 Then
                custom_toll = 30000 / currency_rate
            End If
    Else
        msg_string = "�������� ��� ������. " _
                      & vbCrLf & _
                     "������� ������� ������������ ��������. " _
                      & vbCrLf & _
                     "��������� ������, �� ������� ��������� �������!"
        custom_toll = CVErr(xlErrValue)
    End If

    If TypeName(custom_toll) <> "Error" Then _
        custom_toll = Round(custom_toll, 2)
    
    If msg_string <> "" And msg_flag Then MsgBox msg_string

    On Error GoTo 0

End Function
