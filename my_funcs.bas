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

Function CUSTOM_TOLL( _
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
        CUSTOM_TOLL = CVErr(xlErrValue)
    ElseIf bool_1 Then
        msg_string = "������� ������� �������� ��� ������; " _
                      & vbCrLf & _
                     "�������� ��������� �� ���������� ��������!"
        CUSTOM_TOLL = CVErr(xlErrValue)
    ElseIf Application.WorksheetFunction.IsText(custom_sum) Then
        msg_string = "������� ������� �������� ��� ������; " _
                      & vbCrLf & _
                     "�������� ��������� �� ��������� �������� (�����)!"
        CUSTOM_TOLL = CVErr(xlErrValue)
    ElseIf bool_2 Then
        msg_string = "������� �������� ����������� ��� ��������� ���:" _
                      & vbCrLf & _
                     "�������� ����������� �������� ��� ������ �� ������;" _
                      & vbCrLf & _
                      "������ ����������. ��������� ��������� ������."
        CUSTOM_TOLL = CVErr(xlErrName)
    ElseIf custom_sum < 0 Or currency_rate < 0 Then
        msg_string = "���������� ��������� ��� ���� ������ �� ����� " & _
                     "���� ������������� ������. " & vbCrLf & _
                     "��������� ���������� ������� ���������."
        CUSTOM_TOLL = CVErr(xlErrNum)
    ElseIf currency_rate = 0 Then
        msg_string = "� ������� ����������� ������� ������� �� ����." _
                      & vbCrLf & _
                      "��������� ��������� � ������ ���������� � �������!"
        CUSTOM_TOLL = CVErr(xlErrDiv0)
    ElseIf custom_sum_ru >= 0 And _
           IsNumeric(custom_sum_ru) And _
           custom_sum_ru <> "" Then
            If custom_sum_ru >= 0 And custom_sum_ru <= 200000 Then
                CUSTOM_TOLL = 775 / currency_rate
            ElseIf custom_sum_ru > 200000 And custom_sum_ru <= 450000 Then
                CUSTOM_TOLL = 1550 / currency_rate
            ElseIf custom_sum_ru > 450000 And custom_sum_ru <= 1200000 Then
                CUSTOM_TOLL = 3100 / currency_rate
            ElseIf custom_sum_ru > 1200000 And custom_sum_ru <= 2700000 Then
                CUSTOM_TOLL = 8530 / currency_rate
            ElseIf custom_sum_ru > 2700000 And custom_sum_ru <= 4200000 Then
                CUSTOM_TOLL = 12000 / currency_rate
            ElseIf custom_sum_ru > 4200000 And custom_sum_ru <= 5500000 Then
                CUSTOM_TOLL = 15500 / currency_rate
            ElseIf custom_sum_ru > 5500000 And custom_sum_ru <= 7000000 Then
                CUSTOM_TOLL = 20000 / currency_rate
            ElseIf custom_sum_ru > 7000000 And custom_sum_ru <= 8000000 Then
                CUSTOM_TOLL = 23000 / currency_rate
            ElseIf custom_sum_ru > 8000000 And custom_sum_ru <= 9000000 Then
                CUSTOM_TOLL = 25000 / currency_rate
            ElseIf custom_sum_ru > 9000000 And custom_sum_ru <= 10000000 Then
                CUSTOM_TOLL = 27000 / currency_rate
            ElseIf custom_sum_ru > 10000000 Then
                CUSTOM_TOLL = 30000 / currency_rate
            End If
    Else
        msg_string = "�������� ��� ������. " _
                      & vbCrLf & _
                     "������� ������� ������������ ��������. " _
                      & vbCrLf & _
                     "��������� ������, �� ������� ��������� �������!"
        CUSTOM_TOLL = CVErr(xlErrValue)
    End If

    If TypeName(CUSTOM_TOLL) <> "Error" Then _
        CUSTOM_TOLL = Round(CUSTOM_TOLL, 2)
    
    If msg_string <> "" And msg_flag Then MsgBox msg_string

    On Error GoTo 0

End Function

Public Function CYR2LATIN(text_string As String) As String
    
'����������� ������� �������.
'������ ������ ������������� �������� �� ��������.
'�������� ��������������.

'������� ������������ ����� � ��������� ������ (����������
'���������� ��� �� ������) ������������� ��������, �������
'�� ��������� ����� ������ �� ��������� ����� � �������� ��
'���������� ���������.

'��������������� ��� ������� ����������� � ������� ���������� ��������.
'Sic! ������� (�������) �������� �� ��� ������������� �������!
'�������� ������ �� ������������� �������, ������� ��������� ������
'� ����������.

    Dim latin As Variant
    Dim cyril As Variant
    Dim i As Double
    Dim j As Double
    Dim one_symbol As String
    Dim find_flag As Boolean
    Dim symbol_translit As String
    Dim merged_text As String

'    latin = Array("a", "B", "c", "e", "k", "m", "n", "H", _
'                  "o", "p", "T", "u", "y", "A", "B", "E", _
'                  "K", "M", "O", "P", "C", "T", "H", "Y")
'    cyril = Array("�", "�", "�", "�", "�", "�", "�", "�", _
'                  "�", "�", "�", "�", "�", "�", "�", "�", _
'                  "�", "�", "�", "�", "�", "�", "�", "�")

    '������� "�" �� ��� ������ �� ��������� "N" ��� "U"?
    '������ ��������. ���� � ������ ��������.
    cyril = Array("�", "�", "�", "�", "�", "�", "�", _
                  "�", "�", "�", "�", "�", "�", "�", _
                  "�", "�", "�", "�", "�", "�", "�", _
                  "�", "�", "�", "�", "�", "�", "�")
    latin = Array("a", "B", "c", "e", "k", "m", "H", _
                  "o", "n", "p", "T", "u", "y", "x", _
                  "A", "B", "C", "E", "K", "M", "H", _
                  "O", "n", "P", "T", "U", "Y", "X")

    For i = 1 To Len(text_string)
        one_symbol = Mid(text_string, i, 1)
        find_flag = False
        For j = LBound(cyril) To UBound(cyril)
            If cyril(j) = one_symbol Then
                symbol_translit = latin(j)
                find_flag = True
                Exit For
            End If
        Next
        If find_flag Then
            merged_text = merged_text & symbol_translit
        Else: merged_text = merged_text & one_symbol
        End If
    Next

    CYR2LATIN = merged_text

End Function

Public Function REPLACE_CYRIL_LATIN(txt_ref) As String
    
'������� ������ ������������� �������� (������� �� ���������)
'� ���������, ������ �� �� ��������� ������� � ���������� ������
'������ �� ������� (������������� �������);

    Dim latin As Variant
    Dim cyril As Variant
    Dim i As Double
    Dim j As Double
    Dim one_symbol As String
    Dim find_flag As Boolean
    Dim symbol_translit As String
    Dim merged_text As String

'    latin = Array("a", "B", "c", "e", "k", "m", "n", "H", _
'                  "o", "p", "T", "u", "y", "A", "B", "E", _
'                  "K", "M", "O", "P", "C", "T", "H", "Y")
'    cyril = Array("�", "�", "�", "�", "�", "�", "�", "�", _
'                  "�", "�", "�", "�", "�", "�", "�", "�", _
'                  "�", "�", "�", "�", "�", "�", "�", "�")

    '������� "�" �� ��� ������ �� ��������� "N" ��� "U"?
    '������ ��������. ���� � ������ ��������.
    cyril = Array("�", "�", "�", "�", "�", "�", "�", _
                  "�", "�", "�", "�", "�", "�", "�", _
                  "�", "�", "�", "�", "�", "�", "�", _
                  "�", "�", "�", "�", "�", "�", "�")
    latin = Array("a", "B", "c", "e", "k", "m", "H", _
                  "o", "n", "p", "T", "u", "y", "x", _
                  "A", "B", "C", "E", "K", "M", "H", _
                  "O", "n", "P", "T", "U", "Y", "X")

    If TypeName(txt_ref) = "Range" Then
        text_string = txt_ref.Text
    Else
        text_string = txt_ref
    End If

    For i = 1 To Len(text_string)
        one_symbol = Mid(text_string, i, 1)
        find_flag = False
        For j = LBound(cyril) To UBound(cyril)
            If cyril(j) = one_symbol Then
                symbol_translit = latin(j)
                find_flag = True
                Exit For
            End If
        Next
        If find_flag Then
            merged_text = merged_text & symbol_translit
            If TypeName(txt_ref) = "Range" Then _
                txt_ref.Characters(i, 1).Font.ColorIndex = 7
        Else: merged_text = merged_text & one_symbol
        End If
    Next

    REPLACE_CYRIL_LATIN = merged_text
    
End Function
